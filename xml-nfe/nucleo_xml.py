"""Regras das ferramentas de XML de NF-e, separadas da tela para poderem ser testadas.

Duas operações, as duas sempre em dois tempos: primeiro monta o plano (nada
muda no disco), depois executa o que estiver no plano.
"""
from __future__ import annotations

import os
import re
import shutil
from dataclasses import dataclass
from pathlib import Path

from lxml import etree

NS = {"nfe": "http://www.portalfiscal.inf.br/nfe"}


@dataclass
class Movimento:
    """Um arquivo que o plano de separação pretende mover ou copiar."""
    arquivo: Path
    numero: str
    destino: Path
    situacao: str  # "ok" | "já existe" | "erro"
    detalhe: str = ""


@dataclass
class Alteracao:
    """Um item de nota cujo CFOP o plano pretende trocar."""
    arquivo: Path
    nota: str
    codigo: str
    cfop_atual: str
    cfop_novo: str


def _mesmo_arquivo(a: Path, b: Path) -> bool:
    """Se são o mesmo arquivo no disco, mesmo escrito de jeitos diferentes (unidade mapeada x caminho UNC)."""
    try:
        if a.exists() and b.exists():
            return a.samefile(b)
    except OSError:
        pass
    return os.path.normcase(os.path.abspath(a)) == os.path.normcase(os.path.abspath(b))


def mesma_pasta(a: Path, b: Path) -> bool:
    """Mesma comparação, para a tela conferir origem x destino antes de executar."""
    return _mesmo_arquivo(a, b)


def listar_xmls(pasta: Path, incluir_subpastas: bool) -> list[Path]:
    padrao = "**/*.xml" if incluir_subpastas else "*.xml"
    return sorted(p for p in pasta.glob(padrao) if p.is_file())


def separar_valores(texto: str) -> list[str]:
    """Quebra o que a pessoa digitou (vírgula, ponto e vírgula, espaço ou linha) numa lista."""
    partes = [p.strip() for p in re.split(r"[,;\s]+", texto or "")]
    vistos, saida = set(), []
    for p in partes:
        if p and p not in vistos:
            vistos.add(p)
            saida.append(p)
    return saida


def _sem_namespace(tag) -> str:
    """Nome da tag sem o namespace: '{...}nNF' vira 'nNF'."""
    return tag.rsplit("}", 1)[-1] if isinstance(tag, str) else ""


def _buscar(elemento, nome: str) -> list:
    """Todos os descendentes com esse nome, com namespace ou sem.

    XML de NF-e normalmente usa o namespace do portalfiscal, mas aparece arquivo salvo
    de outro jeito, sem namespace ou com outro. Procurar pelo nome da tag faz a
    ferramenta funcionar nos dois casos em vez de devolver "nada encontrado".
    """
    return [e for e in elemento.iter() if _sem_namespace(e.tag) == nome]


def _primeiro(elemento, nome: str):
    achados = _buscar(elemento, nome)
    return achados[0] if achados else None


def _texto(elemento) -> str:
    return (elemento.text or "").strip() if elemento is not None else ""


def notas_do_xml(raiz) -> list:
    """[(elemento da NFe, número da nota)] — um arquivo pode trazer mais de uma NFe."""
    notas = []
    for inf in _buscar(raiz, "infNFe"):
        ide = _primeiro(inf, "ide")
        notas.append((inf, _texto(_primeiro(ide, "nNF")) if ide is not None else ""))
    if not notas:                      # XML sem infNFe: trata o arquivo inteiro como uma nota
        notas.append((raiz, _texto(_primeiro(raiz, "nNF"))))
    return notas


def ler_nota(caminho: Path) -> str:
    """Número da nota (nNF) de dentro do XML. Devolve "" se não der para ler."""
    try:
        raiz = etree.parse(str(caminho)).getroot()
    except Exception:
        return ""
    notas = notas_do_xml(raiz)
    return notas[0][1] if notas else ""


# ----------------------------------------------------------------- separar por número
def planejar_separacao(origem: Path, destino: Path, numeros: list[str], incluir_subpastas: bool,
                       procurar_dentro: bool) -> tuple[list[Movimento], list[str]]:
    """Monta a lista de arquivos a mover. Devolve (movimentos, números sem nenhum arquivo)."""
    arquivos = listar_xmls(origem, incluir_subpastas)
    movimentos: list[Movimento] = []
    sem_nada: list[str] = []
    ja_escolhidos: set[Path] = set()
    nomes_no_destino: dict[Path, Path] = {}   # destino -> arquivo que já vai ocupar esse nome

    notas_por_arquivo: dict[Path, str] = {}
    if procurar_dentro:
        notas_por_arquivo = {a: ler_nota(a) for a in arquivos}

    for numero in numeros:
        encontrados = [a for a in arquivos
                       if a.name.startswith(f"{numero}_") and a not in ja_escolhidos]
        if procurar_dentro:
            encontrados += [a for a in arquivos
                            if notas_por_arquivo.get(a) == numero
                            and a not in ja_escolhidos and a not in encontrados]
        if not encontrados:
            sem_nada.append(numero)
            continue
        for arquivo in sorted(encontrados):
            ja_escolhidos.add(arquivo)
            alvo = destino / arquivo.name
            if alvo.exists():
                situacao, detalhe = "já existe", "já existe um arquivo com esse nome no destino"
            elif alvo in nomes_no_destino:
                # duas subpastas com XML de mesmo nome cairiam no mesmo lugar e um
                # apagaria o outro em silêncio
                situacao = "nome repetido"
                detalhe = (f"outro arquivo desta busca ({nomes_no_destino[alvo].name}, em "
                           f"{nomes_no_destino[alvo].parent.name}) tem o mesmo nome e iria "
                           f"para o mesmo lugar")
            else:
                situacao, detalhe = "ok", ""
                nomes_no_destino[alvo] = arquivo
            movimentos.append(Movimento(arquivo=arquivo, numero=numero, destino=alvo,
                                        situacao=situacao, detalhe=detalhe))
    return movimentos, sem_nada


def executar_separacao(movimentos: list[Movimento], mover: bool) -> tuple[int, list[str]]:
    """Move (ou copia) o que está com situação "ok". Devolve (quantos, erros)."""
    feitos, erros = 0, []
    for m in movimentos:
        if m.situacao != "ok":
            continue
        try:
            m.destino.parent.mkdir(parents=True, exist_ok=True)
            if mover:
                shutil.move(str(m.arquivo), str(m.destino))
            else:
                shutil.copy2(str(m.arquivo), str(m.destino))
            feitos += 1
        except Exception as erro:
            erros.append(f"{m.arquivo.name}: {erro}")
    return feitos, erros


# ----------------------------------------------------------------- trocar CFOP
def planejar_cfop(pasta: Path, notas: list[str], codigos: list[str], cfop_novo: str,
                  incluir_subpastas: bool) -> tuple[list[Alteracao], list[str]]:
    """Lista os itens que teriam o CFOP trocado. Notas ou códigos vazios querem dizer "todos"."""
    alteracoes: list[Alteracao] = []
    erros: list[str] = []
    for caminho in listar_xmls(pasta, incluir_subpastas):
        try:
            raiz = etree.parse(str(caminho)).getroot()
        except Exception as erro:
            erros.append(f"{caminho.name}: não consegui abrir ({erro})")
            continue
        # um arquivo pode trazer mais de uma NFe; cada uma tem o seu número e os seus itens
        for elemento, nota in notas_do_xml(raiz):
            if notas and nota not in notas:
                continue
            for det in _buscar(elemento, "det"):
                prod = _primeiro(det, "prod")
                if prod is None:
                    continue
                c_prod, cfop = _primeiro(prod, "cProd"), _primeiro(prod, "CFOP")
                if c_prod is None or cfop is None:
                    continue
                codigo, atual = _texto(c_prod), _texto(cfop)
                if codigos and codigo not in codigos:
                    continue
                if atual == cfop_novo:
                    continue
                alteracoes.append(Alteracao(arquivo=caminho, nota=nota, codigo=codigo,
                                            cfop_atual=atual, cfop_novo=cfop_novo))
    return alteracoes, erros


def executar_cfop(alteracoes: list[Alteracao], saida: Path,
                  substituir: bool) -> tuple[int, int, list[str]]:
    """Grava uma cópia alterada de cada XML na pasta de saída.

    Devolve (arquivos gravados, itens trocados, erros). O XML de origem não é tocado.
    """
    erros: list[str] = []
    gravados = trocados = 0
    por_arquivo: dict[Path, list[Alteracao]] = {}
    for a in alteracoes:
        por_arquivo.setdefault(a.arquivo, []).append(a)

    for caminho, itens in por_arquivo.items():
        destino = saida / caminho.name
        if destino.exists() and not substituir:
            erros.append(f"{caminho.name}: já existe no destino, não mexi")
            continue
        if _mesmo_arquivo(destino, caminho):
            erros.append(f"{caminho.name}: a saída é o próprio arquivo de origem, não mexi")
            continue
        try:
            arvore = etree.parse(str(caminho))
            raiz = arvore.getroot()
            # a chave inclui a NOTA: num arquivo com várias NFe, trocar por código e CFOP
            # mexeria em itens de notas que o usuário não pediu
            alvo = {(a.nota, a.codigo, a.cfop_atual): a.cfop_novo for a in itens}
            feitos_aqui = 0
            for elemento, nota in notas_do_xml(raiz):
                for det in _buscar(elemento, "det"):
                    prod = _primeiro(det, "prod")
                    if prod is None:
                        continue
                    c_prod, cfop = _primeiro(prod, "cProd"), _primeiro(prod, "CFOP")
                    if c_prod is None or cfop is None:
                        continue
                    chave = (nota, _texto(c_prod), _texto(cfop))
                    if chave in alvo:
                        cfop.text = alvo[chave]
                        feitos_aqui += 1
            destino.parent.mkdir(parents=True, exist_ok=True)
            arvore.write(str(destino), encoding="utf-8", xml_declaration=True)
            gravados += 1
            trocados += feitos_aqui
        except Exception as erro:
            erros.append(f"{caminho.name}: {erro}")
    return gravados, trocados, erros
