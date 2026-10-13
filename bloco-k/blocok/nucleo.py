"""Estruturas e regras comuns a todas as empresas.

Os leitores (pasta leitores/) transformam a planilha de estoque de cada empresa numa
lista de ItemLido. Daqui para frente o caminho é o mesmo para todo mundo: consolidar
por código, conferir e gerar as duas planilhas de importação (blocok/saidas.py).

Regra do "Saldo" no modelo de importação: é o VALOR TOTAL do item em estoque. Quando o
mesmo código aparece em várias linhas, as quantidades e os totais são somados.
"""
from __future__ import annotations

import calendar
import datetime as dt
import re
import unicodedata
from dataclasses import dataclass, field
from pathlib import Path

# O sistema contábil guarda só os 14 últimos caracteres do código do produto na importação.
LIMITE_CODIGO_DOMINIO = 14
# Quando a planilha não traz NCM, o cadastro sai "zerado" para não travar o processo.
NCM_SEM_INFORMACAO = "00000000"

MESES_NOMES = {
    "JANEIRO": 1, "JAN": 1, "FEVEREIRO": 2, "FEV": 2, "MARCO": 3, "MAR": 3, "ABRIL": 4, "ABR": 4,
    "MAIO": 5, "MAI": 5, "JUNHO": 6, "JUN": 6, "JULHO": 7, "JUL": 7, "AGOSTO": 8, "AGO": 8,
    "SETEMBRO": 9, "SET": 9, "OUTUBRO": 10, "OUT": 10, "NOVEMBRO": 11, "NOV": 11, "DEZEMBRO": 12, "DEZ": 12,
}


class ErroLeitura(Exception):
    """Problema na planilha de estoque, com mensagem pronta para mostrar na tela."""


@dataclass
class ItemLido:
    """Uma linha da planilha de estoque.

    Informe valor_total quando a planilha tiver o custo total da linha; sem ele,
    o total é quantidade × valor_unitario.
    """
    codigo: str
    descricao: str
    ncm: str
    unidade: str
    quantidade: float
    valor_unitario: float | None = None
    valor_total: float | None = None
    linha: int | None = None          # linha na planilha de origem, para as mensagens
    indicador: int | None = None      # 0, 1 ou 2; None = padrão do leitor
    participante: str = ""            # "C" (cliente) ou "F" (fornecedor), só quando indicador 1 ou 2
    cnpj_participante: str = ""

    @property
    def total(self) -> float:
        if self.valor_total is not None:
            return self.valor_total
        return self.quantidade * (self.valor_unitario or 0.0)

    @property
    def unitario(self) -> float:
        if self.valor_unitario is not None:
            return self.valor_unitario
        return self.total / self.quantidade if self.quantidade else 0.0


@dataclass
class Leitura:
    """O que o leitor devolve para um arquivo.

    conta: conta contábil deste arquivo, quando ela vem na planilha (senão vale a do leitor).
    sufixo: vai no nome dos arquivos de saída, para separar importações da mesma empresa
            (ex.: " - FILIAL" quando a empresa manda um arquivo por filial).
    """
    itens: list[ItemLido]
    aba: str = ""
    competencia: dt.date | None = None
    avisos: list[str] = field(default_factory=list)
    conta: str | None = None
    sufixo: str = ""


@dataclass
class Aviso:
    """Algo para olhar antes de importar. Não impede a geração."""
    tipo: str      # Leitura, Dados, Divergência, Código longo
    codigo: str
    texto: str


@dataclass
class ItemConsolidado:
    """Um código no modelo de importação: quantidades e totais somados."""
    codigo: str
    descricao: str
    ncm: str
    unidade: str
    quantidade: float
    total: float                      # vai na coluna "Saldo"
    indicador: int | None
    participante: str
    cnpj_participante: str
    linhas: list[int]

    @property
    def valor_medio(self) -> float:
        return self.total / self.quantidade if self.quantidade else 0.0


# ---------------------------------------------------------------- conversões

def normalizar(valor) -> str:
    """Texto em maiúsculas, sem acento, sem pontuação nas pontas e com espaços simples."""
    if valor is None:
        return ""
    s = unicodedata.normalize("NFKD", str(valor))
    s = "".join(c for c in s if not unicodedata.combining(c)).upper()
    s = re.sub(r"\s+", " ", s)
    return s.strip(" :.;-_\t\r\n")


def texto(valor) -> str:
    """Conteúdo de célula como texto, sem espaços nas pontas. 10357.0 vira '10357'.

    Espaços no meio ficam como estão, para a descrição sair igual à planilha de origem.
    """
    if valor is None:
        return ""
    if isinstance(valor, float) and valor.is_integer():
        valor = int(valor)
    return re.sub(r"[\r\n\t]+", " ", str(valor)).strip()


def numero(valor) -> float | None:
    """Número de uma célula; aceita '1.234,56', 'R$ 12,50', '12.5'. None se não der.

    Texto só com pontos no padrão de milhar ('1.500') é lido como milhar, como no Brasil.
    """
    if valor is None or isinstance(valor, bool):
        return None
    if isinstance(valor, (int, float)):
        return float(valor)
    s = str(valor).strip().replace("R$", "").replace(" ", "").replace(" ", "")
    if not s or s in ("-", "--"):
        return None
    negativo = s.startswith("(") and s.endswith(")")
    s = s.strip("()")
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    elif re.fullmatch(r"-?\d{1,3}(\.\d{3})+", s):
        s = s.replace(".", "")
    try:
        n = float(s)
    except ValueError:
        return None
    return -n if negativo else n


def ncm(valor) -> str:
    """NCM só com dígitos. Se o Excel comeu o zero da frente (7 dígitos), ele volta."""
    s = re.sub(r"\D", "", texto(valor))
    if len(s) == 7:
        s = "0" + s
    return s


def ultimo_dia(ano: int, mes: int) -> dt.date:
    return dt.date(ano, mes, calendar.monthrange(ano, mes)[1])


def competencia_do_texto(txt) -> dt.date | None:
    """Último dia do mês citado no texto ou na data.

    'Estoque 30-06-2026' -> 30/06/2026; 'EST 08 2026', '072026', '07-2026', '2026-07' -> fim do mês.
    Usa só mês e ano, então abas como 'Estoque 31-04-2021' também funcionam (30/04/2021).
    """
    if isinstance(txt, (dt.date, dt.datetime)):
        return ultimo_dia(txt.year, txt.month)
    txt = str(txt or "")
    m = re.search(r"(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})", txt)
    if m:
        mes, ano = int(m.group(2)), int(m.group(3))
        return ultimo_dia(ano, mes) if 1 <= mes <= 12 else None
    m = re.search(r"(?<!\d)(?:(0[1-9]|1[0-2])[-_./ ]?|([1-9])[-_./ ])(20\d{2})(?!\d)", txt)
    if m:
        return ultimo_dia(int(m.group(3)), int(m.group(1) or m.group(2)))
    m = re.search(r"(?<!\d)(20\d{2})[-_./ ]?(0[1-9]|1[0-2])(?!\d)", txt)
    if m:
        return ultimo_dia(int(m.group(1)), int(m.group(2)))
    return None


def competencia_das_pastas(caminho, niveis: int = 3) -> dt.date | None:
    """Mês nas pastas que guardam o arquivo, da mais próxima para cima: '...\\2026\\07-2026\\arquivo' -> 31/07/2026."""
    for pasta in list(Path(caminho).parents)[:niveis]:
        data = competencia_do_texto(pasta.name)
        if data:
            return data
    return None


def competencia_do_nome(caminho) -> dt.date | None:
    """Mês no nome do arquivo: '...28-02-2026', 'BLOCO K 08 2026' ou 'ESTOQUE - JULHO' (ano tirado do caminho)."""
    p = Path(caminho)
    data = competencia_do_texto(p.stem)
    if data:
        return data
    anos = re.findall(r"(?<!\d)(20\d{2})(?!\d)", str(p))
    if anos:
        for palavra in re.findall(r"[A-Z]+", normalizar(p.stem)):
            if palavra in MESES_NOMES:
                return ultimo_dia(int(anos[-1]), MESES_NOMES[palavra])
    return None


def competencia_do_caminho(caminho) -> dt.date | None:
    """Competência pela pasta do mês (mais confiável) ou, sem ela, pelo nome do arquivo."""
    return competencia_das_pastas(caminho) or competencia_do_nome(caminho)


def escolher_competencia(fontes: list[tuple[str, dt.date | None]]) -> tuple[dt.date | None, str | None]:
    """Primeira fonte que tiver data, na ordem de preferência; avisa se as outras discordarem.

    fontes: [("pasta", data), ("nome do arquivo", data), ...]. Devolve (competência, aviso ou None).
    """
    achadas = [(origem, data) for origem, data in fontes if data]
    if not achadas:
        return None, None
    origem, escolhida = achadas[0]
    outras = [f"{o} {d:%m/%Y}" for o, d in achadas[1:] if (d.year, d.month) != (escolhida.year, escolhida.month)]
    if not outras:
        return escolhida, None
    return escolhida, (f"Competência {escolhida:%m/%Y} ({origem}), mas outras fontes indicam outro mês: "
                       + "; ".join(outras) + ". Confira.")


def _alvos(titulo) -> list[str]:
    return [normalizar(t) for t in (titulo if isinstance(titulo, (list, tuple)) else [titulo])]


def achar_cabecalho(linhas: list[list], titulos: dict, opcionais: dict | None = None,
                    max_linhas: int = 15) -> tuple[int, dict[str, int]]:
    """Acha a linha de títulos e a coluna de cada campo.

    titulos: {"codigo": "CÓDIGO", "quantidade": ("Quantidade Final", "Qtde Final"), ...}
        Cada campo aceita um título ou uma lista de títulos alternativos.
    opcionais: mesmo formato; entram só se existirem na linha de títulos.
    A comparação ignora acento, maiúscula, espaços nas pontas e ':' no fim.
    Devolve (índice da linha de títulos, {campo: índice da coluna}), com índices a partir de 0.
    """
    obrigatorios = {campo: _alvos(t) for campo, t in titulos.items()}
    extras = {campo: _alvos(t) for campo, t in (opcionais or {}).items()}
    for i, linha in enumerate(linhas[:max_linhas]):
        valores = [normalizar(v) for v in linha]
        achados: dict[str, int] = {}
        for campo, alvos in list(obrigatorios.items()) + list(extras.items()):
            for alvo in alvos:
                if alvo in valores:
                    achados[campo] = valores.index(alvo)
                    break
        if all(campo in achados for campo in obrigatorios):
            return i, achados
    faltam = ", ".join("'" + (t if isinstance(t, str) else "' ou '".join(t)) + "'" for t in titulos.values())
    raise ErroLeitura(f"Não encontrei a linha de títulos ({faltam}) nas primeiras {max_linhas} linhas. "
                      "O layout da planilha mudou?")


def achar_aba(abas: list, titulos: dict, opcionais: dict | None = None):
    """Primeira aba que tem a linha de títulos. Devolve (aba, linha dos títulos, colunas)."""
    for aba in abas:
        try:
            linha, colunas = achar_cabecalho(aba.linhas, titulos, opcionais)
            return aba, linha, colunas
        except ErroLeitura:
            continue
    nomes = ", ".join(f"'{a.nome}'" for a in abas)
    raise ErroLeitura(f"Nenhuma aba tem os títulos esperados ({', '.join(map(str, titulos.values()))}). "
                      f"Abas do arquivo: {nomes}. O layout da planilha mudou?")


def celula(linha: list, coluna: int | None):
    if coluna is None:
        return None
    return linha[coluna] if coluna < len(linha) else None


# ---------------------------------------------------------------- consolidação

def consolidar(itens: list[ItemLido], indicador_padrao: int = 0) -> list[ItemConsolidado]:
    """Soma quantidades e totais por código (e posse), na ordem em que o código aparece primeiro."""
    grupos: dict[tuple, dict] = {}
    for it in itens:
        indicador = indicador_padrao if it.indicador is None else it.indicador
        chave = (it.codigo, indicador, it.participante, it.cnpj_participante)
        g = grupos.get(chave)
        if g is None:
            g = grupos[chave] = {"primeiro": it, "indicador": indicador, "qtd": 0.0, "total": 0.0, "linhas": []}
        g["qtd"] += it.quantidade
        g["total"] += it.total
        g["linhas"].append(it.linha)

    saida = []
    for (codigo, _, participante, cnpj), g in grupos.items():
        p = g["primeiro"]
        saida.append(ItemConsolidado(
            codigo=codigo, descricao=p.descricao, ncm=p.ncm, unidade=p.unidade,
            quantidade=round(g["qtd"], 6), total=round(g["total"], 2),
            indicador=g["indicador"], participante=participante, cnpj_participante=cnpj,
            linhas=g["linhas"],
        ))
    return saida


VALOR_UNITARIO_SUSPEITO = 1_000_000.0   # preço unitário acima disso é outra coluna lida como preço


def _conferir_layout(itens: list[ItemLido], conferir_valores: bool) -> list[Aviso]:
    """Avisa quando a planilha mudou a ordem das colunas mantendo os títulos.

    Em um layout real a empresa trocou a ordem das colunas e deixou os títulos no lugar:
    o número da ordem de serviço entrou como quantidade e o NCM como preço unitário, e a
    Posse saiu na casa dos trilhões sem nenhum erro. Item a item isso vira dezenas de
    avisos que passam batido; aqui vira um aviso só, no topo, dizendo o que aconteceu.
    """
    if not itens:
        return []
    avisos = []

    sem_ncm = sum(1 for it in itens if len(it.ncm) != 8)
    if sem_ncm >= len(itens) * 0.8 and sem_ncm > 3:
        avisos.append(Aviso("Coluna trocada", "",
                            f"{sem_ncm} de {len(itens)} itens estão sem NCM de 8 dígitos. Quando é a "
                            f"planilha inteira, quase sempre quer dizer que a planilha mudou a ordem "
                            f"das colunas e o que foi lido não é o que o título promete. Confira a "
                            f"aba 'Linhas lidas' antes de gerar."))

    if conferir_valores:
        caros = [it for it in itens if (it.valor_unitario or 0) >= VALOR_UNITARIO_SUSPEITO]
        if caros:
            exemplo = caros[0]
            avisos.append(Aviso("Coluna trocada", exemplo.codigo,
                                f"{len(caros)} item(ns) com valor unitário acima de "
                                f"{moeda(VALOR_UNITARIO_SUSPEITO)} (ex.: linha {exemplo.linha}, "
                                f"{moeda(exemplo.valor_unitario or 0)}). Preço desse tamanho costuma ser "
                                f"outra coluna — NCM, código ou número de OS — lida como preço."))

        datas = [it for it in itens if it.linha and it.valor_total is None and it.valor_unitario]
        if len(datas) >= len(itens) * 0.8 and len(datas) > 3:
            avisos.append(Aviso("Coluna trocada", "",
                                f"A coluna de valor total não trouxe número em {len(datas)} de "
                                f"{len(itens)} linhas; o Saldo saiu de quantidade × unitário. "
                                f"Confira se a coluna certa está sendo lida."))
    return avisos


def conferir(itens: list[ItemLido], consolidados: list[ItemConsolidado], conferir_valores: bool = True,
             conferir_descricao: bool = True, conferir_total_calculado: bool = True) -> list[Aviso]:
    """Avisos para olhar antes de importar. Não bloqueiam a geração.

    conferir_valores=False para empresas cuja planilha não tem valores (Saldo zerado de propósito).
    conferir_descricao=False para empresas cuja planilha não tem descrição (cadastro em branco de propósito).
    conferir_total_calculado=False quando o total informado não é quantidade × unitário (custo médio).
    """
    avisos = []
    avisos += _conferir_layout(itens, conferir_valores)
    for it in itens:
        onde = f"Linha {it.linha}: " if it.linha else ""
        if conferir_descricao and not it.descricao:
            avisos.append(Aviso("Dados", it.codigo, f"{onde}sem descrição."))
        if len(it.ncm) != 8:
            avisos.append(Aviso("Dados", it.codigo, f"{onde}NCM '{it.ncm or 'vazio'}' não tem 8 dígitos."))
        if it.quantidade <= 0:
            avisos.append(Aviso("Dados", it.codigo, f"{onde}quantidade {formatar(it.quantidade, 4)}."))
        if conferir_valores and it.total <= 0:
            avisos.append(Aviso("Dados", it.codigo, f"{onde}valor total {moeda(it.total)}."))
        if (conferir_valores and conferir_total_calculado and it.valor_total is not None
                and it.valor_unitario and it.quantidade):
            calculado = it.quantidade * it.valor_unitario
            if calculado > 0 and abs(calculado - it.valor_total) > max(0.05, calculado * 0.02):
                avisos.append(Aviso("Divergência", it.codigo,
                                    f"{onde}total {moeda(it.valor_total)} não bate com quantidade × "
                                    f"unitário ({moeda(calculado)})."))

    por_codigo: dict[str, list[ItemLido]] = {}
    for it in itens:
        por_codigo.setdefault(it.codigo, []).append(it)
    for codigo, lista in por_codigo.items():
        if len(codigo) > LIMITE_CODIGO_DOMINIO:
            avisos.append(Aviso("Código longo", codigo, f"{len(codigo)} caracteres; o sistema contábil pode guardar só os "
                                                         f"últimos {LIMITE_CODIGO_DOMINIO}."))
        for campo, nome in (("descricao", "Descrição"), ("ncm", "NCM"), ("unidade", "Unidade")):
            valores = {getattr(i, campo) for i in lista}
            if len(valores) > 1:
                usado = getattr(lista[0], campo)
                avisos.append(Aviso("Divergência", codigo, f"{nome} diferente entre as linhas "
                                                           f"({' / '.join(sorted(valores))}); usado '{usado}'."))
    return avisos


# ---------------------------------------------------------------- formatação

def formatar(valor: float, casas: int = 2, fixo: bool = False) -> str:
    """Número no padrão brasileiro. fixo=False tira zeros à direita (28,6 em vez de 28,6000)."""
    s = f"{valor:,.{casas}f}".replace(",", "X").replace(".", ",").replace("X", ".")
    if not fixo and "," in s:
        s = s.rstrip("0").rstrip(",")
    return s


def moeda(valor: float) -> str:
    return "R$ " + formatar(valor, 2, fixo=True)
