"""Compara duas relações de notas já entendidas.

A chave do casamento é escolhida pelo que os dois lados têm, da mais firme para a mais
frouxa: chave de acesso da NFe (44 dígitos, não tem como errar) → CNPJ do participante
mais número da nota → só o número da nota. A ferramenta diz qual usou, porque casar só
pelo número junta notas de fornecedores diferentes que têm o mesmo número.
"""
from __future__ import annotations

import datetime as dt
from collections import Counter
from dataclasses import dataclass, field

from .entender import Documento

TOLERANCIA_PADRAO = 0.01      # diferença de até 1 centavo é arredondamento, não divergência


def reais(valor: float) -> str:
    return "R$ " + f"{valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

SITUACOES = {
    "ok": "Confere",
    "valor": "Valor diferente",
    "so_a": "Só no arquivo 1",
    "so_b": "Só no arquivo 2",
}


@dataclass
class Linha:
    identidade: str = ""
    chave: str = ""
    numero: str = ""
    serie: str = ""
    participante: str = ""
    cnpj: str = ""
    data: dt.date | None = None
    valor_a: float | None = None
    valor_b: float | None = None
    icms_a: float | None = None
    icms_b: float | None = None
    cfop_a: str = ""
    cfop_b: str = ""
    situacao: str = "ok"          # sobre o VALOR, que é o que manda na conferência
    icms_difere: bool = False
    cfop_difere: bool = False
    motivos: list = field(default_factory=list)
    parecida_com: str = ""        # nota parecida do outro lado, quando esta ficou sem par

    @property
    def diferenca_valor(self) -> float | None:
        if self.valor_a is None or self.valor_b is None:
            return None
        return round(self.valor_a - self.valor_b, 2)

    @property
    def diferenca_icms(self) -> float | None:
        if self.icms_a is None or self.icms_b is None:
            return None
        return round(self.icms_a - self.icms_b, 2)

    @property
    def diferenca_percentual(self) -> float | None:
        d = self.diferenca_valor
        if d is None or not self.valor_a:
            return None
        return round(d / self.valor_a * 100, 2)

    @property
    def confere(self) -> bool:
        return self.situacao == "ok"

    @property
    def situacao_texto(self) -> str:
        return SITUACOES.get(self.situacao, self.situacao)


@dataclass
class Resultado:
    doc_a: Documento
    doc_b: Documento
    chave_usada: str = ""
    explicacao_chave: str = ""
    linhas: list = field(default_factory=list)
    avisos: list = field(default_factory=list)
    comparou_icms: bool = False
    comparou_cfop: bool = False

    def por_situacao(self, *situacoes) -> list:
        return [l for l in self.linhas if l.situacao in situacoes]

    @property
    def totais(self) -> dict:
        divergentes = self.por_situacao("valor")
        casadas = [l for l in self.linhas if l.situacao in ("ok", "valor")]
        soma_a = sum(l.valor_a or 0 for l in self.linhas)
        soma_b = sum(l.valor_b or 0 for l in self.linhas)
        return {
            "notas": len(self.linhas),
            "casadas": len(casadas),
            "confere": len(self.por_situacao("ok")),
            "divergentes": len(divergentes),
            "so_a": len(self.por_situacao("so_a")),
            "so_b": len(self.por_situacao("so_b")),
            "icms_difere": sum(1 for l in casadas if l.icms_difere),
            "cfop_difere": sum(1 for l in casadas if l.cfop_difere),
            "soma_a": round(soma_a, 2),
            "soma_b": round(soma_b, 2),
            "diferenca": round(soma_a - soma_b, 2),
            "maior_diferenca": max((abs(l.diferenca_valor or 0) for l in divergentes), default=0.0),
        }


def _escolher_chave(a: Documento, b: Documento) -> tuple[str, str]:
    """(modo, explicação). Da chave mais firme para a mais frouxa."""
    pa, pb = a.tabela.papeis, b.tabela.papeis
    com_chave_a = sum(1 for n in a.tabela.notas if n.chave)
    com_chave_b = sum(1 for n in b.tabela.notas if n.chave)
    if com_chave_a > len(a.tabela.notas) * 0.8 and com_chave_b > len(b.tabela.notas) * 0.8:
        return "chave", "casei pela chave de acesso da NFe, que é única — não tem como errar"
    com_cnpj_a = sum(1 for n in a.tabela.notas if n.cnpj)
    com_cnpj_b = sum(1 for n in b.tabela.notas if n.cnpj)
    if ("cnpj" in pa and "cnpj" in pb and com_cnpj_a > len(a.tabela.notas) * 0.8
            and com_cnpj_b > len(b.tabela.notas) * 0.8):
        return ("cnpj+numero",
                "casei por CNPJ do participante + número da nota, porque só um dos arquivos "
                "tem a chave de acesso")
    return ("numero",
            "casei só pelo número da nota, que é a única coisa que os dois arquivos têm em comum. "
            "Confira o fornecedor nas divergências: notas de fornecedores diferentes podem ter "
            "o mesmo número")


def _indexar(notas: list, modo: str) -> tuple[dict, list]:
    indice, repetidas = {}, []
    for nota in notas:
        identidade = nota.identidade(modo)
        if not identidade:
            continue
        if identidade in indice:
            repetidas.append(nota)
            anterior = indice[identidade]
            if nota.valor is not None:
                anterior.valor = round((anterior.valor or 0) + nota.valor, 2)
            if nota.icms is not None:
                anterior.icms = round((anterior.icms or 0) + nota.icms, 2)
            for c in nota.cfops:
                if c not in anterior.cfops:
                    anterior.cfops.append(c)
            continue
        indice[identidade] = nota
    return indice, repetidas


SOCIETARIO = {"LTDA", "SA", "S/A", "ME", "EPP", "EIRELI", "MEI", "CIA", "E", "DE", "DA", "DO",
              "DOS", "DAS", "COM", "COMERCIO", "INDUSTRIA", "IND", "LTD", "S", "A"}


def _palavras(nome: str) -> set:
    """Palavras que identificam a empresa, sem o que é comum a todas."""
    limpo = "".join(c if c.isalnum() or c.isspace() else " " for c in (nome or "").upper())
    return {p for p in limpo.split() if len(p) > 2 and p not in SOCIETARIO}


def parecidos(a: str, b: str) -> float:
    """0 a 1: o quanto dois nomes de empresa são o mesmo. Sem depender de biblioteca."""
    pa, pb = _palavras(a), _palavras(b)
    if not pa or not pb:
        return 0.0
    return len(pa & pb) / min(len(pa), len(pb))


def _sugerir_parecidas(res: "Resultado", tolerancia: float) -> None:
    """Para cada nota sem par, aponta a nota parecida do outro lado.

    Sem isso a tela só diz "só no arquivo 1" e a pessoa vai caçar na mão. Exemplo típico:
    a nota 42957 no portal e a 42966 no sistema, mesmo fornecedor e mesmo valor — é a
    mesma nota com o número digitado errado, e é isso que o usuário precisa ver.
    """
    notas = {"so_a": (res.doc_b.tabela.notas, "no arquivo 2"),
             "so_b": (res.doc_a.tabela.notas, "no arquivo 1")}
    casadas = {l.numero for l in res.linhas if l.situacao in ("ok", "valor")}

    for linha in res.linhas:
        if linha.situacao not in notas:
            continue
        outros, onde = notas[linha.situacao]
        valor = linha.valor_a if linha.situacao == "so_a" else linha.valor_b
        candidatos = []
        for n in outros:
            mesmo_valor = (valor is not None and n.valor is not None
                           and abs(n.valor - valor) <= max(tolerancia, 0.01))
            semelhanca = parecidos(linha.participante, n.participante)
            if mesmo_valor or semelhanca >= 0.8:
                candidatos.append((mesmo_valor, semelhanca, n))
        if not candidatos:
            continue
        candidatos.sort(key=lambda c: (-int(c[0]), -c[1]))
        partes = []
        for mesmo_valor, semelhanca, n in candidatos[:3]:
            marca = []
            if mesmo_valor:
                marca.append("mesmo valor")
            if semelhanca >= 0.8:
                marca.append("mesmo fornecedor")
            if n.numero in casadas:
                marca.append("já casada com outra")
            partes.append(f"nº {n.numero} ({', '.join(marca)})")
        linha.parecida_com = f"parecida {onde}: " + "; ".join(partes)


def _direcao(notas: list) -> str:
    """'entrada' (CFOP 1/2/3), 'saída' (5/6/7) ou '' quando vem misturado."""
    contagem = Counter()
    for nota in notas:
        for c in nota.cfops:
            d = "".join(ch for ch in str(c) if ch.isdigit())
            if len(d) == 4:
                contagem["entrada" if d[0] in "123" else "saída" if d[0] in "567" else "?"] += 1
    if not contagem:
        return ""
    lado, quantos = contagem.most_common(1)[0]
    return lado if quantos >= sum(contagem.values()) * 0.8 and lado != "?" else ""


def _mesmo_cfop(a: list, b: list) -> bool:
    """CFOP de entrada e de saída são espelhos: 1102 de um lado é 5102 do outro."""
    if not a or not b:
        return True
    def familias(cfops):
        f = set()
        for c in cfops:
            d = "".join(ch for ch in str(c) if ch.isdigit())
            if len(d) == 4:
                f.add(d[1:])          # tira o primeiro dígito (dentro/fora do estado, entrada/saída)
        return f
    fa, fb = familias(a), familias(b)
    return not fa or not fb or bool(fa & fb)


def comparar(doc_a: Documento, doc_b: Documento, tolerancia: float = TOLERANCIA_PADRAO,
             tolerancia_icms: float | None = None, comparar_cfop: bool = True) -> Resultado:
    """Casa as notas dos dois arquivos e classifica cada uma."""
    if not doc_a.ok or not doc_b.ok:
        raise ValueError("Os dois arquivos precisam ter uma tabela de notas entendida.")
    if tolerancia_icms is None:
        tolerancia_icms = tolerancia

    modo, explicacao = _escolher_chave(doc_a, doc_b)
    res = Resultado(doc_a=doc_a, doc_b=doc_b, chave_usada=modo, explicacao_chave=explicacao)

    indice_a, repetidas_a = _indexar(doc_a.tabela.notas, modo)
    indice_b, repetidas_b = _indexar(doc_b.tabela.notas, modo)
    for repetidas, doc in ((repetidas_a, doc_a), (repetidas_b, doc_b)):
        if repetidas:
            exemplos = ", ".join(sorted({n.numero for n in repetidas})[:5])
            res.avisos.append(
                f"{doc.nome}: {len(repetidas)} nota(s) apareciam mais de uma vez com a mesma "
                f"identidade e foram somadas ({exemplos}).")

    tem_icms_a = any(n.icms is not None for n in doc_a.tabela.notas)
    tem_icms_b = any(n.icms is not None for n in doc_b.tabela.notas)
    res.comparou_icms = tem_icms_a and tem_icms_b
    tem_cfop_a = any(n.cfops for n in doc_a.tabela.notas)
    tem_cfop_b = any(n.cfops for n in doc_b.tabela.notas)
    res.comparou_cfop = comparar_cfop and tem_cfop_a and tem_cfop_b

    if not res.comparou_icms and (tem_icms_a or tem_icms_b):
        faltou = doc_b.nome if tem_icms_a else doc_a.nome
        res.avisos.append(f"Não comparei ICMS: não achei essa coluna em {faltou}.")
    if comparar_cfop and not res.comparou_cfop and (tem_cfop_a or tem_cfop_b):
        faltou = doc_b.nome if tem_cfop_a else doc_a.nome
        res.avisos.append(f"Não comparei CFOP: não achei essa coluna em {faltou}.")

    # Um arquivo com CFOP de entrada e outro com CFOP de saída são os dois lados da mesma
    # operação: o fornecedor emite 5102, quem recebe escritura 1102. Comparar os dois
    # apontaria divergência em toda nota, sem querer dizer nada.
    if res.comparou_cfop:
        dir_a, dir_b = _direcao(doc_a.tabela.notas), _direcao(doc_b.tabela.notas)
        if dir_a and dir_b and dir_a != dir_b:
            res.comparou_cfop = False
            res.avisos.append(
                f"Não comparei CFOP: {doc_a.nome} traz CFOP de {dir_a} e {doc_b.nome} de {dir_b}. "
                f"São os dois lados da mesma operação (1102 aqui é 5102 lá), então a diferença "
                f"não quer dizer nada.")

    vistas_b = set()
    nomes_diferentes: dict = {}
    for identidade, a in indice_a.items():
        b = indice_b.get(identidade)
        linha = Linha(identidade=identidade, chave=a.chave, numero=a.numero, serie=a.serie,
                      participante=a.participante, cnpj=a.cnpj, data=a.data,
                      valor_a=a.valor, icms_a=a.icms, cfop_a=", ".join(a.cfops))
        if b is None:
            linha.situacao = "so_a"
            res.linhas.append(linha)
            continue
        vistas_b.add(identidade)
        linha.valor_b, linha.icms_b = b.valor, b.icms
        linha.cfop_b = ", ".join(b.cfops)
        if not linha.participante:
            linha.participante = b.participante

        if a.valor is None or b.valor is None:
            linha.situacao = "valor"
            linha.motivos.append("um dos arquivos não tem valor para esta nota")
        elif abs(a.valor - b.valor) > tolerancia:
            linha.situacao = "valor"
            linha.motivos.append("valor difere em " + reais(abs(a.valor - b.valor)))
        else:
            linha.situacao = "ok"

        if res.comparou_icms and a.icms is not None and b.icms is not None:
            if abs(a.icms - b.icms) > tolerancia_icms:
                linha.icms_difere = True
                linha.motivos.append("ICMS difere em " + reais(abs(a.icms - b.icms)))
        if res.comparou_cfop and not _mesmo_cfop(a.cfops, b.cfops):
            linha.cfop_difere = True
            linha.motivos.append(f"CFOP {', '.join(a.cfops)} x {', '.join(b.cfops)}")
        # casando só pelo número, duas notas de fornecedores diferentes casam caladas
        if modo == "numero" and a.participante and b.participante:
            if parecidos(a.participante, b.participante) < 0.5:
                linha.motivos.append(f"fornecedores diferentes: “{a.participante[:30]}” x "
                                     f"“{b.participante[:30]}” — pode não ser a mesma nota")
                nomes_diferentes[(a.participante, b.participante)] = \
                    nomes_diferentes.get((a.participante, b.participante), 0) + 1
        res.linhas.append(linha)

    for identidade, b in indice_b.items():
        if identidade in vistas_b:
            continue
        res.linhas.append(Linha(identidade=identidade, chave=b.chave, numero=b.numero,
                                serie=b.serie, participante=b.participante, cnpj=b.cnpj, data=b.data,
                                valor_b=b.valor, icms_b=b.icms, cfop_b=", ".join(b.cfops),
                                situacao="so_b"))

    # o mesmo par de nomes repetindo costuma ser a mesma empresa com razão social antiga
    # e nova nos dois sistemas ("COMERCIO MODELO LTDA" num lado, "MODELO PARTICIPACOES" no outro)
    for (nome_a, nome_b), quantas in sorted(nomes_diferentes.items(), key=lambda x: -x[1]):
        if quantas >= 3:
            res.avisos.append(
                f"{quantas} notas casaram pelo número mas o fornecedor tem nome diferente nos dois "
                f"arquivos: “{nome_a}” x “{nome_b}”. Costuma ser a mesma empresa com razão social "
                f"antiga e nova — confira uma e, se for, o resto segue.")

    _sugerir_parecidas(res, tolerancia)

    # o que precisa de olho primeiro: valor diferente, depois nota faltando, depois imposto
    ordem = {"valor": 0, "so_a": 1, "so_b": 2, "ok": 3}
    res.linhas.sort(key=lambda l: (ordem.get(l.situacao, 9),
                                   0 if (l.icms_difere or l.cfop_difere) else 1,
                                   -abs(l.diferenca_valor or 0)))

    if doc_a.competencia and doc_b.competencia and doc_a.competencia != doc_b.competencia:
        res.avisos.append(
            f"Atenção: os arquivos são de competências diferentes "
            f"({doc_a.competencia_texto} e {doc_b.competencia_texto}).")
    return res


def _numeros(doc) -> set:
    saida = set()
    for n in doc.tabela.notas:
        if n.numero:
            saida.add(str(int(n.numero)) if n.numero.isdigit() else n.numero)
    return saida


def mesmo_relatorio(a: Documento, b: Documento) -> bool:
    """Os dois arquivos são o mesmo relatório (ex.: exportado em .xls e em .xlsx)."""
    if len(a.tabela.notas) != len(b.tabela.notas):
        return False
    soma_a = sum(n.valor or 0 for n in a.tabela.notas)
    soma_b = sum(n.valor or 0 for n in b.tabela.notas)
    if abs(soma_a - soma_b) > 0.01:
        return False
    na, nb = _numeros(a), _numeros(b)
    return bool(na) and len(na & nb) >= len(na) * 0.99


def sugerir_par(documentos: list) -> tuple:
    """Escolhe os dois arquivos que mais têm cara de par para comparar.

    O que pesa: as notas serem as mesmas (é o que uma pessoa olha primeiro), mesma
    competência e origens diferentes. O mesmo relatório salvo em dois formatos é
    descartado — comparar um arquivo com a cópia dele não diz nada.
    """
    bons = [d for d in documentos if d.ok]
    if len(bons) < 2:
        return (bons[0] if bons else None, None)

    numeros = {id(d): _numeros(d) for d in bons}
    melhor, melhor_nota = (bons[0], bons[1]), -99.0
    for i, a in enumerate(bons):
        for b in bons[i + 1:]:
            nota = 0.0
            if a.competencia and b.competencia:
                nota += 3.0 if a.competencia == b.competencia else -2.0
            if a.origem and b.origem and a.origem != b.origem:
                nota += 1.5
            na, nb = numeros[id(a)], numeros[id(b)]
            menor_conjunto = min(len(na), len(nb)) or 1
            sobreposicao = len(na & nb) / menor_conjunto
            nota += sobreposicao * 2.5              # as mesmas notas dos dois lados
            menor = min(len(a.tabela.notas), len(b.tabela.notas))
            maior = max(len(a.tabela.notas), len(b.tabela.notas), 1)
            nota += (menor / maior) * 1.0
            nota += min(menor / 300, 1.0)
            if mesmo_relatorio(a, b):
                nota -= 8.0
            if nota > melhor_nota:
                melhor, melhor_nota = (a, b), nota
    return melhor
