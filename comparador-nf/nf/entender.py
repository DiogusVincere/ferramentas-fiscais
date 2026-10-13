"""Entende uma planilha de notas fiscais olhando o conteúdo, não o título da coluna.

É o que uma pessoa faz ao abrir o arquivo: procura onde a tabela começa de verdade,
repara que aquela coluna de 44 dígitos é a chave da NFe, que aquela de 4 dígitos
começando em 1 ou 5 é o CFOP, que aquela que dá uns 18% da outra é o ICMS — e só
então lê os números.

Isso é necessário, não é luxo: há relatório de sistema contábil que sai **sem título
nenhum de coluna**, então qualquer ferramenta que dependa do nome da coluna falha
justamente nele.
"""
from __future__ import annotations

import datetime as dt
import re
import statistics
import unicodedata
from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path

from .planilha import Aba, data_de_serial, ler, texto_celula

AMOSTRA = 400          # linhas olhadas para decidir o papel de cada coluna
MIN_LINHAS_TABELA = 2

PAPEIS = ("chave", "nota", "serie", "cnpj", "participante", "data", "valor", "icms", "cfop", "ncm")

NOMES_PAPEL = {
    "chave": "Chave da NFe",
    "nota": "Número da nota",
    "serie": "Série",
    "cnpj": "CNPJ do participante",
    "participante": "Fornecedor / cliente",
    "data": "Data",
    "valor": "Valor contábil",
    "icms": "ICMS",
    "cfop": "CFOP",
    "ncm": "NCM",
}

# Pistas de título. Reforçam o que o conteúdo já disse; sozinhas não decidem.
PISTAS = {
    "chave": ("CHAVE", "CHAVE DE ACESSO", "CHAVE NFE", "CHAVE DA NOTA"),
    "nota": ("NOTA", "NUMERO", "NUM", "N NF", "NF", "NFE", "NUM NOTA", "NUMERO NOTA", "DOCUMENTO", "N NF."),
    "serie": ("SERIE", "N SERIE"),
    "cnpj": ("CNPJ", "CPF", "CNPJ/CPF", "INSCRICAO", "PARTICIPANTE"),
    "participante": ("FORNECEDOR", "CLIENTE", "EMITENTE", "DESTINATARIO", "RAZAO", "RAZAO SOCIAL",
                     "NOME", "EMPRESA", "REMETENTE"),
    "data": ("DATA", "DT", "EMISSAO", "ENTRADA", "SAIDA", "LANCAMENTO", "DATA EMISSAO"),
    "valor": ("VALOR", "VALOR CONTABIL", "VALOR TOTAL", "TOTAL", "VLR", "VL CONTABIL",
              "VALOR DA NOTA", "VALOR NF", "MONTANTE", "LIQUIDO"),
    "icms": ("ICMS", "VALOR ICMS", "VL ICMS", "CREDITO", "CREDITO ICMS", "IMPOSTO"),
    "cfop": ("CFOP", "CFOP ENTRADA", "CFOP SAIDA", "NATUREZA"),
    "ncm": ("NCM", "CLASSIFICACAO FISCAL"),
}

# Títulos que indicam explicitamente que NÃO é o papel (evita "Valor ICMS" virar valor contábil)
ANTI_PISTAS = {
    "valor": ("ICMS", "IPI", "PIS", "COFINS", "ST", "SUBSTITUICAO", "FRETE", "SEGURO", "DESCONTO",
              "UNITARIO", "UNIT", "BASE", "CBS", "IBS", "FCP", "RETENCAO", "ALIQ", "ALIQUOTA"),
    "nota": ("SERIE", "ITEM", "SEQUENCIA", "LANCAMENTO", "PEDIDO", "CFOP", "NCM"),
    # sem isto, "Valor S.T." e "Alíq." viram série: são inteiros pequenos e pouco variados
    "serie": ("VALOR", "ALIQ", "ALIQUOTA", "BASE", "ICMS", "IPI", "PIS", "COFINS", "ST", "FCP",
              "CBS", "IBS", "RETENCAO", "DESCONTO", "FRETE", "QUANTIDADE", "QTD", "CODIGO"),
}


def normalizar(texto) -> str:
    if texto is None:
        return ""
    s = unicodedata.normalize("NFKD", str(texto))
    s = "".join(c for c in s if not unicodedata.combining(c)).upper()
    return re.sub(r"\s+", " ", s).strip(" :.;-_\t\r\n")


def digitos(valor) -> str:
    return re.sub(r"\D", "", texto_celula(valor))


def como_numero(valor) -> float | None:
    """Número de uma célula, aceitando o formato brasileiro."""
    if valor is None or isinstance(valor, bool):
        return None
    if isinstance(valor, (int, float)):
        return float(valor)
    s = str(valor).strip().replace("R$", "").replace("\xa0", "").replace(" ", "")
    if not s or s in ("-", "--", "#N/D", "#VALOR!"):
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


def como_data(valor, aceitar_serial: bool = False) -> dt.date | None:
    """Data de uma célula.

    O número de série do Excel (46183 = 10/06/2026) só conta quando quem chama pede,
    porque número de nota fiscal cai na mesma faixa: 35163 seria 07/04/1996. Quem decide
    é a coluna inteira, não a célula (ver frac_serial em _perfil).
    """
    if isinstance(valor, dt.datetime):
        return valor.date()
    if isinstance(valor, dt.date):
        return valor
    if aceitar_serial:
        serial = data_de_serial(valor)
        if serial:
            return serial.date()
    s = texto_celula(valor)
    m = re.match(r"(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})", s)
    if m:
        d, mes, a = int(m.group(1)), int(m.group(2)), int(m.group(3))
        a = a + 2000 if a < 100 else a
        try:
            return dt.date(a, mes, d)
        except ValueError:
            return None
    m = re.match(r"(\d{4})-(\d{2})-(\d{2})", s)
    if m:
        try:
            return dt.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except ValueError:
            return None
    return None


# ------------------------------------------------------------------ perfil de coluna
@dataclass
class Coluna:
    indice: int
    titulo: str = ""
    papel: str = ""
    confianca: float = 0.0
    exemplos: list = field(default_factory=list)
    preenchidos: int = 0
    distintos: int = 0
    serial: bool = False       # data guardada como número de série do Excel
    por_grupo: float = 0.0     # média de valores preenchidos por nota (1 = nível nota, >1 = nível item)

    @property
    def letra(self) -> str:
        """Letra da coluna como o Excel mostra: A, B, ... Z, AA."""
        n, letras = self.indice, ""
        while True:
            letras = chr(ord("A") + n % 26) + letras
            n = n // 26 - 1
            if n < 0:
                return letras

    @property
    def nome(self) -> str:
        return self.titulo or f"coluna {self.letra}"


def _perfil(valores: list) -> dict:
    """Números crus sobre uma coluna: o que ela parece conter."""
    cheios = [v for v in valores if v not in (None, "")]
    total = len(cheios) or 1
    numeros = [como_numero(v) for v in cheios]
    numeros = [n for n in numeros if n is not None]
    datas = sum(1 for v in cheios if como_data(v) is not None)
    dig = [digitos(v) for v in cheios]

    # Contador interno do sistema (o número do lançamento) ocupa a faixa inteira:
    # uns 700 valores entre 9800 e 10500, densidade ~1. Número de nota de vários fornecedores
    # fica espalhado numa faixa enorme, densidade perto de zero — mesmo quando há
    # sequências dentro dele (várias notas seguidas do mesmo fornecedor).
    inteiros = sorted({int(n) for n in numeros if float(n).is_integer() and 0 < n < 10 ** 12})
    densidade = 0.0
    if len(inteiros) >= 8:
        faixa = inteiros[-1] - inteiros[0] + 1
        densidade = len(inteiros) / faixa if faixa > 0 else 1.0

    # Data em número de série do Excel: só vale se a coluna INTEIRA cair na faixa e
    # cobrir um intervalo curto de tempo. Número de nota também cai na faixa, mas espalhado.
    seriais = [n for n in numeros if 30000 <= n <= 60000]
    frac_serial = 0.0
    if numeros and len(seriais) / len(numeros) > 0.9 and len(seriais) >= 5:
        if max(seriais) - min(seriais) <= 800:      # pouco mais de dois anos
            frac_serial = len(seriais) / total

    return {
        "cheios": len(cheios),
        "distintos": len({texto_celula(v) for v in cheios}),
        "frac_numero": len(numeros) / total,
        "frac_data": datas / total,
        "frac_44": sum(1 for d in dig if len(d) == 44) / total,
        "frac_14": sum(1 for d in dig if len(d) == 14) / total,
        "frac_8": sum(1 for d in dig if len(d) == 8) / total,
        "frac_cfop": sum(1 for n in numeros
                         if float(n).is_integer() and 1000 <= n <= 7999 and str(int(n))[0] in "1234567")
        / total,
        "frac_inteiro": sum(1 for n in numeros if float(n).is_integer()) / total,
        "frac_decimal": sum(1 for n in numeros if not float(n).is_integer()) / total,
        "frac_texto": sum(1 for v in cheios if como_numero(v) is None and como_data(v) is None) / total,
        "densidade": densidade,
        "frac_serial": frac_serial,
        "numeros": numeros,
        "comprimento_medio": statistics.fmean([len(texto_celula(v)) for v in cheios]) if cheios else 0,
    }


def _pontuar(papel: str, p: dict, titulo: str, n_linhas: int) -> float:
    """0 a 1: o quanto esta coluna parece ser este papel, pelo conteúdo."""
    if not p["cheios"]:
        return 0.0
    # Coluna de data não é outra coisa. Sem isso, "30/06/2026 21:22:20" vira CNPJ
    # (14 dígitos) e "10/06/2026" vira NCM (8 dígitos).
    if papel != "data" and (p["frac_data"] > 0.3 or p["frac_serial"] > 0.8):
        return 0.0
    cobertura = p["cheios"] / max(n_linhas, 1)
    variedade = p["distintos"] / max(p["cheios"], 1)
    numeros = p["numeros"]
    nota = 0.0

    if papel == "chave":
        nota = p["frac_44"]
    elif papel == "cnpj":
        nota = p["frac_14"] * (0.7 + 0.3 * (1 - variedade))
    elif papel == "ncm":
        nota = p["frac_8"] * 0.8 if p["frac_8"] > 0.7 else 0.0
    elif papel == "cfop":
        nota = p["frac_cfop"] if p["distintos"] <= 60 else p["frac_cfop"] * 0.4
    elif papel == "data":
        nota = max(p["frac_data"], p["frac_serial"])
    elif papel == "serie":
        # série é 1, 2, 3… — coluna quase toda zero é valor de imposto, não série.
        # O limite de distintos é folgado: há empresa com mais de vinte séries num mês só.
        if numeros and p["frac_inteiro"] > 0.9 and max(numeros) <= 999 and p["distintos"] <= 60:
            com_valor = sum(1 for n in numeros if n >= 1) / len(numeros)
            nota = 0.6 * com_valor if com_valor > 0.5 else 0.0
    elif papel == "nota":
        if numeros and p["frac_inteiro"] > 0.85 and p["frac_44"] < 0.2 and p["frac_14"] < 0.2:
            faixa = 1 <= min(numeros) and max(numeros) <= 999_999_999
            nota = (0.55 + 0.45 * variedade) if faixa else 0.0
            if p["frac_cfop"] > 0.8 or p["frac_8"] > 0.8:
                nota *= 0.2
            if p["densidade"] > 0.7:
                nota *= 0.35   # contador interno do sistema, não número de nota
    elif papel == "valor":
        if numeros and min(numeros) >= -0.01:
            media = statistics.fmean(numeros)
            nota = 0.35 + 0.35 * variedade + (0.3 if p["frac_decimal"] > 0.3 else 0.0)
            if media < 1 or max(numeros) < 10:
                nota *= 0.3
            if p["frac_44"] > 0.2 or p["frac_14"] > 0.2 or p["frac_cfop"] > 0.8 or p["frac_data"] > 0.3:
                nota = 0.0
    elif papel == "icms":
        if numeros and min(numeros) >= -0.01 and p["frac_decimal"] > 0.15:
            nota = 0.25 + 0.25 * variedade
    elif papel == "participante":
        if p["frac_texto"] > 0.8 and p["comprimento_medio"] >= 6:
            nota = 0.4 + 0.4 * min(variedade * 2, 1)

    t = normalizar(titulo)
    alvos = PISTAS.get(papel, ())
    exato = bool(t) and t in alvos
    parcial = bool(t) and any(a in t for a in alvos)
    anti = any(a in t.split() or t.startswith(a + " ") or t.endswith(" " + a)
               for a in ANTI_PISTAS.get(papel, ())) if t else False

    if nota <= 0:
        # O conteúdo não fechou a regra, mas o título diz exatamente o que a coluna é.
        # Uma pessoa confiaria no título aqui; a nota fica baixa de propósito, para
        # perder de qualquer coluna cujo conteúdo sustente o papel de verdade.
        if not exato or anti or not p["cheios"]:
            return 0.0
        nota = 0.45

    # Sem teto: o título é o desempate. "Nota" na coluna certa tem que ganhar do
    # contador interno, que sozinho também pontua alto por ser tudo distinto.
    if exato:
        nota += 0.35
    elif parcial:
        nota += 0.2
    if anti:
        nota *= 0.15
    return nota * (0.5 + 0.5 * min(cobertura * 1.5, 1))


PALAVRAS_TOTAL = ("TOTAL", "TOTAIS", "SUBTOTAL", "SOMA", "RESUMO", "TOTAL GERAL")
PALAVRAS_RODAPE = ("SISTEMA LICENCIADO", "PAGINA", "PÁGINA", "EMITIDO EM", "USUARIO",
                   "ACOMPANHAMENTO DE", "RELATORIO DE", "PERIODO")


def _e_lixo(linha: list, titulos_norm: set) -> str:
    """Diz por que esta linha não é um registro: cabeçalho repetido, total ou rodapé.

    Relatório de sistema contábil costuma repetir o cabeçalho a cada página e fechar com
    linhas como "Total da Situação" e "Total Geral". Sem tirar isso, o
    cabeçalho repetido vira uma "nota" chamada Nota e o texto no meio dos números estraga
    a leitura do que cada coluna é.
    """
    textos = [normalizar(v) for v in linha if isinstance(v, str) and v.strip()]
    if not textos:
        return ""
    if titulos_norm and len([t for t in textos if t in titulos_norm]) >= max(2, len(titulos_norm) // 3):
        return "cabeçalho repetido"
    primeiro = textos[0]
    if any(primeiro.startswith(p) for p in PALAVRAS_TOTAL):
        return "linha de total"
    if any(primeiro.startswith(p) for p in PALAVRAS_RODAPE):
        return "cabeçalho ou rodapé do relatório"
    return ""


def achar_linha_titulos(linhas: list[list]) -> int | None:
    """Linha de títulos, ou None quando não há nenhuma.

    Aceita cabeçalho **esparso**: há relatório que traz só três rótulos em mais de
    trinta colunas ("Nota", "Valor Contábil", "Tipo"), e esses três já resolvem
    as duas colunas que mais importam. O que descarta uma linha é o texto dela se
    repetir nas linhas de baixo — aí é dado, não título.
    """
    melhor, melhor_nota = None, 0.0
    for i, linha in enumerate(linhas[:25]):
        cheias = [v for v in linha if v not in (None, "")]
        if not cheias:
            continue
        textos = [v for v in cheias if como_numero(v) is None and como_data(v) is None
                  and 0 < len(texto_celula(v)) <= 42]
        if not textos:
            continue
        abaixo = [l for l in linhas[i + 1: i + 12] if any(v not in (None, "") for v in l)]
        if len(abaixo) < 2:
            continue
        valores_abaixo = {texto_celula(v) for l in abaixo for v in l if v not in (None, "")}
        repetidos = sum(1 for v in textos if texto_celula(v) in valores_abaixo)
        nota = (len(textos) / len(cheias)) * (1 - repetidos / len(textos)) * min(len(textos) / 3, 1.0)
        if any(como_numero(v) is not None or como_data(v) is not None for l in abaixo for v in l):
            nota += 0.2
        if nota > melhor_nota:
            melhor, melhor_nota = i, nota
    return melhor if melhor_nota >= 0.5 else None


# ------------------------------------------------------------------ nota consolidada
@dataclass
class Nota:
    chave: str = ""
    numero: str = ""
    serie: str = ""
    cnpj: str = ""
    participante: str = ""
    data: dt.date | None = None
    valor: float | None = None
    icms: float | None = None
    cfops: list = field(default_factory=list)
    linhas: list = field(default_factory=list)

    def identidade(self, modo: str) -> str:
        """Como esta nota é identificada no casamento. O modo é o mesmo para os dois arquivos.

        modo: "chave", "cnpj+numero" ou "numero". Se um lado tem chave e o outro não,
        os dois têm que cair no mesmo modo, senão nada casa.
        """
        numero = str(int(self.numero)) if self.numero.isdigit() else self.numero
        if modo == "chave":
            return ("C:" + self.chave) if self.chave else ""
        if modo == "cnpj+numero" and self.cnpj and numero:
            return f"D:{self.cnpj}:{numero}"
        return ("N:" + numero) if numero else ""


@dataclass
class Tabela:
    aba: str
    linha_titulos: int | None
    colunas: list
    total_linhas: int
    notas: list = field(default_factory=list)
    nivel: str = "nota"          # "nota" (uma linha por NF) ou "item" (uma linha por produto)
    avisos: list = field(default_factory=list)

    def coluna(self, papel: str):
        return next((c for c in self.colunas if c.papel == papel), None)

    @property
    def papeis(self) -> dict:
        return {c.papel: c for c in self.colunas if c.papel}

    @property
    def qualidade(self) -> float:
        """O quanto esta tabela serve para comparar notas."""
        p = self.papeis
        if "valor" not in p or not self.notas:
            return 0.0
        # Relatório de nota fiscal sempre traz pelo menos um destes: chave de acesso,
        # CNPJ do participante ou data. Sem nenhum, é outra planilha (estoque, cadastro).
        if not ({"chave", "cnpj", "data"} & set(p)):
            return 0.0
        nota = 1.0 if ("chave" in p or "nota" in p) else 0.0
        nota += 1.0 if "chave" in p else 0.0
        nota += 0.5 if "icms" in p else 0.0
        nota += 0.3 if "cfop" in p else 0.0
        nota += 0.2 if "cnpj" in p else 0.0
        preenchido = sum(1 for n in self.notas if n.valor is not None) / max(len(self.notas), 1)
        nota = nota * (0.4 + 0.6 * preenchido) + min(len(self.notas) / 200, 1.0)
        # Aba que já vem uma linha por nota vale mais do que somar itens: o valor contábil
        # da nota inclui frete, IPI e desconto, que a soma dos produtos não pega.
        if self.nivel == "item":
            nota *= 0.7
        return nota


def entender_aba(aba: Aba) -> Tabela | None:
    """Acha a tabela dentro da aba e diz o que cada coluna é."""
    linhas = aba.linhas
    if len(linhas) < MIN_LINHAS_TABELA:
        return None

    linha_titulos = achar_linha_titulos(linhas)
    inicio = (linha_titulos + 1) if linha_titulos is not None else 0
    titulos = linhas[linha_titulos] if linha_titulos is not None else []
    titulos_norm = {normalizar(v) for v in titulos if isinstance(v, str) and v.strip()}

    corpo, descartadas = [], {}
    for l in linhas[inicio:]:
        if not any(v not in (None, "") for v in l):
            continue
        motivo = _e_lixo(l, titulos_norm)
        if motivo:
            descartadas[motivo] = descartadas.get(motivo, 0) + 1
            continue
        corpo.append(l)
    if len(corpo) < MIN_LINHAS_TABELA:
        return None

    largura = max(len(l) for l in corpo)
    amostra = corpo[:AMOSTRA]

    colunas = []
    perfis = {}
    for c in range(largura):
        valores = [l[c] if c < len(l) else None for l in amostra]
        p = _perfil(valores)
        titulo = texto_celula(titulos[c]) if c < len(titulos) else ""
        col = Coluna(indice=c, titulo=titulo, preenchidos=p["cheios"], distintos=p["distintos"],
                     exemplos=[texto_celula(v) for v in valores if v not in (None, "")][:3])
        colunas.append(col)
        perfis[c] = p

    # cada papel fica com a coluna de maior nota; uma coluna só tem um papel
    candidatos = []
    for col in colunas:
        if not col.preenchidos:
            continue
        for papel in PAPEIS:
            nota = _pontuar(papel, perfis[col.indice], col.titulo, len(amostra))
            if nota > 0.35:
                candidatos.append((nota, papel, col))
    candidatos.sort(key=lambda x: -x[0])
    usados_papel, usadas_coluna = set(), set()
    for nota, papel, col in candidatos:
        if papel in usados_papel or col.indice in usadas_coluna:
            continue
        col.papel, col.confianca = papel, nota
        if papel == "data":
            col.serial = perfis[col.indice]["frac_serial"] > perfis[col.indice]["frac_data"]
        usados_papel.add(papel)
        usadas_coluna.add(col.indice)

    tabela = Tabela(aba=aba.nome, linha_titulos=linha_titulos, colunas=colunas, total_linhas=len(corpo))
    for motivo, quantas in descartadas.items():
        tabela.avisos.append(f"Deixei de fora {quantas} linha(s): {motivo}.")
    _ajustar_icms(tabela, corpo)
    _montar_notas(tabela, corpo)
    return tabela


def _ajustar_icms(tabela: Tabela, corpo: list[list]) -> None:
    """ICMS costuma ficar entre 1% e 30% do valor contábil. Usa isso para separar dos outros valores."""
    col_valor = tabela.coluna("valor")
    col_icms = tabela.coluna("icms")
    if col_valor is None:
        return

    ALIQUOTAS = (0.04, 0.07, 0.12, 0.18, 0.20, 0.25)

    def razao(col, minimo=8):
        """Mediana de icms/valor, só nas linhas em que os dois têm valor."""
        razoes = []
        for l in corpo[:AMOSTRA]:
            v = como_numero(l[col_valor.indice]) if col_valor.indice < len(l) else None
            i = como_numero(l[col.indice]) if col.indice < len(l) else None
            if v and i and v > 0:
                razoes.append(i / v)
        return (statistics.median(razoes), len(razoes)) if len(razoes) >= minimo else (None, len(razoes))

    def perto_de_aliquota(r):
        return any(abs(r - a) <= a * 0.3 for a in ALIQUOTAS)

    if col_icms is not None:
        r, n = razao(col_icms, minimo=5)
        if r is not None and not (0.001 <= r <= 0.45):
            tabela.avisos.append(
                f"'{col_icms.nome}' dá {r * 100:.0f}% do valor contábil, fora da faixa de ICMS: "
                f"não usei como ICMS.")
            col_icms.papel, col_icms.confianca = "", 0.0
            col_icms = None

    if col_icms is None:
        # Sem coluna com cara de ICMS pelo nome, procura a que dá uma alíquota de verdade.
        melhor, melhor_r = None, None
        for col in tabela.colunas:
            if col.papel or col.indice == col_valor.indice or not col.preenchidos:
                continue
            r, n = razao(col)
            if r is None or not perto_de_aliquota(r):
                continue
            if melhor is None or abs(r - 0.18) < abs(melhor_r - 0.18):
                melhor, melhor_r = col, r
        if melhor is not None:
            melhor.papel, melhor.confianca = "icms", 0.55
            tabela.avisos.append(
                f"Peguei '{melhor.nome}' como ICMS: dá {melhor_r * 100:.0f}% do valor contábil.")


def _montar_notas(tabela: Tabela, corpo: list[list]) -> None:
    """Agrupa as linhas por nota. Relatório por produto vira uma linha por NF."""
    p = tabela.papeis
    if "valor" not in p and "chave" not in p and "nota" not in p:
        return

    def pega(linha, papel):
        col = p.get(papel)
        if col is None or col.indice >= len(linha):
            return None
        v = linha[col.indice]
        return None if v in (None, "") else v

    grupos: dict[str, Nota] = {}
    ordem: list[str] = []
    com_cnpj = "cnpj" in p
    for n, linha in enumerate(corpo, start=1):
        chave = digitos(pega(linha, "chave"))
        numero = digitos(pega(linha, "nota")) or texto_celula(pega(linha, "nota"))
        if not chave and not numero:
            continue
        cnpj = digitos(pega(linha, "cnpj"))
        identidade = ("C:" + chave) if chave else (
            f"D:{cnpj}:{numero}" if com_cnpj and cnpj else "N:" + numero)

        nota = grupos.get(identidade)
        if nota is None:
            nota = grupos[identidade] = Nota(chave=chave, numero=numero, cnpj=cnpj)
            ordem.append(identidade)
        nota.linhas.append(n)
        if not nota.participante:
            nota.participante = texto_celula(pega(linha, "participante"))
        if not nota.serie:
            nota.serie = texto_celula(pega(linha, "serie"))
        if nota.data is None:
            col_data = p.get("data")
            nota.data = como_data(pega(linha, "data"),
                                  aceitar_serial=bool(col_data and col_data.serial))
        cfop = pega(linha, "cfop")
        if cfop is not None:
            c = texto_celula(cfop)
            if c and c not in nota.cfops:
                nota.cfops.append(c)

    repetidas = sum(1 for i in ordem if len(grupos[i].linhas) > 1)
    tabela.nivel = "item" if repetidas > len(ordem) * 0.25 else "nota"

    for papel in ("valor", "icms"):
        col = p.get(papel)
        if col is None:
            continue
        por_nota = {}
        for identidade in ordem:
            vals = []
            for n in grupos[identidade].linhas:
                v = como_numero(corpo[n - 1][col.indice]) if col.indice < len(corpo[n - 1]) else None
                if v is not None:
                    vals.append(v)
            por_nota[identidade] = vals
        # Como juntar quando a nota ocupa mais de uma linha. Três casos reais:
        #  - o valor aparece uma vez só (portal, "Valor contábil" na primeira linha da nota)
        #  - o valor se repete igual em todas as linhas (total da nota repetido)
        #  - o valor é diferente em cada linha (o sistema quebra a nota por CFOP) -> somar
        multiplas = [i for i in ordem if len(grupos[i].linhas) > 1]
        unicos = iguais = 0
        for i in multiplas:
            vals = por_nota[i]
            if len(vals) <= 1:
                unicos += 1
            elif max(vals) - min(vals) < 0.005:
                iguais += 1
        n_mult = len(multiplas) or 1
        nivel_nota = not multiplas or (unicos + iguais) / n_mult >= 0.6

        for identidade in ordem:
            vals = por_nota[identidade]
            if not vals:
                continue
            setattr(grupos[identidade], papel, max(vals) if nivel_nota else round(sum(vals), 2))
        cheios = [len(v) for v in por_nota.values() if v]
        col.por_grupo = statistics.fmean(cheios) if cheios else 0
        if multiplas:
            tabela.avisos.append(
                f"'{col.nome}': {len(multiplas)} nota(s) ocupam mais de uma linha e "
                + ("peguei o valor da nota." if nivel_nota else "somei as linhas de cada nota."))

    tabela.notas = [grupos[i] for i in ordem]


# ------------------------------------------------------------------ documento
@dataclass
class Documento:
    caminho: Path
    tabela: Tabela | None = None
    todas_tabelas: list = field(default_factory=list)
    competencia: tuple | None = None      # (ano, mes)
    origem: str = ""
    erro: str = ""
    resultado_anterior: bool = False      # é um Excel de resultado gerado por esta ferramenta

    @property
    def nome(self) -> str:
        return self.caminho.name

    @property
    def ok(self) -> bool:
        return self.tabela is not None and bool(self.tabela.notas)

    @property
    def competencia_texto(self) -> str:
        return f"{self.competencia[1]:02d}/{self.competencia[0]}" if self.competencia else "?"

    def resumo(self) -> str:
        """Uma frase em português sobre o que foi entendido do arquivo."""
        if self.erro:
            return self.erro
        if self.resultado_anterior:
            return "resultado de uma comparação anterior — deixei de fora"
        if not self.ok:
            return "não achei uma tabela de notas neste arquivo"
        t = self.tabela
        partes = [f"{len(t.notas)} nota(s)"]
        if t.nivel == "item":
            partes.append(f"a partir de {t.total_linhas} linhas de produto")
        if self.competencia:
            partes.append(f"competência {self.competencia_texto}")
        if self.origem:
            partes.append(f"cara de {self.origem}")
        achadas = [NOMES_PAPEL[c.papel] for c in t.colunas if c.papel]
        return f"aba '{t.aba}': " + ", ".join(partes) + ". Colunas entendidas: " + ", ".join(achadas) + "."


def _competencia(tabela: Tabela, caminho: Path) -> tuple | None:
    datas = [n.data for n in tabela.notas if n.data]
    if datas:
        comum = Counter((d.year, d.month) for d in datas).most_common(1)[0]
        if comum[1] >= max(2, len(datas) * 0.3):
            return comum[0]
    texto = normalizar(caminho.stem)
    m = re.search(r"(?<!\d)(0[1-9]|1[0-2])[-_./ ]?(20\d{2})(?!\d)", texto)
    if m:
        return (int(m.group(2)), int(m.group(1)))
    m = re.search(r"(?<!\d)(20\d{2})[-_./ ]?(0[1-9]|1[0-2])(?!\d)", texto)
    if m:
        return (int(m.group(1)), int(m.group(2)))
    return None


def _origem(tabela: Tabela, caminho: Path) -> str:
    p = tabela.papeis
    nome = normalizar(caminho.stem) + " " + normalizar(tabela.aba)
    if "chave" in p and tabela.linha_titulos is not None:
        return "relatório do portal (tem chave de acesso)"
    if tabela.linha_titulos is None:
        return "relatório de sistema contábil (sai sem título de coluna)"
    if "PORTAL" in nome or "PREFEITURA" in nome or "SEFAZ" in nome:
        return "portal"
    if "SISTEMA" in nome or "DOMINIO" in nome:
        return "sistema contábil"
    return ""


def entender_arquivo(caminho) -> Documento:
    """Lê o arquivo, escolhe a melhor tabela de notas e explica o que entendeu."""
    caminho = Path(caminho)
    doc = Documento(caminho=caminho)
    try:
        abas = ler(caminho)
    except Exception as erro:
        doc.erro = str(erro)
        return doc

    # Excel que esta ferramenta gerou: não faz sentido comparar resultado com resultado
    nomes_abas = {normalizar(a.nome) for a in abas}
    marcas = {"DIVERGENCIAS", "NAO ENCONTRADOS", "NAO_ENCONTRADOS", "RESUMO"}
    primeira = {normalizar(v) for v in (abas[0].linhas[0] if abas[0].linhas else [])}
    if len(marcas & nomes_abas) >= 2 or {"VALOR ARQ1", "VALOR_ARQ1", "CHAVE_ARQ1"} & primeira:
        doc.resultado_anterior = True
        return doc

    tabelas = []
    for aba in abas:
        try:
            t = entender_aba(aba)
        except Exception as erro:
            doc.erro = f"aba '{aba.nome}': {type(erro).__name__}: {erro}"
            continue
        if t is not None:
            tabelas.append(t)
    doc.todas_tabelas = tabelas
    if not tabelas:
        return doc

    melhor = max(tabelas, key=lambda t: t.qualidade)
    if melhor.qualidade <= 0:
        return doc
    doc.tabela = melhor
    _completar_de_outra_aba(melhor, tabelas)
    doc.competencia = _competencia(melhor, caminho)
    doc.origem = _origem(melhor, caminho)
    return doc


def _completar_de_outra_aba(escolhida: Tabela, tabelas: list) -> None:
    """Busca ICMS e CFOP em outra aba do mesmo arquivo, quando a aba escolhida não tem.

    É o que a pessoa faz: o relatório por nota traz o valor contábil, mas o ICMS só está
    no relatório por produto. Casa pela chave da NFe; sem chave, pelo número da nota.
    """
    faltando = [papel for papel in ("icms", "cfop") if papel not in escolhida.papeis]
    if not faltando or not escolhida.notas:
        return
    for outra in tabelas:
        if outra is escolhida or not outra.notas:
            continue
        disponiveis = [p for p in faltando if p in outra.papeis]
        if not disponiveis:
            continue
        por_chave = {n.chave: n for n in outra.notas if n.chave}
        por_numero = {n.numero: n for n in outra.notas if n.numero}
        casadas = 0
        for nota in escolhida.notas:
            origem = por_chave.get(nota.chave) if nota.chave else None
            if origem is None:
                origem = por_numero.get(nota.numero)
            if origem is None:
                continue
            casadas += 1
            if "icms" in disponiveis and nota.icms is None:
                nota.icms = origem.icms
            if "cfop" in disponiveis and not nota.cfops:
                nota.cfops = list(origem.cfops)
        if casadas >= len(escolhida.notas) * 0.5:
            nomes = " e ".join(NOMES_PAPEL[p] for p in disponiveis)
            escolhida.avisos.append(
                f"{nomes} vieram da aba '{outra.aba}' ({casadas} de {len(escolhida.notas)} notas casadas).")
            for p in disponiveis:
                faltando.remove(p)
        if not faltando:
            return


def entender_pasta(pasta, incluir_subpastas: bool = False, progresso=None) -> list:
    """Entende todas as planilhas da pasta, uma por uma."""
    from .planilha import arquivos_da_pasta
    arquivos = arquivos_da_pasta(pasta, incluir_subpastas)
    documentos = []
    for i, caminho in enumerate(arquivos, start=1):
        if progresso:
            progresso(i, len(arquivos), caminho.name)
        documentos.append(entender_arquivo(caminho))
    return documentos
