"""Gera dois relatórios fictícios de notas fiscais para experimentar o comparador.

    python gerar_relatorios_exemplo.py            -> grava ao lado deste script
    python gerar_relatorios_exemplo.py <pasta>    -> grava na pasta indicada

Saem dois arquivos, do jeito que costumam chegar:

- "relatorio-portal 07-2026.xlsx": com título em toda coluna, chave de acesso de 44
  dígitos, CNPJ do emitente, data, valor, ICMS e CFOP. É a cara de um relatório de
  portal de notas.
- "relatorio-sistema 07-2026.xlsx": sem CNPJ, com cabeçalho esparso (só "Nota",
  "Valor Contábil" e "Data" têm título), um contador interno de lançamento, cabeçalho
  repetido a cada página e linha "Total Geral" no fim. É a cara de um relatório
  exportado de sistema contábil.

Os dois têm as mesmas 40 notas, com três divergências plantadas de propósito:

1. a nota 5560 tem valor diferente nos dois arquivos (R$ 90,00 a mais no sistema);
2. a nota 5564 existe só no portal;
3. a nota 125778 do portal foi digitada no sistema como 125787 (mesmo fornecedor, mesmo
   valor) — aparece como "só no arquivo" dos dois lados, e o comparador aponta a
   parecida do outro lado.

Além disso, três notas de um fornecedor vêm com razão social antiga no sistema e nova
no portal, para o comparador mostrar o aviso de "mesma empresa com nome diferente".

Os CNPJs e as chaves são montados em tempo de execução, com dígitos verificadores
calculados, e não correspondem a nenhuma empresa.
"""
import datetime as dt
import os
import random
import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

PASTA = Path(os.path.abspath(__file__)).parent
NOME_PORTAL = "relatorio-portal 07-2026.xlsx"
NOME_SISTEMA = "relatorio-sistema 07-2026.xlsx"

FORNECEDORES = [
    "COMERCIO MODELO LTDA",
    "INDUSTRIA FICTICIA SA",
    "DISTRIBUIDORA EXEMPLO ME",
    "METALURGICA ALFA LTDA",
    "PAPELARIA BETA EPP",
    "TRANSPORTES GAMA LTDA",
    "QUIMICA DELTA IND E COM LTDA",
    "ELETRICA OMEGA LTDA",
]
# razão social antiga, que ainda aparece no sistema contábil para este fornecedor
NOME_ANTIGO = {"INDUSTRIA FICTICIA SA": "NOVO NOME HOLDING SA"}

# Cada fornecedor tem a própria numeração, como na vida real. Isso importa para o
# comparador: 40 números seguidos (48101, 48102, …) têm densidade 1 e são tomados, com
# razão, por contador interno do sistema; e número de 4 dígitos começando em 1-7 tem
# cara de CFOP. Número da nota i = BASE[i % 8] + (i // 8) * PASSO[i % 8].
BASES = [48101, 1203, 77012, 915, 30450, 2081, 125770, 5560]
PASSOS = [1, 3, 2, 1, 5, 2, 4, 1]

NOTA_VALOR_DIFERENTE = 5560             # nota i=7
NOTA_SO_NO_PORTAL = 5564                # nota i=39
NOTA_DIGITADA_ERRADA = (125778, 125787)  # nota i=22: (número no portal, número digitado no sistema)


def numero_da_nota(i: int) -> int:
    return BASES[i % len(BASES)] + (i // len(BASES)) * PASSOS[i % len(PASSOS)]


def _dv_mod11(digitos: str, pesos: list) -> int:
    soma = sum(int(d) * p for d, p in zip(digitos, pesos))
    resto = soma % 11
    return 0 if resto < 2 else 11 - resto


def cnpj_ficticio(indice: int) -> str:
    """CNPJ com dígitos verificadores válidos, montado a partir de um número qualquer."""
    base = f"{(indice * 7919) % 100000000:08d}0001"
    d1 = _dv_mod11(base, [5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2])
    d2 = _dv_mod11(base + str(d1), [6, 5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2])
    c = base + str(d1) + str(d2)
    return f"{c[:2]}.{c[2:5]}.{c[5:8]}/{c[8:12]}-{c[12:]}"


def chave_ficticia(cnpj: str, serie: int, numero: int, codigo: int) -> str:
    """Chave de acesso de 44 dígitos no formato oficial (UF, AAMM, CNPJ, modelo, série, número...)."""
    so_digitos = "".join(ch for ch in cnpj if ch.isdigit())
    corpo = f"35{'2607'}{so_digitos}55{serie:03d}{numero:09d}1{codigo:08d}"
    pesos = [2, 3, 4, 5, 6, 7, 8, 9]
    soma = sum(int(d) * pesos[i % 8] for i, d in enumerate(reversed(corpo)))
    resto = soma % 11
    dv = 0 if resto < 2 else 11 - resto
    return corpo + str(dv)


def notas() -> list:
    """As 40 notas-base, iguais nos dois lados."""
    sorteio = random.Random(2026)
    saida = []
    for i in range(40):
        numero = numero_da_nota(i)
        fornecedor = FORNECEDORES[i % len(FORNECEDORES)]
        cnpj = cnpj_ficticio(FORNECEDORES.index(fornecedor) + 1)
        valor = round(sorteio.uniform(180, 9800), 2)
        aliquota = sorteio.choice([0.12, 0.18, 0.18, 0.18, 0.04])
        icms = round(valor * aliquota, 2)
        data = dt.date(2026, 7, 1 + (i * 3) % 28)
        cfop = sorteio.choice(["1102", "1102", "2102", "1403"])
        saida.append({
            "numero": numero, "serie": 1, "fornecedor": fornecedor, "cnpj": cnpj,
            "valor": valor, "icms": icms, "data": data, "cfop": cfop,
            "chave": chave_ficticia(cnpj, 1, numero, 11000000 + i),
        })
    return saida


def _cabecalho(ws, linha, titulos):
    for col, t in enumerate(titulos, start=1):
        c = ws.cell(linha, col, t)
        c.font = Font(bold=True)
        c.fill = PatternFill("solid", fgColor="DDDDDD")


def gerar_portal(destino: Path, lista: list) -> Path:
    livro = Workbook()
    ws = livro.active
    ws.title = "Notas"
    _cabecalho(ws, 1, ["Chave de acesso", "Número", "Série", "Emitente", "CNPJ Emitente",
                       "Data Emissão", "Valor Total", "Valor ICMS", "CFOP"])
    linha = 2
    for n in lista:
        ws.cell(linha, 1, n["chave"]).number_format = "@"
        ws.cell(linha, 2, n["numero"])
        ws.cell(linha, 3, n["serie"])
        ws.cell(linha, 4, n["fornecedor"])
        ws.cell(linha, 5, n["cnpj"])
        ws.cell(linha, 6, n["data"]).number_format = "DD/MM/YYYY"
        ws.cell(linha, 7, n["valor"]).number_format = "#,##0.00"
        ws.cell(linha, 8, n["icms"]).number_format = "#,##0.00"
        ws.cell(linha, 9, n["cfop"])
        linha += 1
    for letra, largura in zip("ABCDEFGHI", (48, 10, 7, 32, 20, 12, 14, 12, 8)):
        ws.column_dimensions[letra].width = largura
    livro.save(destino)
    return destino


def gerar_sistema(destino: Path, lista: list) -> Path:
    livro = Workbook()
    ws = livro.active
    ws.title = "Relatorio"
    ws.cell(1, 1, "Relatório de entradas").font = Font(bold=True, size=12)
    ws.cell(2, 1, "Período: 01/07/2026 a 31/07/2026")
    # cabeçalho esparso: só três colunas têm título, como em exportação de sistema
    cabecalho = [None, "Data", "Nota", None, None, "Valor Contábil", None, None, None]

    def escreve_cabecalho(l):
        for col, t in enumerate(cabecalho, start=1):
            if t:
                ws.cell(l, col, t).font = Font(bold=True)

    linha = 4
    escreve_cabecalho(linha)
    linha += 1
    lancamento = 9801
    soma = 0.0
    for i, n in enumerate(lista):
        if i and i % 18 == 0:               # "quebra de página": cabeçalho repetido
            linha += 1
            escreve_cabecalho(linha)
            linha += 1
        valor = n["valor"]
        numero = n["numero"]
        if numero == NOTA_SO_NO_PORTAL:
            continue
        if numero == NOTA_VALOR_DIFERENTE:
            valor = round(valor + 90.00, 2)
        if numero == NOTA_DIGITADA_ERRADA[0]:
            numero = NOTA_DIGITADA_ERRADA[1]
        fornecedor = NOME_ANTIGO.get(n["fornecedor"], n["fornecedor"])
        base_icms = round(valor * 0.9, 2)
        ws.cell(linha, 1, lancamento)
        ws.cell(linha, 2, n["data"]).number_format = "DD/MM/YYYY"
        ws.cell(linha, 3, numero)
        ws.cell(linha, 4, n["serie"])
        ws.cell(linha, 5, fornecedor)
        ws.cell(linha, 6, valor).number_format = "#,##0.00"
        ws.cell(linha, 7, base_icms).number_format = "#,##0.00"
        ws.cell(linha, 8, n["icms"]).number_format = "#,##0.00"
        ws.cell(linha, 9, n["cfop"])
        soma += valor
        lancamento += 1
        linha += 1
    linha += 1
    ws.cell(linha, 5, "Total Geral").font = Font(bold=True)
    ws.cell(linha, 6, round(soma, 2)).number_format = "#,##0.00"
    for letra, largura in zip("ABCDEFGHI", (10, 12, 8, 6, 34, 14, 14, 12, 8)):
        ws.column_dimensions[letra].width = largura
    livro.save(destino)
    return destino


def gerar(pasta: Path) -> tuple:
    pasta.mkdir(parents=True, exist_ok=True)
    lista = notas()
    return gerar_portal(pasta / NOME_PORTAL, lista), gerar_sistema(pasta / NOME_SISTEMA, lista)


if __name__ == "__main__":
    alvo = Path(sys.argv[1]) if len(sys.argv) > 1 else PASTA
    for arquivo in gerar(alvo):
        print("Gravado:", arquivo)
