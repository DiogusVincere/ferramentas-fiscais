"""Leitura de planilha sem depender do Excel.

.xlsx/.xlsm pelo openpyxl, .xls pelo xlrd (0,1 s, contra 5-14 s abrindo o Excel),
.csv/.txt direto, detectando separador e codificação.

Devolve sempre a mesma coisa: uma lista de Aba(nome, linhas), com as células já
convertidas — data do Excel vira datetime, número fica número, o resto vira texto.
"""
from __future__ import annotations

import csv
import datetime as dt
import io
import re
from dataclasses import dataclass, field
from pathlib import Path

EXTENSOES = (".xlsx", ".xlsm", ".xls", ".csv", ".txt")
LIMITE_LINHAS = 200_000


class ErroPlanilha(Exception):
    """Problema ao abrir o arquivo, com mensagem pronta para a tela."""


@dataclass
class Aba:
    nome: str
    linhas: list[list] = field(default_factory=list)

    @property
    def util(self) -> bool:
        return any(any(v not in (None, "") for v in linha) for linha in self.linhas)


def _limpar(linhas: list[list]) -> list[list]:
    """Corta linhas e colunas totalmente vazias das bordas."""
    while linhas and not any(v not in (None, "") for v in linhas[-1]):
        linhas.pop()
    if not linhas:
        return []
    largura = max(len(l) for l in linhas)
    usadas = [c for c in range(largura)
              if any(c < len(l) and l[c] not in (None, "") for l in linhas)]
    if not usadas:
        return []
    primeira, ultima = usadas[0], usadas[-1]
    return [[(l[c] if c < len(l) else None) for c in range(primeira, ultima + 1)] for l in linhas]


def _ler_xlsx(dados: bytes) -> list[Aba]:
    import openpyxl
    livro = openpyxl.load_workbook(io.BytesIO(dados), read_only=True, data_only=True)
    abas = []
    for ws in livro.worksheets:
        linhas = []
        for i, linha in enumerate(ws.iter_rows(values_only=True)):
            linhas.append(list(linha))
            if i >= LIMITE_LINHAS:
                break
        abas.append(Aba(ws.title, _limpar(linhas)))
    livro.close()
    return abas


def _ler_xls(dados: bytes) -> list[Aba]:
    import xlrd
    livro = xlrd.open_workbook(file_contents=dados)
    abas = []
    for ws in livro.sheets():
        linhas = []
        for r in range(min(ws.nrows, LIMITE_LINHAS)):
            linha = []
            for c in range(ws.ncols):
                celula = ws.cell(r, c)
                valor = celula.value
                if celula.ctype == xlrd.XL_CELL_DATE:
                    try:
                        valor = xlrd.xldate.xldate_as_datetime(valor, livro.datemode)
                    except Exception:
                        pass
                elif celula.ctype == xlrd.XL_CELL_EMPTY:
                    valor = None
                linha.append(valor)
            linhas.append(linha)
        abas.append(Aba(ws.name, _limpar(linhas)))
    return abas


def _ler_texto(dados: bytes, nome: str) -> list[Aba]:
    for codificacao in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            texto = dados.decode(codificacao)
            break
        except UnicodeDecodeError:
            continue
    else:
        raise ErroPlanilha(f"Não consegui descobrir a codificação de {nome}.")

    amostra = "\n".join(texto.splitlines()[:40])
    separador = max((";", ",", "\t", "|"), key=amostra.count)
    if amostra.count(separador) == 0:
        separador = ";"
    linhas = [list(l) for l in csv.reader(io.StringIO(texto), delimiter=separador)]
    return [Aba(Path(nome).stem, _limpar(linhas[:LIMITE_LINHAS]))]


_SERIAL_MIN, _SERIAL_MAX = 30000, 60000   # 1982 a 2064, faixa plausível de data em serial do Excel


def data_de_serial(valor) -> dt.datetime | None:
    """Converte o número de data do Excel, quando o valor estiver nessa faixa."""
    if isinstance(valor, (int, float)) and not isinstance(valor, bool):
        if _SERIAL_MIN <= valor <= _SERIAL_MAX:
            try:
                return dt.datetime(1899, 12, 30) + dt.timedelta(days=float(valor))
            except (OverflowError, ValueError):
                return None
    return None


def ler(caminho) -> list[Aba]:
    """Todas as abas do arquivo. Levanta ErroPlanilha com mensagem para a tela."""
    caminho = Path(caminho)
    if not caminho.is_file():
        raise ErroPlanilha(f"Não achei o arquivo {caminho.name}.")
    sufixo = caminho.suffix.lower()
    if sufixo not in EXTENSOES:
        raise ErroPlanilha(f"{caminho.name}: não sei ler arquivo {sufixo}.")
    try:
        dados = caminho.read_bytes()
    except PermissionError:
        raise ErroPlanilha(f"{caminho.name} está aberto em outro programa. Feche e tente de novo.")
    except OSError as erro:
        raise ErroPlanilha(f"Não consegui ler {caminho.name}: {erro}")

    try:
        if sufixo in (".xlsx", ".xlsm"):
            abas = _ler_xlsx(dados)
        elif sufixo == ".xls":
            try:
                abas = _ler_xls(dados)
            except Exception:
                # .xls com nome errado: alguns sistemas mandam .xlsx ou HTML chamado de .xls
                abas = _ler_xlsx(dados)
        else:
            abas = _ler_texto(dados, caminho.name)
    except ErroPlanilha:
        raise
    except Exception as erro:
        raise ErroPlanilha(f"{caminho.name}: não consegui abrir ({type(erro).__name__}: {erro}).")

    uteis = [a for a in abas if a.util]
    if not uteis:
        raise ErroPlanilha(f"{caminho.name}: todas as abas estão vazias.")
    return uteis


def arquivos_da_pasta(pasta, incluir_subpastas: bool = False) -> list[Path]:
    """Planilhas da pasta, sem temporários do Excel nem arquivos escondidos."""
    pasta = Path(pasta)
    padrao = "**/*" if incluir_subpastas else "*"
    achados = []
    for p in sorted(pasta.glob(padrao)):
        if not p.is_file() or p.suffix.lower() not in EXTENSOES:
            continue
        if p.name.startswith(("~$", ".")):
            continue
        achados.append(p)
    return achados


def texto_celula(valor) -> str:
    if valor is None:
        return ""
    if isinstance(valor, float) and valor.is_integer():
        valor = int(valor)
    if isinstance(valor, (dt.datetime, dt.date)):
        return valor.strftime("%d/%m/%Y")
    return re.sub(r"\s+", " ", str(valor)).strip()
