"""Tira tabela de dentro de PDF, sem depender de Ghostscript nem de Java.

Três jeitos, do mais estruturado para o mais bruto:

  tabelas  - pdfplumber acha a grade da tabela (funciona quando o PDF tem linhas
             desenhadas ou espaçamento regular);
  colunas  - quando não há grade, agrupa as palavras pela posição horizontal na
             página e monta as colunas a partir disso. É o que uma pessoa faz ao
             olhar um relatório alinhado: enxerga as colunas pelo alinhamento;
  linhas   - último recurso: cada linha de texto vira uma célula só.

O modo "auto" experimenta os três e fica com o que devolve a tabela mais bem
formada. PDF que é foto (digitalizado) não tem texto nenhum e nenhum dos três
funciona — aí a ferramenta avisa em vez de devolver planilha vazia.
"""
from __future__ import annotations

import re
import statistics
from dataclasses import dataclass, field
from pathlib import Path

import pdfplumber


@dataclass
class Extracao:
    modo: str = ""
    linhas: list = field(default_factory=list)
    paginas: int = 0
    paginas_com_texto: int = 0
    avisos: list = field(default_factory=list)

    @property
    def colunas(self) -> int:
        return max((len(l) for l in self.linhas), default=0)

    @property
    def vazia(self) -> bool:
        return not self.linhas

    @property
    def qualidade(self) -> float:
        """Tabela boa tem várias colunas e largura constante."""
        if len(self.linhas) < 2:
            return 0.0
        larguras = [len([c for c in l if str(c).strip()]) for l in self.linhas]
        media = statistics.fmean(larguras)
        if media < 1.5:
            return 0.2 * min(len(self.linhas) / 50, 1)
        desvio = statistics.pstdev(larguras) if len(larguras) > 1 else 0
        regularidade = 1 / (1 + desvio)
        return min(media / 6, 1.0) * regularidade + min(len(self.linhas) / 200, 0.5)


def _limpar_celula(v) -> str:
    if v is None:
        return ""
    return re.sub(r"\s+", " ", str(v)).strip()


def _normalizar_largura(linhas: list) -> list:
    largura = max((len(l) for l in linhas), default=0)
    return [list(l) + [""] * (largura - len(l)) for l in linhas]


def _tirar_cabecalho_repetido(linhas: list) -> tuple[list, int]:
    """Relatório de várias páginas repete o cabeçalho; fica só o primeiro."""
    if len(linhas) < 3:
        return linhas, 0
    primeira = [c.lower() for c in linhas[0]]
    saida, removidas = [linhas[0]], 0
    for l in linhas[1:]:
        if [c.lower() for c in l] == primeira:
            removidas += 1
            continue
        saida.append(l)
    return saida, removidas


def _por_tabelas(pdf) -> list:
    linhas = []
    for pagina in pdf.pages:
        for tabela in pagina.extract_tables() or []:
            for linha in tabela:
                celulas = [_limpar_celula(c) for c in linha]
                if any(celulas):
                    linhas.append(celulas)
    return linhas


def _agrupar_colunas(palavras: list, tolerancia: float = 12.0) -> list:
    """Descobre onde ficam as colunas pelas posições x em que as palavras começam."""
    inicios = sorted(p["x0"] for p in palavras)
    if not inicios:
        return []
    grupos, atual = [], [inicios[0]]
    for x in inicios[1:]:
        if x - atual[-1] <= tolerancia:
            atual.append(x)
        else:
            grupos.append(atual)
            atual = [x]
    grupos.append(atual)
    # só vale coluna com gente suficiente, senão qualquer palavra solta vira coluna
    minimo = max(2, len(palavras) // 200)
    return [statistics.fmean(g) for g in grupos if len(g) >= minimo]


def _por_colunas(pdf) -> list:
    todas = []
    for pagina in pdf.pages:
        palavras = pagina.extract_words(use_text_flow=False, keep_blank_chars=False)
        if palavras:
            todas.append((pagina, palavras))
    if not todas:
        return []

    referencia = [p for _, palavras in todas for p in palavras]
    limites = _agrupar_colunas(referencia)
    if len(limites) < 2:
        return []

    linhas = []
    for _pagina, palavras in todas:
        por_linha: dict = {}
        for p in palavras:
            chave = round(p["top"] / 3)
            por_linha.setdefault(chave, []).append(p)
        for chave in sorted(por_linha):
            celulas = [""] * len(limites)
            for p in sorted(por_linha[chave], key=lambda w: w["x0"]):
                indice = min(range(len(limites)), key=lambda i: abs(limites[i] - p["x0"]))
                celulas[indice] = (celulas[indice] + " " + p["text"]).strip()
            if any(celulas):
                linhas.append(celulas)
    return linhas


def _por_linhas(pdf) -> list:
    linhas = []
    for pagina in pdf.pages:
        texto = pagina.extract_text() or ""
        for l in texto.split("\n"):
            l = re.sub(r"\s+", " ", l).strip()
            if l:
                linhas.append([l])
    return linhas


MODOS = {
    "tabelas": ("Tabela com grade", _por_tabelas),
    "colunas": ("Colunas pelo alinhamento", _por_colunas),
    "linhas": ("Uma linha de texto por célula", _por_linhas),
}


def extrair(caminho, modo: str = "auto", tirar_repetidos: bool = True) -> Extracao:
    """Extrai a tabela do PDF. modo: 'auto', 'tabelas', 'colunas' ou 'linhas'."""
    caminho = Path(caminho)
    resultado = Extracao(modo=modo)
    try:
        pdf = pdfplumber.open(str(caminho))
    except Exception as erro:
        resultado.avisos.append(f"Não consegui abrir o PDF: {erro}")
        return resultado

    with pdf:
        resultado.paginas = len(pdf.pages)
        resultado.paginas_com_texto = sum(1 for p in pdf.pages if (p.extract_text() or "").strip())
        if not resultado.paginas_com_texto:
            resultado.avisos.append(
                "Este PDF não tem texto: ele é uma imagem digitalizada. Nenhum extrator lê "
                "imagem — para esse caso o caminho é o OCR.")
            return resultado
        if resultado.paginas_com_texto < resultado.paginas:
            resultado.avisos.append(
                f"{resultado.paginas - resultado.paginas_com_texto} de {resultado.paginas} páginas "
                f"não têm texto e ficaram de fora.")

        candidatos = []
        for nome in (MODOS if modo == "auto" else [modo]):
            try:
                linhas = MODOS[nome][1](pdf)
            except Exception as erro:
                resultado.avisos.append(f"Modo '{nome}' falhou: {type(erro).__name__}: {erro}")
                continue
            if linhas:
                candidatos.append((nome, _normalizar_largura(linhas)))

    if not candidatos:
        resultado.avisos.append("Achei texto no PDF, mas não consegui montar tabela nenhuma.")
        return resultado

    melhor_nome, melhor_linhas = None, None
    melhor_nota = -1.0
    for nome, linhas in candidatos:
        teste = Extracao(modo=nome, linhas=linhas)
        if teste.qualidade > melhor_nota:
            melhor_nome, melhor_linhas, melhor_nota = nome, linhas, teste.qualidade

    resultado.modo = melhor_nome
    resultado.linhas = melhor_linhas
    if tirar_repetidos:
        resultado.linhas, removidas = _tirar_cabecalho_repetido(resultado.linhas)
        if removidas:
            resultado.avisos.append(f"Tirei {removidas} repetição(ões) do cabeçalho entre as páginas.")
    if modo == "auto":
        resultado.avisos.insert(0, f"Usei o modo “{MODOS[melhor_nome][0]}”, que deu a tabela "
                                   f"mais bem formada ({len(resultado.linhas)} linhas, "
                                   f"{resultado.colunas} colunas).")
    return resultado


def gerar_excel(extracao: Extracao, destino, primeira_linha_e_cabecalho: bool = True) -> Path:
    """Grava a tabela num .xlsx simples, com cabeçalho escuro (mesmas cores de hub/tema.py)."""
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    destino = Path(destino)
    livro = Workbook()
    ws = livro.active
    ws.title = "Tabela"
    arial = Font(name="Segoe UI", size=9)

    linhas = extracao.linhas
    inicio = 1
    if primeira_linha_e_cabecalho and linhas:
        fonte = Font(name="Segoe UI", size=9, bold=True, color="FFEAEEF3")
        fundo = PatternFill("solid", fgColor="FF1F2A37")
        for col, v in enumerate(linhas[0], start=1):
            c = ws.cell(1, col, v)
            c.font, c.fill = fonte, fundo
            c.alignment = Alignment(vertical="center", wrap_text=True)
        ws.freeze_panes = "A2"
        linhas = linhas[1:]
        inicio = 2

    for i, linha in enumerate(linhas, start=inicio):
        for col, v in enumerate(linha, start=1):
            c = ws.cell(i, col, v)
            c.font = arial

    for col in range(1, extracao.colunas + 1):
        letra = get_column_letter(col)
        largura = max((len(str(l[col - 1])) for l in extracao.linhas if col <= len(l)), default=10)
        ws.column_dimensions[letra].width = min(max(largura + 2, 10), 55)

    destino.parent.mkdir(parents=True, exist_ok=True)
    livro.save(destino)
    return destino
