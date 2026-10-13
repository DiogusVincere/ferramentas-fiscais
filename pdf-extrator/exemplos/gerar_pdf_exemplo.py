"""Gera os PDFs de exemplo do extrator, sem biblioteca de PDF: o arquivo é escrito à mão.

    python gerar_pdf_exemplo.py            -> grava ao lado deste script
    python gerar_pdf_exemplo.py <pasta>    -> grava na pasta indicada

- tabela-exemplo.pdf: relatório de estoque fictício em duas páginas, com grade
  desenhada, cabeçalho repetido na segunda página e rodapé "Página N de 2". Serve
  para os três modos do extrator (grade, alinhamento das palavras e uma linha por
  célula).
- digitalizado-exemplo.pdf: uma página que é só imagem (feita com Pillow), para
  mostrar o aviso de PDF digitalizado. Se o Pillow não estiver instalado, é pulado.
"""
import os
import sys
from pathlib import Path

PASTA = Path(os.path.abspath(__file__)).parent

LINHAS = [
    ("P0001", "Parafuso sextavado M6 x 40 zincado", "73181500", "1.200", "0,42", "504,00"),
    ("P0002", "Porca sextavada M6 zincada", "73181600", "1.500", "0,18", "270,00"),
    ("P0003", "Arruela lisa M6", "73182200", "2.000", "0,07", "140,00"),
    ("P0004", "Chapa de aço carbono 1,5 mm", "72085200", "18", "312,50", "5.625,00"),
    ("P0005", "Chapa de aço inox 304 1,0 mm", "72193300", "6", "890,00", "5.340,00"),
    ("P0006", "Tubo de aço quadrado 30 x 30", "73066100", "40", "98,70", "3.948,00"),
    ("P0007", "Barra chata 1 x 1/8 pol", "72166100", "25", "44,90", "1.122,50"),
    ("P0008", "Eletrodo revestido 2,5 mm", "83111000", "12", "86,00", "1.032,00"),
    ("P0009", "Disco de corte 7 pol", "68042290", "60", "9,80", "588,00"),
    ("P0010", "Disco de desbaste 7 pol", "68042290", "30", "14,60", "438,00"),
    ("P0011", "Tinta esmalte sintético cinza", "32089010", "15", "78,90", "1.183,50"),
    ("P0012", "Primer anticorrosivo", "32089010", "10", "92,40", "924,00"),
    ("P0013", "Solvente para esmalte 5 L", "38140010", "8", "61,00", "488,00"),
    ("P0014", "Lixa ferro grão 80", "68052000", "200", "1,95", "390,00"),
    ("P0015", "Lixa ferro grão 120", "68052000", "150", "1,95", "292,50"),
    ("P0016", "Fita isolante 19 mm x 20 m", "39191020", "48", "6,30", "302,40"),
    ("P0017", "Cabo flexível 2,5 mm² preto", "85444900", "7", "189,00", "1.323,00"),
    ("P0018", "Cabo flexível 2,5 mm² azul", "85444900", "5", "189,00", "945,00"),
    ("P0019", "Disjuntor monopolar 20 A", "85362000", "36", "17,80", "640,80"),
    ("P0020", "Tomada 2P+T 10 A branca", "85366990", "80", "8,45", "676,00"),
    ("P0021", "Luva de raspa cano curto", "42032900", "90", "11,20", "1.008,00"),
    ("P0022", "Óculos de proteção incolor", "90049090", "50", "7,90", "395,00"),
    ("P0023", "Protetor auricular tipo concha", "63079010", "25", "29,00", "725,00"),
    ("P0024", "Capacete de segurança amarelo", "65061000", "20", "24,50", "490,00"),
    ("P0025", "Óleo lubrificante 20W50 1 L", "27101932", "24", "32,00", "768,00"),
    ("P0026", "Graxa de lítio 500 g", "27101990", "18", "21,70", "390,60"),
    ("P0027", "Estopa branca (fardo 1 kg)", "63109000", "30", "12,80", "384,00"),
    ("P0028", "Rolamento 6204 2RS", "84821010", "40", "15,90", "636,00"),
    ("P0029", "Correia em V A-42", "40103100", "22", "19,60", "431,20"),
    ("P0030", "Parafuso sextavado M8 x 60 inox", "73181500", "300", "1,35", "405,00"),
]
TITULOS = ("Código", "Descrição", "NCM", "Qtd", "Valor unit.", "Total")
COLUNAS_X = (40, 100, 310, 380, 430, 495)     # início de cada coluna
DIREITA = 560
ALTURA_LINHA = 16
POR_PAGINA = 18


def _texto_pdf(s: str) -> bytes:
    """String PDF entre parênteses, em WinAnsiEncoding (cp1252), com os escapes necessários."""
    b = s.encode("cp1252", errors="replace")
    b = b.replace(b"\\", b"\\\\").replace(b"(", b"\\(").replace(b")", b"\\)")
    return b"(" + b + b")"


def _pagina(numero: int, total_paginas: int, linhas: list) -> bytes:
    """Conteúdo (content stream) de uma página: título, grade e texto."""
    partes = []
    y = 800
    partes.append(b"BT /F2 13 Tf 40 " + str(y).encode() + b" Td " + _texto_pdf("Empresa Exemplo Ltda") + b" Tj ET")
    partes.append(b"BT /F1 9 Tf 40 " + str(y - 16).encode() + b" Td "
                  + _texto_pdf("Relatório de estoque em 31/07/2026") + b" Tj ET")

    topo = y - 44
    todas = [TITULOS] + linhas
    base = topo - ALTURA_LINHA * len(todas)
    # grade: horizontais e verticais
    partes.append(b"0.5 w 0.3 G")
    for i in range(len(todas) + 1):
        yy = topo - i * ALTURA_LINHA
        partes.append(f"{COLUNAS_X[0]} {yy} m {DIREITA} {yy} l S".encode())
    for x in COLUNAS_X + (DIREITA,):
        partes.append(f"{x} {topo} m {x} {base} l S".encode())
    # fundo cinza do cabeçalho
    partes.append(f"0.85 g {COLUNAS_X[0]} {topo - ALTURA_LINHA} {DIREITA - COLUNAS_X[0]} {ALTURA_LINHA} re f 0 g".encode())
    for i in range(len(todas) + 1):          # redesenha as linhas por cima do fundo
        yy = topo - i * ALTURA_LINHA
        partes.append(f"{COLUNAS_X[0]} {yy} m {DIREITA} {yy} l S".encode())

    for i, celulas in enumerate(todas):
        fonte = b"/F2" if i == 0 else b"/F1"
        yy = topo - (i + 1) * ALTURA_LINHA + 4
        for x, texto in zip(COLUNAS_X, celulas):
            partes.append(b"BT " + fonte + b" 9 Tf " + f"{x + 3} {yy}".encode() + b" Td "
                          + _texto_pdf(texto) + b" Tj ET")

    partes.append(b"BT /F1 8 Tf 40 40 Td " + _texto_pdf(f"Página {numero} de {total_paginas}") + b" Tj ET")
    return b"\n".join(partes) + b"\n"


def gerar_tabela(destino: Path) -> Path:
    paginas = [LINHAS[i:i + POR_PAGINA] for i in range(0, len(LINHAS), POR_PAGINA)]
    conteudos = [_pagina(i + 1, len(paginas), p) for i, p in enumerate(paginas)]

    # objetos: 1 catálogo, 2 páginas, 3 fonte normal, 4 fonte negrito, depois página+conteúdo aos pares
    objetos = []
    n_pag = len(conteudos)
    ids_paginas = [5 + 2 * i for i in range(n_pag)]
    objetos.append(b"<< /Type /Catalog /Pages 2 0 R >>")
    kids = b" ".join(f"{i} 0 R".encode() for i in ids_paginas)
    objetos.append(b"<< /Type /Pages /Kids [" + kids + b"] /Count " + str(n_pag).encode() + b" >>")
    objetos.append(b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>")
    objetos.append(b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold /Encoding /WinAnsiEncoding >>")
    for i, conteudo in enumerate(conteudos):
        id_conteudo = ids_paginas[i] + 1
        objetos.append(b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] "
                       b"/Resources << /Font << /F1 3 0 R /F2 4 0 R >> >> "
                       b"/Contents " + f"{id_conteudo} 0 R".encode() + b" >>")
        objetos.append(b"<< /Length " + str(len(conteudo)).encode() + b" >>\nstream\n" + conteudo + b"endstream")

    saida = bytearray(b"%PDF-1.4\n%\xe2\xe3\xcf\xd3\n")
    posicoes = []
    for i, obj in enumerate(objetos, start=1):
        posicoes.append(len(saida))
        saida += f"{i} 0 obj\n".encode() + obj + b"\nendobj\n"
    xref = len(saida)
    saida += f"xref\n0 {len(objetos) + 1}\n".encode()
    saida += b"0000000000 65535 f \n"
    for p in posicoes:
        saida += f"{p:010d} 00000 n \n".encode()
    saida += (f"trailer\n<< /Size {len(objetos) + 1} /Root 1 0 R >>\nstartxref\n{xref}\n%%EOF\n").encode()

    destino.parent.mkdir(parents=True, exist_ok=True)
    destino.write_bytes(bytes(saida))
    return destino


def gerar_digitalizado(destino: Path):
    """Página que é só imagem: o que um scanner produz. Precisa do Pillow."""
    try:
        from PIL import Image, ImageDraw
    except ImportError:
        return None
    imagem = Image.new("RGB", (1240, 1754), "white")
    desenho = ImageDraw.Draw(imagem)
    desenho.text((80, 80), "Empresa Exemplo Ltda - relatorio digitalizado", fill="black")
    y = 160
    for codigo, descricao, ncm, qtd, unit, total in LINHAS[:12]:
        desenho.text((80, y), f"{codigo}   {descricao}   {ncm}   {qtd}   {unit}   {total}", fill="black")
        y += 36
    desenho.rectangle((60, 140, 1180, y + 10), outline="gray")
    destino.parent.mkdir(parents=True, exist_ok=True)
    imagem.save(destino, "PDF", resolution=150)
    return destino


if __name__ == "__main__":
    alvo = Path(sys.argv[1]) if len(sys.argv) > 1 else PASTA
    print("Gravado:", gerar_tabela(alvo / "tabela-exemplo.pdf"))
    digitalizado = gerar_digitalizado(alvo / "digitalizado-exemplo.pdf")
    print("Gravado:", digitalizado) if digitalizado else print("Pillow não instalado: PDF digitalizado pulado.")
