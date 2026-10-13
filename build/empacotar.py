r"""Monta o pacote portátil para Windows: a pasta "Ferramentas Fiscais" e o zip dela.

    python build\empacotar.py
        -> dist\Ferramentas Fiscais\  e  dist\ferramentas-fiscais-v1.0.0-windows.zip
    python build\empacotar.py --versao 1.0.1 --saida C:\tmp\ff --sem-zip

Passos:
1. PyInstaller com build\ferramentas.spec: "Ferramentas Fiscais.exe" (sem console) + _internal\
   com o Python, o tkinter e as bibliotecas (openpyxl, xlrd, pdfplumber, lxml).
2. Copia hub\, comparador-nf\, pdf-extrator\, xml-nfe\ e bloco-k\ para o lado do executável, com
   exemplos, geradores e converter_excel.ps1, sem __pycache__, saidas, config.json e *.log. Gera
   os dois relatórios de exemplo do comparador, que no repositório não são versionados.
3. Escreve LEIA-ME.txt, copia a licença e compacta.

Precisa do Python com as bibliotecas do requirements.txt e do PyInstaller
(pip install pyinstaller). O código das ferramentas não é congelado: ver build\lancador.py.
"""
import argparse
import os
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path

RAIZ = Path(os.path.abspath(__file__)).parent.parent
NOME = "Ferramentas Fiscais"
PASTAS = ["hub", "comparador-nf", "pdf-extrator", "xml-nfe", "bloco-k"]
IGNORAR = shutil.ignore_patterns("__pycache__", "saidas", "config.json", "*.log", "*.pyc")

LEIA_ME = r"""FERRAMENTAS FISCAIS - pacote portátil para Windows (versão {versao})

Como usar
1. Descompacte este zip numa pasta qualquer (Área de Trabalho, Documentos...).
2. Abra "Ferramentas Fiscais.exe". Não precisa instalar nada: o Python e as
   bibliotecas vêm dentro da pasta _internal.
3. Na primeira vez o Windows pode avisar que o editor é desconhecido
   (SmartScreen). Clique em "Mais informações" e depois em "Executar assim mesmo".

O que vem dentro
- Ferramentas Fiscais.exe      abre o hub, que lista e abre as quatro ferramentas
- hub, comparador-nf, pdf-extrator, xml-nfe, bloco-k
                               o código de cada tela, igual ao do repositório
- _internal                    o Python e as bibliotecas (não precisa mexer)

Exemplos para experimentar, com dados fictícios, na pasta "exemplos" de cada ferramenta:
- comparador-nf\exemplos   dois relatórios de notas com divergências plantadas:
                           aponte essa pasta na tela e clique em "Ler a pasta"
- pdf-extrator\exemplos    tabela-exemplo.pdf e digitalizado-exemplo.pdf
- xml-nfe\exemplos         1001_nfe-exemplo.xml e 1002_nfe-exemplo.xml
- bloco-k\exemplos         estoque-exemplo.xlsx (empresa "EX1 - EMPRESA EXEMPLO LTDA")

As planilhas que as ferramentas geram ficam em <ferramenta>\saidas.
Os geradores de exemplo também rodam pelo executável, por exemplo:
    "Ferramentas Fiscais.exe" comparador-nf\exemplos\gerar_relatorios_exemplo.py

Código-fonte e documentação: https://github.com/DiogusVincere/ferramentas-fiscais
Licença MIT (LICENCA.txt).
"""


def rodar(comando, **kw):
    print(">", " ".join(str(c) for c in comando))
    subprocess.run([str(c) for c in comando], check=True, **kw)


def congelar(saida: Path, trabalho: Path, versao: str) -> Path:
    """Passo 1: PyInstaller. Devolve a pasta do pacote."""
    ambiente = dict(os.environ, FF_VERSAO=versao)
    rodar([sys.executable, "-m", "PyInstaller", "--noconfirm", "--clean",
           "--distpath", saida, "--workpath", trabalho, RAIZ / "build" / "ferramentas.spec"],
          env=ambiente, cwd=str(RAIZ))
    pacote = saida / NOME
    if not (pacote / f"{NOME}.exe").is_file():
        raise SystemExit(f"O PyInstaller não gerou {pacote / (NOME + '.exe')}.")
    return pacote


def copiar_ferramentas(pacote: Path) -> None:
    """Passo 2: o código das ferramentas vai para o lado do executável."""
    for nome in PASTAS:
        destino = pacote / nome
        if destino.exists():
            shutil.rmtree(destino)
        shutil.copytree(RAIZ / nome, destino, ignore=IGNORAR)
        print("copiado", nome)
    # os relatórios do comparador são gerados na hora (no repositório ficam fora do Git)
    rodar([sys.executable, pacote / "comparador-nf" / "exemplos" / "gerar_relatorios_exemplo.py"])


def documentos(pacote: Path, versao: str) -> None:
    """Passo 3a: LEIA-ME e licença."""
    (pacote / "LEIA-ME.txt").write_text(LEIA_ME.format(versao=versao), encoding="utf-8-sig", newline="\r\n")
    shutil.copy(RAIZ / "LICENSE", pacote / "LICENCA.txt")


def compactar(pacote: Path, zip_destino: Path) -> Path:
    """Passo 3b: zip com a pasta "Ferramentas Fiscais" dentro."""
    if zip_destino.exists():
        zip_destino.unlink()
    with zipfile.ZipFile(zip_destino, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as zf:
        for arquivo in sorted(pacote.rglob("*")):
            if arquivo.is_file():
                zf.write(arquivo, Path(NOME) / arquivo.relative_to(pacote))
    return zip_destino


def main() -> None:
    p = argparse.ArgumentParser(description="Monta o pacote portátil das Ferramentas Fiscais.")
    p.add_argument("--versao", default="1.0.0", help="versão que vai no nome do zip e nas propriedades do .exe")
    p.add_argument("--saida", default=str(RAIZ / "dist"), help="pasta onde ficam o pacote e o zip")
    p.add_argument("--trabalho", default=str(RAIZ / "build" / "pyinstaller"), help="pasta de trabalho do PyInstaller")
    p.add_argument("--sem-zip", action="store_true", help="só monta a pasta, não compacta")
    args = p.parse_args()

    saida, trabalho = Path(args.saida), Path(args.trabalho)
    saida.mkdir(parents=True, exist_ok=True)

    pacote = congelar(saida, trabalho, args.versao)
    copiar_ferramentas(pacote)
    documentos(pacote, args.versao)
    tamanho = sum(f.stat().st_size for f in pacote.rglob("*") if f.is_file())
    print(f"\nPacote: {pacote}  ({tamanho / 1e6:.1f} MB)")
    if not args.sem_zip:
        zip_destino = compactar(pacote, saida / f"ferramentas-fiscais-v{args.versao}-windows.zip")
        print(f"Zip: {zip_destino}  ({zip_destino.stat().st_size / 1e6:.1f} MB)")


if __name__ == "__main__":
    main()
