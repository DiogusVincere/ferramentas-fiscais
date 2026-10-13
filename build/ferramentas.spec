# -*- mode: python ; coding: utf-8 -*-
r"""Receita do PyInstaller para o pacote portátil das Ferramentas Fiscais.

Rode por build\empacotar.py, que chama
    python -m PyInstaller --noconfirm --clean --distpath <dist> --workpath <trabalho> build\ferramentas.spec
e depois copia as pastas das ferramentas para o lado do executável.

O que sai daqui é "Ferramentas Fiscais.exe" (sem console) + _internal\ com o Python, o tkinter e
as bibliotecas. O código das ferramentas NÃO é congelado: fica ao lado do executável, igual ao
repositório, e o lançador (build\lancador.py) roda hub\hub.pyw dali. Por isso os módulos que
esse código importa entram como hidden imports, lidos dos próprios arquivos.
"""
import os
import re
from pathlib import Path

RAIZ = Path(SPECPATH).parent  # SPECPATH é a pasta deste arquivo (build\)
NOME = "Ferramentas Fiscais"
VERSAO = os.environ.get("FF_VERSAO", "1.0.0")
PASTAS = ["hub", "comparador-nf", "pdf-extrator", "xml-nfe", "bloco-k"]
LOCAIS = {"tema", "nf", "pdfx", "nucleo_xml", "blocok", "leitores"}  # vêm do disco, não do executável


def modulos_usados() -> list:
    """Tudo que o código das ferramentas importa (biblioteca padrão e de terceiros).

    Lê os import de cada arquivo com o ast. Em "from tkinter import ttk, messagebox" cada nome
    também pode ser um submódulo, e esses precisam entrar por extenso (senão o tkinter congelado
    vem sem ttk); importlib.util.find_spec diz se o nome é módulo ou só algo dentro dele.
    """
    import ast
    import importlib.util

    def e_modulo(nome: str) -> bool:
        try:
            return importlib.util.find_spec(nome) is not None
        except (ImportError, ValueError, AttributeError):
            return False

    nomes = set()
    for pasta in PASTAS:
        for arquivo in (RAIZ / pasta).rglob("*"):
            if arquivo.suffix not in (".py", ".pyw"):
                continue
            arvore = ast.parse(arquivo.read_text(encoding="utf-8"), filename=str(arquivo))
            for no in ast.walk(arvore):
                if isinstance(no, ast.Import):
                    candidatos = [a.name for a in no.names]
                elif isinstance(no, ast.ImportFrom) and no.level == 0 and no.module:
                    candidatos = [no.module] + [no.module + "." + a.name for a in no.names]
                else:
                    continue
                for nome in candidatos:
                    if nome.split(".")[0] in LOCAIS or nome == "__future__":
                        continue
                    if "." in nome and not e_modulo(nome):
                        continue
                    nomes.add(nome)
    return sorted(nomes)


def gerar_icone(destino: Path):
    """Ícone do executável: um quadrado escuro com uma planilha, nas cores do tema."""
    try:
        from PIL import Image, ImageDraw
    except ImportError:
        return None
    tam = 256
    imagem = Image.new("RGBA", (tam, tam), (0, 0, 0, 0))
    d = ImageDraw.Draw(imagem)
    d.rounded_rectangle((8, 8, tam - 8, tam - 8), radius=48, fill="#1F2A37")
    d.rounded_rectangle((52, 60, 204, 98), radius=8, fill="#2F5D8A")
    for y in (116, 148, 180):
        d.rounded_rectangle((52, y, 204, y + 20), radius=6, fill="#9DB8D4")
    destino.parent.mkdir(parents=True, exist_ok=True)
    imagem.save(destino, format="ICO",
                sizes=[(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])
    return str(destino)


def info_de_versao():
    """Propriedades do executável (botão direito > Propriedades > Detalhes)."""
    from PyInstaller.utils.win32.versioninfo import (FixedFileInfo, StringFileInfo, StringStruct,
                                                     StringTable, VarFileInfo, VarStruct, VSVersionInfo)
    partes = tuple(int(p) for p in (VERSAO.split(".") + ["0", "0", "0"])[:3]) + (0,)
    return VSVersionInfo(
        ffi=FixedFileInfo(filevers=partes, prodvers=partes, mask=0x3F, flags=0x0, OS=0x40004,
                          fileType=0x1, subtype=0x0, date=(0, 0)),
        kids=[StringFileInfo([StringTable("041604B0", [
            StringStruct("FileDescription", NOME),
            StringStruct("ProductName", NOME),
            StringStruct("FileVersion", VERSAO),
            StringStruct("ProductVersion", VERSAO),
            StringStruct("OriginalFilename", NOME + ".exe"),
            StringStruct("LegalCopyright", "Licença MIT"),
        ])]), VarFileInfo([VarStruct("Translation", [1046, 1200])])])


a = Analysis(
    [str(RAIZ / "build" / "lancador.py")],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=modulos_usados(),
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    # instalados na máquina de quem empacota, mas que nenhuma ferramenta usa
    excludes=["numpy", "pandas", "matplotlib", "scipy", "IPython", "jupyter", "streamlit", "rapidfuzz",
              "win32com", "pythoncom", "pywintypes", "pytest", "test", "tkinter.test", "idlelib",
              "turtle", "turtledemo", "lib2to3", "pydoc_data",
              "PIL._avif"],  # codec de imagem de 7 MB que nenhuma tela usa
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name=NOME,
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    icon=gerar_icone(Path(workpath) / "icone.ico"),
    version=info_de_versao(),
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    name=NOME,
)
