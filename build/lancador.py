r"""Lançador do pacote portátil: é o que vira "Ferramentas Fiscais.exe".

    Ferramentas Fiscais.exe                      abre o hub (hub\hub.pyw, ao lado do executável)
    Ferramentas Fiscais.exe <caminho\tela.pyw>   roda essa tela: é assim que o hub abre cada ferramenta

O hub abre as ferramentas com subprocess.Popen([sys.executable, caminho_do_pyw]). Com o Python
instalado, sys.executable é o pythonw.exe; no pacote, é este executável. Então o lançador aceita
o mesmo argumento e roda o .pyw com runpy, no mesmo processo. Só aceita scripts de dentro da
pasta do pacote.

As pastas das ferramentas ficam ao lado do executável (build\empacotar.py copia); os módulos
delas (tema, nf, pdfx, nucleo_xml, blocok, leitores) são importados do disco, como no
repositório. As bibliotecas (tkinter, openpyxl, xlrd, pdfplumber, lxml) vêm congeladas em
_internal\. Ver build\ferramentas.spec.
"""
import os
import runpy
import sys
import traceback

TITULO = "Ferramentas Fiscais"


def raiz_do_pacote() -> str:
    """Pasta onde estão o executável e as pastas das ferramentas."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))  # rodando do repositório


def avisar(texto: str) -> None:
    """Caixa de mensagem do Windows: o executável não tem console para mostrar erro."""
    try:
        import ctypes
        ctypes.windll.user32.MessageBoxW(None, texto, TITULO, 0x10)
    except Exception:
        print(texto, file=sys.stderr)


def dentro(caminho: str, pasta: str) -> bool:
    try:
        return os.path.commonpath([os.path.abspath(caminho), pasta]) == pasta
    except ValueError:
        return False


def main() -> int:
    sys.dont_write_bytecode = True  # não espalhar __pycache__ pela pasta do pacote
    raiz = raiz_do_pacote()
    if len(sys.argv) > 1:
        script = os.path.abspath(sys.argv[1])
        if not (script.lower().endswith((".pyw", ".py")) and os.path.isfile(script) and dentro(script, raiz)):
            avisar("O executável só abre as telas de dentro da pasta do pacote.\n\nRecebi: " + sys.argv[1])
            return 2
        sys.argv = sys.argv[1:]
    else:
        script = os.path.join(raiz, "hub", "hub.pyw")
        if not os.path.isfile(script):
            avisar("Não achei hub\\hub.pyw ao lado do executável.\n\nPasta: " + raiz)
            return 2
        sys.argv = [script]
    try:
        runpy.run_path(script, run_name="__main__")
    except SystemExit:
        raise
    except Exception:
        detalhe = traceback.format_exc()
        print(detalhe, file=sys.stderr)  # o hub guarda isto quando abre a ferramenta
        avisar(os.path.basename(script) + " não abriu.\n\n" + "\n".join(detalhe.splitlines()[-8:]))
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
