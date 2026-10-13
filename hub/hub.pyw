"""Hub das Ferramentas Fiscais: uma tela que lista e abre cada ferramenta.

Abra com dois cliques em "Abrir Ferramentas.bat" (na pasta de cima).
Cada ferramenta vira um cartão: o hub confere as bibliotecas que ela precisa e abre a tela.
"""
from __future__ import annotations

import importlib.util
import os
import queue
import subprocess
import sys
import tempfile
import threading
import traceback
from dataclasses import dataclass, field
from pathlib import Path

import tkinter as tk
from tkinter import messagebox, ttk

PASTA = Path(os.path.abspath(__file__)).parent  # abspath preserva a letra de unidade mapeada (resolve() trocaria pelo caminho de rede)
RAIZ = PASTA.parent
sys.path.insert(0, str(PASTA))

import tema  # noqa: E402
from tema import (AMBAR, AMBAR_T, APAGADO, CLARO, PAINEL, TINTA,  # noqa: E402
                  VERDE, VERDE_T, VERMELHO, VERMELHO_T)

VERSAO = "1.1"
TITULO = "Ferramentas Fiscais"
SEM_JANELA = 0x08000000  # CREATE_NO_WINDOW: não pisca console preto ao abrir a ferramenta


@dataclass
class Ferramenta:
    """Uma ferramenta do hub: um .pyw Tkinter e as bibliotecas que ela precisa.

    requisitos: {nome que se importa: nome que o pip instala}.
    """
    chave: str
    nome: str
    descricao: str
    caminho: str
    requisitos: dict = field(default_factory=dict)
    observacao: str = ""

    @property
    def arquivo(self) -> Path:
        return RAIZ / self.caminho

    def faltando(self) -> list[str]:
        """Pacotes do pip que a ferramenta precisa e não estão instalados."""
        ausentes = []
        for modulo, pacote in self.requisitos.items():
            try:
                achou = importlib.util.find_spec(modulo) is not None
            except (ImportError, ValueError):
                achou = False
            if not achou:
                ausentes.append(pacote)
        return ausentes


# A lista é a fonte da verdade: para acrescentar uma ferramenta, basta um item novo aqui.
FERRAMENTAS = [
    Ferramenta(
        chave="bloco-k",
        nome="Bloco K",
        descricao="Lê a planilha de estoque de uma empresa e gera o cadastro de produtos e a "
                  "Posse de Produtos no layout de importação do sistema contábil.",
        caminho="bloco-k/bloco_k.pyw",
        requisitos={"openpyxl": "openpyxl"},
        observacao="Cada empresa tem um leitor em bloco-k/leitores/; o de exemplo mostra o caminho.",
    ),
    Ferramenta(
        chave="comparador-nf",
        nome="Comparador de notas fiscais",
        descricao="Aponte a pasta com os relatórios: ele abre cada planilha, descobre sozinho o "
                  "que cada coluna é e aponta as diferenças de valor, ICMS e CFOP.",
        caminho="comparador-nf/comparador.pyw",
        requisitos={"openpyxl": "openpyxl", "xlrd": "xlrd"},
    ),
    Ferramenta(
        chave="pdf-extrator",
        nome="Extrator de PDF para Excel",
        descricao="Tira a tabela de dentro do PDF e devolve em planilha, escolhendo sozinho o "
                  "melhor jeito de achar as colunas.",
        caminho="pdf-extrator/extrator.pyw",
        requisitos={"pdfplumber": "pdfplumber", "openpyxl": "openpyxl"},
        observacao="PDF digitalizado (que é foto) não tem texto: para esse caso o caminho é o OCR.",
    ),
    Ferramenta(
        chave="xml-nfe",
        nome="XML de NF-e",
        descricao="Separa XML por número de nota e troca o CFOP de itens em lote, sempre gravando "
                  "em outra pasta.",
        caminho="xml-nfe/xml_nfe.pyw",
        requisitos={"lxml": "lxml"},
    ),
]


def executavel(com_console: bool) -> str:
    """Caminho do python.exe (com console) ou pythonw.exe (sem), a partir do que está rodando."""
    atual = Path(sys.executable)
    alvo = atual.with_name("python.exe" if com_console else "pythonw.exe")
    return str(alvo if alvo.exists() else atual)


class Cartao(ttk.Frame):
    """Um retângulo branco por ferramenta: nome, explicação, selo de estado e botões."""

    def __init__(self, pai, ferramenta: Ferramenta, app: "App"):
        super().__init__(pai, style="Cartao.TFrame", padding=(16, 12))
        self.ferramenta = ferramenta
        self.app = app
        self.columnconfigure(0, weight=1)

        linha_topo = ttk.Frame(self, style="Painel.TFrame")
        linha_topo.grid(row=0, column=0, sticky="ew")
        ttk.Label(linha_topo, text=ferramenta.nome, style="Titulo.TLabel").pack(side="left")
        self.selo = tema.etiqueta(linha_topo, "", CLARO, APAGADO)
        self.selo.pack(side="left", padx=(10, 0))

        ttk.Label(self, text=ferramenta.descricao, style="ApagadoPainel.TLabel",
                  wraplength=560, justify="left").grid(row=1, column=0, sticky="w", pady=(4, 0))
        self.var_detalhe = tk.StringVar()
        ttk.Label(self, textvariable=self.var_detalhe, style="ApagadoPainel.TLabel",
                  wraplength=560, justify="left").grid(row=2, column=0, sticky="w", pady=(4, 0))

        botoes = ttk.Frame(self, style="Painel.TFrame")
        botoes.grid(row=0, column=1, rowspan=3, sticky="e", padx=(16, 0))
        self.bt_abrir = ttk.Button(botoes, text="Abrir", style="Primario.TButton",
                                   command=lambda: app.abrir(ferramenta))
        self.bt_abrir.pack(side="top", fill="x")
        self.bt_instalar = ttk.Button(botoes, text="Instalar o que falta",
                                      command=lambda: app.instalar(ferramenta))
        self.bt_pasta = ttk.Button(botoes, text="Ver pasta",
                                   command=lambda: app.abrir_pasta(ferramenta.arquivo.parent))
        self.bt_pasta.pack(side="top", fill="x", pady=(6, 0))

    def atualizar(self):
        f = self.ferramenta
        detalhes = []
        if not f.arquivo.exists():
            self.selo.config(text="NÃO ENCONTRADA", bg=VERMELHO_T, fg=VERMELHO)
            self.var_detalhe.set(f"Falta o arquivo {f.caminho}.")
            self.bt_abrir.state(["disabled"])
            self.bt_instalar.pack_forget()
            return

        faltando = f.faltando()
        if faltando:
            self.selo.config(text="FALTAM BIBLIOTECAS", bg=AMBAR_T, fg=AMBAR)
            detalhes.append("Falta instalar: " + ", ".join(faltando) + ".")
            self.bt_abrir.state(["disabled"])
            self.bt_instalar.pack(side="top", fill="x", pady=(6, 0), before=self.bt_pasta)
        else:
            self.selo.config(text="PRONTA", bg=VERDE_T, fg=VERDE)
            self.bt_abrir.state(["!disabled"])
            self.bt_instalar.pack_forget()

        if f.observacao:
            detalhes.append(f.observacao)
        self.var_detalhe.set("  ".join(detalhes))


class JanelaInstalacao(tk.Toplevel):
    """Janelinha que mostra o pip trabalhando, linha por linha."""

    def __init__(self, pai: "App", ferramenta: Ferramenta):
        super().__init__(pai)
        self.pai = pai
        self.ferramenta = ferramenta
        self.title(f"Instalando o que falta — {ferramenta.nome}")
        self.geometry("760x420")
        self.configure(bg=tema.FUNDO)
        self.transient(pai)

        ttk.Label(self, text=f"Instalando as bibliotecas de {ferramenta.nome}. "
                             f"Pode demorar alguns minutos.",
                  style="Resumo.TLabel").pack(anchor="w", padx=16, pady=(14, 6))
        quadro = ttk.Frame(self)
        quadro.pack(fill="both", expand=True, padx=16)
        self.texto = tk.Text(quadro, bg=PAINEL, fg=TINTA, font=("Consolas", 9), wrap="word",
                             relief="solid", bd=1, highlightthickness=0)
        rolagem = ttk.Scrollbar(quadro, orient="vertical", command=self.texto.yview)
        self.texto.configure(yscrollcommand=rolagem.set, state="disabled")
        self.texto.pack(side="left", fill="both", expand=True)
        rolagem.pack(side="right", fill="y")

        self.bt_fechar = ttk.Button(self, text="Fechar", command=self.destroy, state="disabled")
        self.bt_fechar.pack(anchor="e", padx=16, pady=12)

        self.fila: queue.Queue = queue.Queue()
        threading.Thread(target=self._rodar, daemon=True).start()
        self.after(120, self._ler_fila)

    def _rodar(self):
        faltando = self.ferramenta.faltando()
        if not faltando:
            self.fila.put("Não falta nada para esta ferramenta.\n")
            self.fila.put(None)
            return
        comando = [executavel(True), "-m", "pip", "install", *faltando]
        self.fila.put("> " + " ".join(comando) + "\n\n")
        try:
            processo = subprocess.Popen(comando, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                                        text=True, encoding="utf-8", errors="replace",
                                        creationflags=SEM_JANELA,
                                        cwd=str(self.ferramenta.arquivo.parent))
            for linha in processo.stdout:
                self.fila.put(linha)
            processo.wait()
            self.fila.put("\n" + ("Instalação concluída." if processo.returncode == 0
                                  else f"O pip terminou com erro (código {processo.returncode})."))
        except Exception:
            self.fila.put("\n" + traceback.format_exc())
        self.fila.put(None)

    def _ler_fila(self):
        try:
            while True:
                item = self.fila.get_nowait()
                if item is None:
                    self.bt_fechar.state(["!disabled"])
                    self.pai.atualizar_cartoes()
                    return
                self.texto.configure(state="normal")
                self.texto.insert("end", item)
                self.texto.see("end")
                self.texto.configure(state="disabled")
        except queue.Empty:
            pass
        self.after(120, self._ler_fila)


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(TITULO)
        self.geometry("980x640")
        self.minsize(860, 520)
        self.configure(bg=tema.FUNDO)
        self.var_status = tk.StringVar(value=f"{len(FERRAMENTAS)} ferramentas.")

        tema.estilos(self)
        self._montar()
        self.atualizar_cartoes()

    def _montar(self):
        tema.faixa(self, TITULO, "Escolha a ferramenta e clique em Abrir.", f"v{VERSAO}")

        corpo = ttk.Frame(self, padding=(20, 14, 20, 0))
        corpo.pack(fill="both", expand=True)

        # Lista rolável, para o hub aguentar ganhar ferramentas novas
        tela = tk.Canvas(corpo, bg=tema.FUNDO, highlightthickness=0)
        rolagem = ttk.Scrollbar(corpo, orient="vertical", command=tela.yview)
        self.lista = ttk.Frame(tela)
        self.lista.bind("<Configure>", lambda _e: tela.configure(scrollregion=tela.bbox("all")))
        janela_lista = tela.create_window((0, 0), window=self.lista, anchor="nw")
        tela.bind("<Configure>", lambda e: tela.itemconfigure(janela_lista, width=e.width))
        tela.configure(yscrollcommand=rolagem.set)
        tela.pack(side="left", fill="both", expand=True)
        rolagem.pack(side="right", fill="y")
        # Roda do mouse só enquanto o ponteiro está sobre a lista
        tela.bind("<Enter>", lambda _e: tela.bind_all(
            "<MouseWheel>", lambda ev: tela.yview_scroll(int(-ev.delta / 120), "units")))
        tela.bind("<Leave>", lambda _e: tela.unbind_all("<MouseWheel>"))

        self.cartoes = []
        for f in FERRAMENTAS:
            cartao = Cartao(self.lista, f, self)
            cartao.pack(fill="x", pady=(0, 10), padx=(0, 4))
            self.cartoes.append(cartao)

        rodape = ttk.Frame(self, padding=(20, 8, 20, 12))
        rodape.pack(fill="x")
        ttk.Label(rodape, text="Cada ferramenta também abre sozinha pelo .bat de dentro da pasta dela. "
                               "Para acrescentar uma, inclua um item na lista FERRAMENTAS de hub.pyw.",
                  style="Apagado.TLabel").pack(side="left")
        ttk.Button(rodape, text="Abrir pasta da suíte",
                   command=lambda: self.abrir_pasta(RAIZ)).pack(side="right")

        tema.barra_status(self, self.var_status)

    def atualizar_cartoes(self):
        for c in self.cartoes:
            c.atualizar()

    def abrir(self, f: Ferramenta):
        """Abre a tela da ferramenta e, se ela morrer na largada, mostra o erro.

        Sem isso o pythonw engole a mensagem e o usuário fica olhando para o nada.
        """
        try:
            erros = tempfile.NamedTemporaryFile(mode="w+", suffix=".log", delete=False,
                                                encoding="utf-8", errors="replace")
            processo = subprocess.Popen([executavel(False), str(f.arquivo)],
                                        cwd=str(f.arquivo.parent), creationflags=SEM_JANELA,
                                        stdout=erros, stderr=subprocess.STDOUT)
            self.var_status.set(f"{f.nome}: abrindo numa janela nova…")
            self.after(2500, lambda: self._conferir_abertura(f, processo, Path(erros.name)))
        except Exception as erro:
            messagebox.showerror(TITULO, f"Não consegui abrir {f.nome}.\n\n{erro}")
            self.var_status.set(f"{f.nome}: falhou ao abrir.")

    def _conferir_abertura(self, f: Ferramenta, processo, log: Path):
        """Dois segundos e meio depois: a tela subiu ou morreu?"""
        if processo.poll() is None:
            self.var_status.set(f"{f.nome}: aberta.")
            return
        detalhe = ""
        try:
            detalhe = log.read_text(encoding="utf-8", errors="replace").strip()
            log.unlink(missing_ok=True)
        except OSError:
            pass
        if processo.returncode == 0 and not detalhe:
            self.var_status.set(f"{f.nome}: a janela fechou sozinha.")
            return
        ultimas = "\n".join(detalhe.splitlines()[-12:]) if detalhe else "(sem mensagem)"
        self.var_status.set(f"{f.nome}: não abriu (código {processo.returncode}).")
        messagebox.showerror(TITULO, f"{f.nome} não abriu.\n\nO Python respondeu:\n\n{ultimas}")

    def instalar(self, f: Ferramenta):
        JanelaInstalacao(self, f)

    def abrir_pasta(self, pasta: Path):
        if pasta.exists():
            os.startfile(str(pasta))
        else:
            messagebox.showinfo(TITULO, f"A pasta ainda não existe:\n{pasta}")


if __name__ == "__main__":
    App().mainloop()
