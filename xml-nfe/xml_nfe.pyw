"""XML de NF-e: separa notas por número e troca o CFOP de itens em lote.

Abra pelo hub (Ferramentas Fiscais) ou com dois cliques em "Abrir XML de NF-e.bat".
As duas ferramentas trabalham em dois tempos: "Conferir" só olha e mostra o que
vai acontecer; o botão de executar só age depois que você conferiu.
O XML de origem nunca é alterado — a troca de CFOP grava cópias em outra pasta.
"""
from __future__ import annotations

import os
import queue
import sys
import threading
import traceback
from pathlib import Path

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

PASTA = Path(os.path.abspath(__file__)).parent  # abspath preserva a letra de unidade mapeada
RAIZ = PASTA.parent
sys.path.insert(0, str(PASTA))
sys.path.insert(0, str(RAIZ / "hub"))

import tema  # noqa: E402
import nucleo_xml as nucleo  # noqa: E402
from tema import VERMELHO  # noqa: E402

VERSAO = "1.0"
TITULO = "XML de NF-e"


class Aba(ttk.Frame):
    """Base das duas abas: campos em cima, tabela no meio, botões e status embaixo."""

    def __init__(self, pai, app: "App"):
        super().__init__(pai, padding=(16, 14, 16, 10))
        self.app = app
        self.fila: queue.Queue = queue.Queue()
        self.trabalhando = False

    def escolher_pasta(self, variavel: tk.StringVar, titulo: str):
        atual = variavel.get().strip()
        pasta = filedialog.askdirectory(title=titulo, initialdir=atual or str(RAIZ))
        if pasta:
            variavel.set(os.path.normpath(pasta))

    def em_fundo(self, tarefa, terminou):
        """Roda a tarefa numa thread e devolve o resultado pela fila, sem travar a tela.

        Quem chama já tem que ter lido as variáveis do Tkinter: dentro da thread
        elas não podem ser lidas ("main thread is not in main loop").
        """
        if self.trabalhando:
            return
        self.trabalhando = True
        self.app.ocupado(True)

        def correr():
            try:
                self.fila.put(("ok", tarefa()))
            except Exception:
                self.fila.put(("erro", traceback.format_exc()))

        threading.Thread(target=correr, daemon=True).start()
        self.after(100, lambda: self._checar(terminou))

    def _checar(self, terminou):
        try:
            tipo, dados = self.fila.get_nowait()
        except queue.Empty:
            self.after(100, lambda: self._checar(terminou))
            return
        self.trabalhando = False
        self.app.ocupado(False)
        if tipo == "erro":
            messagebox.showerror(TITULO, "Deu erro:\n\n" + dados.strip().splitlines()[-1])
            self.app.var_status.set("Deu erro. Nada foi alterado.")
            return
        terminou(dados)


class AbaSeparar(Aba):
    """Move (ou copia) os XML de certas notas para outra pasta."""

    def __init__(self, pai, app):
        super().__init__(pai, app)
        self.var_origem = tk.StringVar()
        self.var_destino = tk.StringVar()
        self.var_sub = tk.BooleanVar(value=True)
        self.var_dentro = tk.BooleanVar(value=False)
        self.var_modo = tk.StringVar(value="mover")
        self.movimentos: list = []

        self.columnconfigure(2, weight=1)

        ttk.Label(self, text="1", style="Passo.TLabel").grid(row=0, column=0, sticky="w", padx=(0, 8))
        ttk.Label(self, text="Pasta com os XMLs").grid(row=0, column=1, sticky="w", padx=(0, 12))
        linha = ttk.Frame(self)
        linha.grid(row=0, column=2, sticky="ew", pady=4)
        linha.columnconfigure(0, weight=1)
        self.ent_origem = ttk.Entry(linha, textvariable=self.var_origem)
        self.ent_origem.grid(row=0, column=0, sticky="ew")
        ttk.Button(linha, text="Procurar…",
                   command=lambda: self.escolher_pasta(self.var_origem, "Pasta com os XMLs")
                   ).grid(row=0, column=1, padx=(8, 0))
        ttk.Checkbutton(self, text="incluir subpastas", variable=self.var_sub).grid(
            row=1, column=2, sticky="w")

        ttk.Label(self, text="2", style="Passo.TLabel").grid(row=2, column=0, sticky="nw", padx=(0, 8), pady=(8, 0))
        ttk.Label(self, text="Números das notas").grid(row=2, column=1, sticky="nw", padx=(0, 12), pady=(8, 0))
        quadro = ttk.Frame(self)
        quadro.grid(row=2, column=2, sticky="ew", pady=(8, 0))
        quadro.columnconfigure(0, weight=1)
        self.txt_numeros = tk.Text(quadro, height=3, font=(tema.FONTE, 10), relief="solid", bd=1,
                                   highlightthickness=0, bg=tema.PAINEL, fg=tema.TINTA)
        self.txt_numeros.grid(row=0, column=0, sticky="ew")
        ttk.Label(self, text="Um por linha, ou separados por vírgula ou espaço. Procura os arquivos que "
                             "começam com o número e um traço baixo (ex.: 12345_...xml).",
                  style="Apagado.TLabel", wraplength=620, justify="left").grid(row=3, column=2, sticky="w")
        ttk.Checkbutton(self, text="procurar também pelo número dentro do XML (nNF) — mais devagar",
                        variable=self.var_dentro).grid(row=4, column=2, sticky="w")

        ttk.Label(self, text="3", style="Passo.TLabel").grid(row=5, column=0, sticky="w", padx=(0, 8), pady=(8, 0))
        ttk.Label(self, text="Levar para").grid(row=5, column=1, sticky="w", padx=(0, 12), pady=(8, 0))
        linha2 = ttk.Frame(self)
        linha2.grid(row=5, column=2, sticky="ew", pady=(8, 0))
        linha2.columnconfigure(0, weight=1)
        self.ent_destino = ttk.Entry(linha2, textvariable=self.var_destino)
        self.ent_destino.grid(row=0, column=0, sticky="ew")
        ttk.Button(linha2, text="Procurar…",
                   command=lambda: self.escolher_pasta(self.var_destino, "Pasta de destino")
                   ).grid(row=0, column=1, padx=(8, 0))

        acoes = ttk.Frame(self)
        acoes.grid(row=6, column=0, columnspan=3, sticky="ew", pady=(12, 6))
        ttk.Radiobutton(acoes, text="Mover", variable=self.var_modo, value="mover").pack(side="left")
        ttk.Radiobutton(acoes, text="Copiar", variable=self.var_modo, value="copiar").pack(side="left", padx=(12, 24))
        ttk.Button(acoes, text="Conferir", command=self._conferir).pack(side="left")
        self.bt_executar = ttk.Button(acoes, text="Mover arquivos", style="Primario.TButton",
                                      command=self._executar, state="disabled")
        self.bt_executar.pack(side="left", padx=(8, 0))
        self.var_modo.trace_add("write", lambda *_: self.bt_executar.config(
            text="Mover arquivos" if self.var_modo.get() == "mover" else "Copiar arquivos"))

        self.tabela = tema.Tabela(self, [
            ("nota", "Nota", 90, "w", False),
            ("arquivo", "Arquivo", 420, "w", True),
            ("situacao", "Situação", 320, "w", True)])
        self.tabela.grid(row=7, column=0, columnspan=3, sticky="nsew")
        self.rowconfigure(7, weight=1)

        # mexeu em algum campo, o plano não vale mais: senão o "Mover" usaria a pasta antiga
        for var in (self.var_origem, self.var_destino, self.var_sub, self.var_dentro):
            var.trace_add("write", lambda *_: self._invalidar())
        self.txt_numeros.bind("<KeyRelease>", lambda _e: self._invalidar())

    def _invalidar(self):
        if self.movimentos:
            self.app.var_status.set("Os campos mudaram: clique em Conferir de novo.")
        self.movimentos = []
        self.tabela.limpar()
        self.bt_executar.state(["disabled"])

    def _conferir(self):
        origem = Path(self.var_origem.get().strip())
        destino = Path(self.var_destino.get().strip())
        numeros = nucleo.separar_valores(self.txt_numeros.get("1.0", "end"))
        if not origem.is_dir():
            messagebox.showwarning(TITULO, "Escolha a pasta onde estão os XMLs.")
            return
        if not numeros:
            messagebox.showwarning(TITULO, "Digite pelo menos um número de nota.")
            return
        if not str(destino).strip() or str(destino) == ".":
            messagebox.showwarning(TITULO, "Escolha a pasta de destino.")
            return

        subpastas, procurar_dentro = self.var_sub.get(), self.var_dentro.get()
        self.bt_executar.state(["disabled"])
        self.app.var_status.set("Procurando os arquivos…")
        self.em_fundo(
            lambda: nucleo.planejar_separacao(origem, destino, numeros, subpastas, procurar_dentro),
            self._mostrar)

    def _mostrar(self, resultado):
        self.movimentos, sem_nada = resultado
        self.tabela.limpar()
        for m in self.movimentos:
            tags = ("alerta",) if m.situacao != "ok" else ()
            situacao = "pronto para ir" if m.situacao == "ok" else m.detalhe or m.situacao
            self.tabela.inserir((m.numero, m.arquivo.name, situacao), tags)
        for numero in sem_nada:
            self.tabela.inserir((numero, "—", "nenhum arquivo encontrado com esse número"), ("alerta",))

        prontos = sum(1 for m in self.movimentos if m.situacao == "ok")
        if prontos:
            self.bt_executar.state(["!disabled"])
        partes = [f"{prontos} arquivo(s) prontos"]
        if len(self.movimentos) - prontos:
            partes.append(f"{len(self.movimentos) - prontos} já existem no destino")
        if sem_nada:
            partes.append(f"{len(sem_nada)} número(s) sem arquivo")
        self.app.var_status.set(" · ".join(partes) + ". Nada foi alterado ainda.")

    def _executar(self):
        mover = self.var_modo.get() == "mover"
        prontos = sum(1 for m in self.movimentos if m.situacao == "ok")
        verbo = "mover" if mover else "copiar"
        if not messagebox.askyesno(TITULO,
                                   f"Confirma {verbo} {prontos} arquivo(s) para\n{self.var_destino.get()}?"):
            return
        self.bt_executar.state(["disabled"])
        self.app.var_status.set("Movendo os arquivos…" if mover else "Copiando os arquivos…")
        self.em_fundo(lambda: nucleo.executar_separacao(self.movimentos, mover), self._terminou)

    def _terminou(self, resultado):
        feitos, erros = resultado
        self.app.var_status.set(f"{feitos} arquivo(s) no destino." +
                                (f" {len(erros)} com erro." if erros else ""))
        if erros:
            messagebox.showwarning(TITULO, "Alguns arquivos não foram:\n\n" + "\n".join(erros[:12]))
        self.movimentos = []
        self.tabela.limpar()


class AbaCfop(Aba):
    """Grava cópias dos XML com o CFOP trocado."""

    def __init__(self, pai, app):
        super().__init__(pai, app)
        self.var_origem = tk.StringVar()
        self.var_saida = tk.StringVar()
        self.var_notas = tk.StringVar()
        self.var_codigos = tk.StringVar()
        self.var_cfop = tk.StringVar(value="5102")
        self.var_sub = tk.BooleanVar(value=False)
        self.var_substituir = tk.BooleanVar(value=False)
        self.alteracoes: list = []

        self.columnconfigure(2, weight=1)

        ttk.Label(self, text="1", style="Passo.TLabel").grid(row=0, column=0, sticky="w", padx=(0, 8))
        ttk.Label(self, text="Pasta com os XMLs").grid(row=0, column=1, sticky="w", padx=(0, 12))
        linha = ttk.Frame(self)
        linha.grid(row=0, column=2, sticky="ew", pady=4)
        linha.columnconfigure(0, weight=1)
        self.ent_origem = ttk.Entry(linha, textvariable=self.var_origem)
        self.ent_origem.grid(row=0, column=0, sticky="ew")
        ttk.Button(linha, text="Procurar…",
                   command=lambda: self.escolher_pasta(self.var_origem, "Pasta com os XMLs")
                   ).grid(row=0, column=1, padx=(8, 0))
        ttk.Checkbutton(self, text="incluir subpastas", variable=self.var_sub).grid(row=1, column=2, sticky="w")

        ttk.Label(self, text="2", style="Passo.TLabel").grid(row=2, column=0, sticky="w", padx=(0, 8), pady=(8, 0))
        ttk.Label(self, text="O que trocar").grid(row=2, column=1, sticky="w", padx=(0, 12), pady=(8, 0))
        filtros = ttk.Frame(self)
        filtros.grid(row=2, column=2, sticky="ew", pady=(8, 0))
        filtros.columnconfigure(1, weight=1)
        filtros.columnconfigure(3, weight=1)
        ttk.Label(filtros, text="Notas").grid(row=0, column=0, sticky="w", padx=(0, 6))
        ttk.Entry(filtros, textvariable=self.var_notas).grid(row=0, column=1, sticky="ew")
        ttk.Label(filtros, text="Códigos de item").grid(row=0, column=2, sticky="w", padx=(16, 6))
        ttk.Entry(filtros, textvariable=self.var_codigos).grid(row=0, column=3, sticky="ew")
        ttk.Label(filtros, text="CFOP novo").grid(row=0, column=4, sticky="w", padx=(16, 6))
        ttk.Entry(filtros, textvariable=self.var_cfop, width=8).grid(row=0, column=5, sticky="w")
        ttk.Label(self, text="Deixe Notas ou Códigos em branco para valer para todas as notas / todos os itens.",
                  style="Apagado.TLabel").grid(row=3, column=2, sticky="w", pady=(2, 0))

        ttk.Label(self, text="3", style="Passo.TLabel").grid(row=4, column=0, sticky="w", padx=(0, 8), pady=(8, 0))
        ttk.Label(self, text="Gravar as cópias em").grid(row=4, column=1, sticky="w", padx=(0, 12), pady=(8, 0))
        linha2 = ttk.Frame(self)
        linha2.grid(row=4, column=2, sticky="ew", pady=(8, 0))
        linha2.columnconfigure(0, weight=1)
        self.ent_saida = ttk.Entry(linha2, textvariable=self.var_saida)
        self.ent_saida.grid(row=0, column=0, sticky="ew")
        ttk.Button(linha2, text="Procurar…",
                   command=lambda: self.escolher_pasta(self.var_saida, "Pasta de saída")
                   ).grid(row=0, column=1, padx=(8, 0))
        ttk.Checkbutton(self, text="substituir se já existir no destino",
                        variable=self.var_substituir).grid(row=5, column=2, sticky="w")

        acoes = ttk.Frame(self)
        acoes.grid(row=6, column=0, columnspan=3, sticky="ew", pady=(12, 6))
        ttk.Button(acoes, text="Conferir", command=self._conferir).pack(side="left")
        self.bt_gravar = ttk.Button(acoes, text="Gravar cópias", style="Primario.TButton",
                                    command=self._executar, state="disabled")
        self.bt_gravar.pack(side="left", padx=(8, 0))
        tk.Label(acoes, text="Mexer no XML derruba a assinatura digital: serve para conferência e "
                             "importação, não para reenviar à SEFAZ.",
                 bg=tema.FUNDO, fg=VERMELHO, font=(tema.FONTE, 9)).pack(side="left", padx=(20, 0))

        self.tabela = tema.Tabela(self, [
            ("arquivo", "Arquivo", 360, "w", True),
            ("nota", "Nota", 90, "w", False),
            ("item", "Código do item", 160, "w", False),
            ("de", "CFOP atual", 100, "center", False),
            ("para", "CFOP novo", 100, "center", False)])
        self.tabela.grid(row=7, column=0, columnspan=3, sticky="nsew")
        self.rowconfigure(7, weight=1)

        for var in (self.var_origem, self.var_saida, self.var_notas, self.var_codigos,
                    self.var_cfop, self.var_sub, self.var_substituir):
            var.trace_add("write", lambda *_: self._invalidar())

    def _invalidar(self):
        if self.alteracoes:
            self.app.var_status.set("Os campos mudaram: clique em Conferir de novo.")
        self.alteracoes = []
        self.tabela.limpar()
        self.bt_gravar.state(["disabled"])

    def _conferir(self):
        origem = Path(self.var_origem.get().strip())
        cfop = self.var_cfop.get().strip()
        if not origem.is_dir():
            messagebox.showwarning(TITULO, "Escolha a pasta onde estão os XMLs.")
            return
        if not (cfop.isdigit() and len(cfop) == 4):
            messagebox.showwarning(TITULO, "O CFOP novo tem que ter 4 números (ex.: 5102).")
            return
        notas = nucleo.separar_valores(self.var_notas.get())
        codigos = nucleo.separar_valores(self.var_codigos.get())
        subpastas = self.var_sub.get()

        self.bt_gravar.state(["disabled"])
        self.app.var_status.set("Lendo os XMLs…")
        self.em_fundo(lambda: nucleo.planejar_cfop(origem, notas, codigos, cfop, subpastas),
                      self._mostrar)

    def _mostrar(self, resultado):
        self.alteracoes, erros = resultado
        self.tabela.limpar()
        for a in self.alteracoes:
            self.tabela.inserir((a.arquivo.name, a.nota, a.codigo, a.cfop_atual, a.cfop_novo))
        arquivos = len({a.arquivo for a in self.alteracoes})
        if self.alteracoes:
            self.bt_gravar.state(["!disabled"])
        self.app.var_status.set(
            f"{len(self.alteracoes)} item(ns) em {arquivos} arquivo(s) para trocar." +
            (f" {len(erros)} XML não abriu." if erros else "") + " Nada foi gravado ainda.")
        if erros:
            messagebox.showwarning(TITULO, "Não consegui abrir:\n\n" + "\n".join(erros[:12]))

    def _executar(self):
        saida = Path(self.var_saida.get().strip())
        if not str(saida).strip() or str(saida) == ".":
            messagebox.showwarning(TITULO, "Escolha a pasta onde gravar as cópias.")
            return
        if nucleo.mesma_pasta(saida, Path(self.var_origem.get().strip())):
            messagebox.showwarning(TITULO, "A pasta de saída tem que ser diferente da de origem, "
                                                 "para o XML original ficar intacto.")
            return
        arquivos = len({a.arquivo for a in self.alteracoes})
        if not messagebox.askyesno(TITULO,
                                   f"Gravar {arquivos} XML com o CFOP trocado em\n{saida}?\n\n"
                                   "Os arquivos de origem não são alterados."):
            return
        substituir = self.var_substituir.get()
        self.bt_gravar.state(["disabled"])
        self.app.var_status.set("Gravando as cópias…")
        self.em_fundo(lambda: nucleo.executar_cfop(self.alteracoes, saida, substituir),
                      self._terminou)

    def _terminou(self, resultado):
        gravados, trocados, erros = resultado
        self.app.var_status.set(f"{gravados} arquivo(s) gravados, {trocados} item(ns) trocados." +
                                (f" {len(erros)} com problema." if erros else ""))
        if erros:
            messagebox.showwarning(TITULO, "Alguns arquivos não foram gravados:\n\n" +
                                   "\n".join(erros[:12]))
        self.alteracoes = []
        self.tabela.limpar()


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(TITULO)
        self.geometry("1080x700")
        self.minsize(900, 560)
        self.configure(bg=tema.FUNDO)
        self.var_status = tk.StringVar(value="Escolha a pasta e clique em Conferir. "
                                             "Nada muda no disco antes de você mandar.")
        tema.estilos(self)
        tema.faixa(self, TITULO, "Separar notas por número e trocar CFOP de itens.", f"v{VERSAO}")

        abas = ttk.Notebook(self, padding=(16, 10, 16, 0))
        abas.pack(fill="both", expand=True)
        abas.add(AbaSeparar(abas, self), text="Separar por número")
        abas.add(AbaCfop(abas, self), text="Trocar CFOP")

        base = ttk.Frame(self)
        base.pack(fill="x", side="bottom")
        self.progresso = ttk.Progressbar(base, mode="indeterminate", length=150)
        ttk.Label(base, textvariable=self.var_status, style="Status.TLabel", wraplength=980,
                  justify="left").pack(side="left", fill="x", expand=True)

    def ocupado(self, sim: bool):
        if sim:
            self.progresso.pack(side="right", padx=10, pady=4)
            self.progresso.start(12)
            self.configure(cursor="watch")
        else:
            self.progresso.stop()
            self.progresso.pack_forget()
            self.configure(cursor="")


if __name__ == "__main__":
    App().mainloop()
