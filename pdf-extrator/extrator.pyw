"""Extrator de PDF para Excel.

Abra pelo hub (Ferramentas Fiscais) ou com dois cliques em "Abrir Extrator.bat".

1. Escolha o PDF (dá para escolher vários de uma vez) e clique em "Extrair".
   No modo automático a ferramenta experimenta três jeitos de achar a tabela e fica
   com o que sair mais bem formado — e diz qual usou.
2. Confira a prévia.
3. Gere o Excel.

PDF que é foto (digitalizado) não tem texto nenhum: a ferramenta avisa em vez de
devolver planilha vazia.
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

import pdfx  # noqa: E402
import tema  # noqa: E402
from tema import APAGADO  # noqa: E402

VERSAO = "2.0"
PASTA_SAIDAS = PASTA / "saidas"
MODOS = [("Automático", "auto"), ("Tabela com grade", "tabelas"),
         ("Colunas pelo alinhamento", "colunas"), ("Uma linha por célula", "linhas")]
MAX_PREVIA = 400


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Extrator de PDF")
        self.geometry("1180x820")
        self.minsize(960, 640)
        self.configure(bg=tema.FUNDO)

        self.extracoes: list = []          # [(caminho, Extracao)]
        self.fila: queue.Queue = queue.Queue()
        self.trabalhando = False

        self.var_pdf = tk.StringVar()
        self.var_modo = tk.StringVar(value=MODOS[0][0])
        self.var_repetidos = tk.BooleanVar(value=True)
        self.var_cabecalho = tk.BooleanVar(value=True)
        self.var_qual = tk.StringVar()
        self.var_saida = tk.StringVar(value=str(PASTA_SAIDAS))
        self.var_status = tk.StringVar(value="Escolha o PDF e clique em “Extrair”.")

        tema.estilos(self)
        self._montar()

    # ------------------------------------------------------------ montagem
    def _montar(self):
        tema.faixa(self, "Extrator de PDF para Excel",
                   "Tira a tabela de dentro do PDF e devolve em planilha.", f"v{VERSAO}")

        corpo = ttk.Frame(self, padding=(20, 14, 20, 0))
        corpo.pack(fill="both", expand=True)
        corpo.columnconfigure(2, weight=1)

        ttk.Label(corpo, text="1", style="Passo.TLabel").grid(row=0, column=0, sticky="w", padx=(0, 8))
        ttk.Label(corpo, text="PDF").grid(row=0, column=1, sticky="w", padx=(0, 12))
        linha = ttk.Frame(corpo)
        linha.grid(row=0, column=2, sticky="ew", pady=4)
        linha.columnconfigure(0, weight=1)
        self.ent_pdf = ttk.Entry(linha, textvariable=self.var_pdf)
        self.ent_pdf.grid(row=0, column=0, sticky="ew")
        ttk.Button(linha, text="Procurar…", command=self._escolher).grid(row=0, column=1, padx=(8, 0))
        self.bt_extrair = ttk.Button(linha, text="Extrair", style="Primario.TButton", command=self._extrair)
        self.bt_extrair.grid(row=0, column=2, padx=(8, 0))
        ttk.Label(corpo, text="Dá para escolher mais de um PDF: cada um vira uma planilha.",
                  style="Apagado.TLabel").grid(row=1, column=2, sticky="w")

        opcoes = ttk.Frame(corpo)
        opcoes.grid(row=2, column=2, sticky="w", pady=(6, 0))
        ttk.Label(opcoes, text="Como achar a tabela").pack(side="left", padx=(0, 8))
        ttk.Combobox(opcoes, textvariable=self.var_modo, state="readonly", width=26,
                     values=[m[0] for m in MODOS]).pack(side="left")
        ttk.Checkbutton(opcoes, text="tirar cabeçalho repetido entre páginas",
                        variable=self.var_repetidos).pack(side="left", padx=(18, 0))

        # prévia
        ttk.Label(corpo, text="2", style="Passo.TLabel").grid(row=3, column=0, sticky="w",
                                                              padx=(0, 8), pady=(14, 0))
        ttk.Label(corpo, text="Prévia").grid(row=3, column=1, sticky="w", padx=(0, 12), pady=(14, 0))
        cabecalho = ttk.Frame(corpo)
        cabecalho.grid(row=3, column=2, sticky="ew", pady=(14, 0))
        cabecalho.columnconfigure(1, weight=1)
        self.combo_arquivo = ttk.Combobox(cabecalho, state="readonly", width=44)
        self.combo_arquivo.grid(row=0, column=0, sticky="w")
        self.combo_arquivo.bind("<<ComboboxSelected>>",
                                lambda _e: self._mostrar(self.combo_arquivo.current()))
        ttk.Label(cabecalho, textvariable=self.var_qual, style="Apagado.TLabel").grid(
            row=0, column=1, sticky="w", padx=(12, 0))
        ttk.Checkbutton(cabecalho, text="primeira linha é o cabeçalho",
                        variable=self.var_cabecalho).grid(row=0, column=2, sticky="e")

        self.quadro_tabela = ttk.Frame(corpo)
        self.quadro_tabela.grid(row=4, column=0, columnspan=3, sticky="nsew", pady=(6, 0))
        corpo.rowconfigure(4, weight=3)
        self.tabela = None

        self.tab_avisos = tema.Tabela(corpo, [("aviso", "Avisos", 1050, "w", True)])
        self.tab_avisos.grid(row=5, column=0, columnspan=3, sticky="nsew", pady=(8, 0))
        corpo.rowconfigure(5, weight=1)

        rodape = ttk.Frame(corpo, padding=(0, 12, 0, 12))
        rodape.grid(row=6, column=0, columnspan=3, sticky="ew")
        rodape.columnconfigure(3, weight=1)
        ttk.Label(rodape, text="3", style="Passo.TLabel").grid(row=0, column=0, sticky="w", padx=(0, 8))
        ttk.Label(rodape, text="Salvar em").grid(row=0, column=1, sticky="w", padx=(0, 12))
        ttk.Entry(rodape, textvariable=self.var_saida).grid(row=0, column=3, sticky="ew")
        ttk.Button(rodape, text="Escolher…", command=self._escolher_saida).grid(row=0, column=4, padx=(8, 0))
        self.bt_abrir = ttk.Button(rodape, text="Abrir pasta", command=self._abrir_pasta, state="disabled")
        self.bt_abrir.grid(row=0, column=5, padx=(8, 0))
        self.bt_gerar = ttk.Button(rodape, text="Gerar Excel", style="Primario.TButton",
                                   command=self._gerar, state="disabled")
        self.bt_gerar.grid(row=0, column=6, padx=(8, 0))

        base = ttk.Frame(self)
        base.pack(fill="x", side="bottom")
        self.progresso = ttk.Progressbar(base, mode="determinate", length=180)
        ttk.Label(base, textvariable=self.var_status, style="Status.TLabel", wraplength=1080,
                  justify="left").pack(side="left", fill="x", expand=True)

    # ------------------------------------------------------------ apoio
    def _arquivos(self) -> list:
        bruto = self.var_pdf.get().strip()
        if not bruto:
            return []
        return [Path(p.strip()) for p in bruto.split("; ") if p.strip()]

    def _escolher(self):
        escolhidos = filedialog.askopenfilenames(
            title="PDF para extrair", filetypes=[("PDF", "*.pdf"), ("Todos", "*.*")],
            initialdir=str(RAIZ))
        if escolhidos:
            self.var_pdf.set("; ".join(os.path.normpath(c) for c in escolhidos))

    def _escolher_saida(self):
        pasta = filedialog.askdirectory(title="Onde salvar", initialdir=self.var_saida.get().strip())
        if pasta:
            self.var_saida.set(os.path.normpath(pasta))

    def _abrir_pasta(self):
        destino = Path(self.var_saida.get().strip())
        if destino.exists():
            os.startfile(str(destino))

    # ------------------------------------------------------------ passo 1
    def _extrair(self):
        if self.trabalhando:
            return
        arquivos = self._arquivos()
        if not arquivos or not all(a.is_file() for a in arquivos):
            messagebox.showwarning("Extrator", "Escolha ao menos um PDF que exista.")
            return
        modo = dict(MODOS)[self.var_modo.get()]
        tirar = self.var_repetidos.get()

        self.extracoes = []
        self.trabalhando = True
        self.bt_extrair.state(["disabled"])
        self.bt_gerar.state(["disabled"])
        self.progresso.pack(side="right", padx=10, pady=4)
        self.progresso.configure(maximum=len(arquivos), value=0)
        self.configure(cursor="watch")
        self.var_status.set(f"Lendo {len(arquivos)} PDF(s)…")

        def trabalhar():
            try:
                for i, caminho in enumerate(arquivos, start=1):
                    self.fila.put(("um", (i, caminho, pdfx.extrair(caminho, modo, tirar))))
                self.fila.put(("fim", None))
            except Exception:
                self.fila.put(("erro", traceback.format_exc()))

        threading.Thread(target=trabalhar, daemon=True).start()
        self.after(80, self._ler_fila)

    def _ler_fila(self):
        try:
            while True:
                tipo, dados = self.fila.get_nowait()
                if tipo == "um":
                    i, caminho, extracao = dados
                    self.extracoes.append((caminho, extracao))
                    self.progresso.configure(value=i)
                    self.var_status.set(f"[{i}] {caminho.name}: {len(extracao.linhas)} linha(s).")
                elif tipo == "erro":
                    self._parar()
                    messagebox.showerror("Extrator", "Deu erro:\n\n" + dados.strip().splitlines()[-1])
                    return
                else:
                    self._parar()
                    self._depois()
                    return
        except queue.Empty:
            pass
        self.after(80, self._ler_fila)

    def _parar(self):
        self.trabalhando = False
        self.progresso.pack_forget()
        self.configure(cursor="")
        self.bt_extrair.state(["!disabled"])

    def _depois(self):
        rotulos = [f"{c.name}  ({len(e.linhas)} linhas)" for c, e in self.extracoes]
        self.combo_arquivo.configure(values=rotulos)
        if rotulos:
            self.combo_arquivo.current(0)
            self._mostrar(0)
        com_dados = sum(1 for _c, e in self.extracoes if e.linhas)
        if com_dados:
            self.bt_gerar.state(["!disabled"])
        self.var_status.set(
            f"{len(self.extracoes)} PDF(s) lidos, {com_dados} com tabela. "
            + ("Confira a prévia e gere o Excel." if com_dados else "Veja os avisos."))

    # ------------------------------------------------------------ passo 2
    def _mostrar(self, indice: int):
        if not (0 <= indice < len(self.extracoes)):
            return
        caminho, extracao = self.extracoes[indice]
        if self.tabela is not None:
            self.tabela.destroy()
        colunas = [(f"c{i}", f"Coluna {i + 1}", 150, "w", True) for i in range(max(extracao.colunas, 1))]
        self.tabela = tema.Tabela(self.quadro_tabela, colunas)
        self.tabela.pack(fill="both", expand=True)
        for linha in extracao.linhas[:MAX_PREVIA]:
            self.tabela.inserir(tuple(linha))

        nome = dict((v, k) for k, v in MODOS).get(extracao.modo, extracao.modo)
        self.var_qual.set(f"{len(extracao.linhas)} linhas × {extracao.colunas} colunas · "
                          f"modo “{nome}” · {extracao.paginas} página(s)"
                          + (f" · prévia das primeiras {MAX_PREVIA}" if len(extracao.linhas) > MAX_PREVIA else ""))
        self.tab_avisos.limpar()
        for aviso in extracao.avisos:
            self.tab_avisos.inserir((aviso,), () if extracao.linhas else ("alerta",))

    # ------------------------------------------------------------ passo 3
    def _gerar(self):
        destino = Path(self.var_saida.get().strip() or str(PASTA_SAIDAS))
        cabecalho = self.var_cabecalho.get()
        feitos, erros = [], []
        for caminho, extracao in self.extracoes:
            if not extracao.linhas:
                continue
            alvo = destino / (caminho.stem + ".xlsx")
            try:
                pdfx.gerar_excel(extracao, alvo, cabecalho)
                feitos.append(alvo.name)
            except PermissionError:
                erros.append(f"{alvo.name} está aberto no Excel.")
            except Exception as erro:
                erros.append(f"{alvo.name}: {erro}")
        if feitos:
            self.bt_abrir.state(["!disabled"])
        self.var_status.set(f"{len(feitos)} planilha(s) gerada(s) em {destino}."
                            + (f" {len(erros)} com erro." if erros else ""))
        if erros:
            messagebox.showwarning("Extrator", "\n".join(erros[:10]))
        elif feitos:
            messagebox.showinfo("Extrator", "Pronto:\n" + "\n".join(feitos[:10]))


if __name__ == "__main__":
    App().mainloop()
