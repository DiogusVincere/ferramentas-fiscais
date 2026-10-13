"""Comparador de Notas Fiscais.

Abra pelo hub (Ferramentas Fiscais) ou com dois cliques em "Abrir Comparador.bat".

1. Aponte a pasta onde estão os arquivos e clique em "Ler a pasta".
   A ferramenta abre cada planilha, descobre sozinha onde a tabela começa e o que cada
   coluna é — olhando o conteúdo, não o título — e diz em português o que entendeu.
2. Confira o par que ela sugeriu (dá para trocar) e clique em "Comparar".
3. Veja o que precisa de olho e gere o Excel.
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
from nf import VERSAO  # noqa: E402
from nf import saida  # noqa: E402
from nf.comparar import comparar, mesmo_relatorio, reais, sugerir_par  # noqa: E402
from nf.entender import entender_arquivo  # noqa: E402
from nf.planilha import arquivos_da_pasta  # noqa: E402
from tema import AMBAR, APAGADO, PAINEL, TINTA, VERDE, VERMELHO  # noqa: E402

PASTA_SAIDAS = PASTA / "saidas"


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Comparador de Notas")
        self.geometry("1240x860")
        self.minsize(1040, 680)
        self.configure(bg=tema.FUNDO)

        self.documentos: list = []
        self.resultado = None
        self.fila: queue.Queue = queue.Queue()
        self.trabalhando = False

        self.var_pasta = tk.StringVar()
        self.var_sub = tk.BooleanVar(value=False)
        self.var_a = tk.StringVar()
        self.var_b = tk.StringVar()
        self.var_tolerancia = tk.StringVar(value="0,01")
        self.var_saida = tk.StringVar(value=str(PASTA_SAIDAS))
        self.var_status = tk.StringVar(value="Aponte a pasta com os arquivos e clique em “Ler a pasta”.")
        self.var_explicacao = tk.StringVar()

        tema.estilos(self)
        self._montar()

    # ------------------------------------------------------------ montagem
    def _montar(self):
        tema.faixa(self, "Comparador de Notas Fiscais",
                   "Aponte a pasta: eu abro cada planilha, entendo as colunas e comparo.", f"v{VERSAO}")

        corpo = ttk.Frame(self, padding=(20, 14, 20, 0))
        corpo.pack(fill="both", expand=True)
        corpo.columnconfigure(2, weight=1)

        # passo 1: pasta
        ttk.Label(corpo, text="1", style="Passo.TLabel").grid(row=0, column=0, sticky="w", padx=(0, 8))
        ttk.Label(corpo, text="Pasta com os arquivos").grid(row=0, column=1, sticky="w", padx=(0, 12))
        linha = ttk.Frame(corpo)
        linha.grid(row=0, column=2, sticky="ew", pady=4)
        linha.columnconfigure(0, weight=1)
        self.ent_pasta = ttk.Entry(linha, textvariable=self.var_pasta)
        self.ent_pasta.grid(row=0, column=0, sticky="ew")
        ttk.Button(linha, text="Procurar…", command=self._escolher_pasta).grid(row=0, column=1, padx=(8, 0))
        self.bt_ler = ttk.Button(linha, text="Ler a pasta", style="Primario.TButton", command=self._ler)
        self.bt_ler.grid(row=0, column=2, padx=(8, 0))
        ttk.Checkbutton(corpo, text="incluir subpastas", variable=self.var_sub).grid(
            row=1, column=2, sticky="w")
        ttk.Label(corpo, text="Lê .xlsx, .xlsm, .xls e .csv. Não precisa do Excel aberto.",
                  style="Apagado.TLabel").grid(row=2, column=2, sticky="w")

        # o que foi entendido de cada arquivo
        self.tab_arquivos = tema.Tabela(corpo, [
            ("arquivo", "Arquivo", 250, "w", False),
            ("entendi", "O que eu entendi", 560, "w", True),
            ("notas", "Notas", 70, "e", False),
            ("comp", "Competência", 90, "center", False),
            ("soma", "Soma", 130, "e", False)])
        self.tab_arquivos.grid(row=3, column=0, columnspan=3, sticky="nsew", pady=(10, 0))
        corpo.rowconfigure(3, weight=1)

        # passo 2: par e comparação
        ttk.Label(corpo, text="2", style="Passo.TLabel").grid(row=4, column=0, sticky="w", padx=(0, 8), pady=(12, 0))
        ttk.Label(corpo, text="Comparar").grid(row=4, column=1, sticky="w", padx=(0, 12), pady=(12, 0))
        par = ttk.Frame(corpo)
        par.grid(row=4, column=2, sticky="ew", pady=(12, 0))
        par.columnconfigure(0, weight=1)
        par.columnconfigure(2, weight=1)
        self.combo_a = ttk.Combobox(par, textvariable=self.var_a, state="readonly")
        self.combo_a.grid(row=0, column=0, sticky="ew")
        ttk.Label(par, text="com").grid(row=0, column=1, padx=8)
        self.combo_b = ttk.Combobox(par, textvariable=self.var_b, state="readonly")
        self.combo_b.grid(row=0, column=2, sticky="ew")
        ttk.Label(par, text="Diferença a ignorar R$").grid(row=0, column=3, padx=(16, 6))
        ttk.Entry(par, textvariable=self.var_tolerancia, width=7).grid(row=0, column=4)
        self.bt_comparar = ttk.Button(par, text="Comparar", style="Primario.TButton",
                                      command=self._comparar, state="disabled")
        self.bt_comparar.grid(row=0, column=5, padx=(12, 0))

        # placar
        self.placar = ttk.Frame(corpo)
        self.placar.grid(row=5, column=0, columnspan=3, sticky="ew", pady=(12, 4))
        self.cartoes = {}
        for i, (chave, rotulo) in enumerate([("casadas", "notas casadas"), ("confere", "valor confere"),
                                             ("divergentes", "valor diferente"), ("so_a", "só no arquivo 1"),
                                             ("so_b", "só no arquivo 2"), ("icms_difere", "ICMS diferente")]):
            quadro = tk.Frame(self.placar, bg=PAINEL, highlightbackground=tema.LINHA, highlightthickness=1)
            quadro.grid(row=0, column=i, sticky="ew", padx=(0, 8))
            self.placar.columnconfigure(i, weight=1)
            valor = tk.Label(quadro, text="—", bg=PAINEL, fg=TINTA, font=(tema.FONTE, 16, "bold"))
            valor.pack(anchor="w", padx=12, pady=(8, 0))
            tk.Label(quadro, text=rotulo, bg=PAINEL, fg=APAGADO, font=(tema.FONTE, 9)).pack(
                anchor="w", padx=12, pady=(0, 8))
            self.cartoes[chave] = valor

        ttk.Label(corpo, textvariable=self.var_explicacao, style="Apagado.TLabel",
                  wraplength=1150, justify="left").grid(row=6, column=0, columnspan=3, sticky="w")

        # resultado
        self.abas = ttk.Notebook(corpo)
        self.abas.grid(row=7, column=0, columnspan=3, sticky="nsew", pady=(8, 0))
        corpo.rowconfigure(7, weight=2)
        colunas_nota = [
            ("situacao", "Situação", 130, "w", False),
            ("nota", "Nota", 90, "w", False),
            ("participante", "Fornecedor / cliente", 230, "w", True),
            ("valor_a", "Valor arquivo 1", 115, "e", False),
            ("valor_b", "Valor arquivo 2", 115, "e", False),
            ("dif", "Diferença", 100, "e", False),
            ("icms_a", "ICMS 1", 95, "e", False),
            ("icms_b", "ICMS 2", 95, "e", False),
            ("motivo", "O que eu vi", 330, "w", True)]
        self.tab_olho = tema.Tabela(self.abas, colunas_nota)
        self.tab_todas = tema.Tabela(self.abas, colunas_nota)
        self.tab_avisos = tema.Tabela(self.abas, [("aviso", "Aviso", 1100, "w", True)])
        self.abas.add(self.tab_olho, text="Precisa de olho")
        self.abas.add(self.tab_todas, text="Todas as notas")
        self.abas.add(self.tab_avisos, text="Avisos")

        # passo 3: saída
        rodape = ttk.Frame(corpo, padding=(0, 12, 0, 12))
        rodape.grid(row=8, column=0, columnspan=3, sticky="ew")
        rodape.columnconfigure(3, weight=1)
        ttk.Label(rodape, text="3", style="Passo.TLabel").grid(row=0, column=0, sticky="w", padx=(0, 8))
        ttk.Label(rodape, text="Salvar o Excel em").grid(row=0, column=1, sticky="w", padx=(0, 12))
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
        ttk.Label(base, textvariable=self.var_status, style="Status.TLabel", wraplength=1140,
                  justify="left").pack(side="left", fill="x", expand=True)

    # ------------------------------------------------------------ apoio
    def _ocupado(self, sim: bool, maximo: int = 0):
        self.trabalhando = sim
        if sim:
            self.progresso.pack(side="right", padx=10, pady=4)
            self.progresso.configure(maximum=max(maximo, 1), value=0)
            self.configure(cursor="watch")
        else:
            self.progresso.pack_forget()
            self.configure(cursor="")

    def _escolher_pasta(self):
        pasta = filedialog.askdirectory(title="Pasta com os arquivos de notas",
                                        initialdir=self.var_pasta.get().strip() or str(RAIZ))
        if pasta:
            self.var_pasta.set(os.path.normpath(pasta))

    def _escolher_saida(self):
        pasta = filedialog.askdirectory(title="Onde salvar o Excel",
                                        initialdir=self.var_saida.get().strip() or str(PASTA_SAIDAS))
        if pasta:
            self.var_saida.set(os.path.normpath(pasta))

    def _abrir_pasta(self):
        destino = Path(self.var_saida.get().strip())
        if destino.exists():
            os.startfile(str(destino))

    def _rotulo(self, doc) -> str:
        return f"{doc.nome}  ({len(doc.tabela.notas)} notas, {doc.competencia_texto})"

    # ------------------------------------------------------------ passo 1
    def _ler(self):
        if self.trabalhando:
            return
        pasta = Path(self.var_pasta.get().strip())
        if not pasta.is_dir():
            messagebox.showwarning("Comparador", "Escolha a pasta onde estão os arquivos.")
            return
        incluir = self.var_sub.get()
        arquivos = arquivos_da_pasta(pasta, incluir)
        if not arquivos:
            messagebox.showinfo("Comparador", f"Não achei planilha nenhuma em:\n{pasta}")
            return

        self.tab_arquivos.limpar()
        self.documentos = []
        self._ocupado(True, len(arquivos))
        self.bt_comparar.state(["disabled"])
        self.var_status.set(f"Lendo {len(arquivos)} arquivo(s)…")

        def trabalhar():
            try:
                for i, caminho in enumerate(arquivos, start=1):
                    self.fila.put(("lido", (i, caminho.name, entender_arquivo(caminho))))
                self.fila.put(("fim", None))
            except Exception:
                self.fila.put(("erro", traceback.format_exc()))

        threading.Thread(target=trabalhar, daemon=True).start()
        self.after(60, self._ler_fila)

    def _ler_fila(self):
        try:
            while True:
                tipo, dados = self.fila.get_nowait()
                if tipo == "lido":
                    i, nome, doc = dados
                    self.documentos.append(doc)
                    self.progresso.configure(value=i)
                    self.var_status.set(f"[{i}] {nome}: {doc.resumo()[:110]}")
                    soma = (reais(sum(n.valor or 0 for n in doc.tabela.notas)) if doc.ok else "")
                    self.tab_arquivos.inserir(
                        (nome, doc.resumo(), len(doc.tabela.notas) if doc.ok else "",
                         doc.competencia_texto if doc.ok else "", soma),
                        () if doc.ok else ("alerta",))
                elif tipo == "erro":
                    self._ocupado(False)
                    messagebox.showerror("Comparador", "Deu erro ao ler a pasta:\n\n" +
                                         dados.strip().splitlines()[-1])
                    return
                else:
                    self._ocupado(False)
                    self._depois_de_ler()
                    return
        except queue.Empty:
            pass
        self.after(60, self._ler_fila)

    def _depois_de_ler(self):
        bons = [d for d in self.documentos if d.ok]
        rotulos = [self._rotulo(d) for d in bons]
        self.combo_a.configure(values=rotulos)
        self.combo_b.configure(values=rotulos)
        if len(bons) < 2:
            self.var_status.set(
                f"{len(self.documentos)} arquivo(s) lidos, mas só {len(bons)} com tabela de notas. "
                f"Preciso de dois para comparar.")
            return
        a, b = sugerir_par(bons)
        self.var_a.set(self._rotulo(a))
        self.var_b.set(self._rotulo(b))
        self.bt_comparar.state(["!disabled"])
        self.var_status.set(
            f"{len(bons)} arquivo(s) com notas. Sugeri comparar {a.nome} com {b.nome} — "
            f"troque nas caixas se não for isso e clique em Comparar.")

    # ------------------------------------------------------------ passo 2
    def _documento(self, rotulo: str):
        return next((d for d in self.documentos if d.ok and self._rotulo(d) == rotulo), None)

    def _comparar(self):
        a = self._documento(self.var_a.get())
        b = self._documento(self.var_b.get())
        if a is None or b is None:
            messagebox.showwarning("Comparador", "Escolha os dois arquivos para comparar.")
            return
        if a is b:
            messagebox.showwarning("Comparador", "Escolha dois arquivos diferentes.")
            return
        try:
            tolerancia = float(self.var_tolerancia.get().replace(",", "."))
        except ValueError:
            messagebox.showwarning("Comparador", "A diferença a ignorar tem que ser um número, ex.: 0,01.")
            return
        if mesmo_relatorio(a, b) and not messagebox.askyesno(
                "Comparador",
                f"{a.nome} e {b.nome} parecem ser o mesmo relatório: mesma quantidade de notas, "
                f"mesma soma e os mesmos números.\n\nComparar os dois não vai apontar nada. "
                f"Quer comparar assim mesmo?"):
            return

        try:
            self.resultado = comparar(a, b, tolerancia=tolerancia)
        except Exception as erro:
            messagebox.showerror("Comparador", f"Não consegui comparar:\n\n{erro}")
            return
        self._mostrar()

    def _mostrar(self):
        res = self.resultado
        t = res.totais
        for chave, rotulo in self.cartoes.items():
            valor = t.get(chave, 0)
            if chave == "icms_difere" and not res.comparou_icms:
                rotulo.configure(text="—", fg=APAGADO)
            else:
                cor = TINTA
                if chave == "confere" and valor:
                    cor = VERDE
                elif chave in ("divergentes",) and valor:
                    cor = VERMELHO
                elif chave in ("so_a", "so_b", "icms_difere") and valor:
                    cor = AMBAR
                rotulo.configure(text=f"{valor}", fg=cor)

        self.var_explicacao.set(
            f"{res.explicacao_chave}.  Soma do arquivo 1: {reais(t['soma_a'])} · "
            f"soma do arquivo 2: {reais(t['soma_b'])} · diferença: {reais(t['diferenca'])}.")

        precisa = [l for l in res.linhas if l.situacao != "ok" or l.icms_difere or l.cfop_difere]
        self._preencher(self.tab_olho, precisa)
        self._preencher(self.tab_todas, res.linhas)
        self.tab_avisos.limpar()
        for aviso in res.avisos:
            self.tab_avisos.inserir((aviso,), ("alerta",))
        for doc in (res.doc_a, res.doc_b):
            for aviso in doc.tabela.avisos:
                self.tab_avisos.inserir((f"{doc.nome}: {aviso}",))
        self.abas.select(self.tab_olho if precisa else self.tab_todas)

        self.bt_gerar.state(["!disabled"])
        self.var_status.set(
            f"{t['casadas']} nota(s) casadas: valor confere em {t['confere']} e difere em "
            f"{t['divergentes']}. {t['so_a']} só no arquivo 1 e {t['so_b']} só no arquivo 2.")

    def _preencher(self, tabela, linhas):
        tabela.limpar()
        for l in linhas:
            tags = ()
            if l.situacao == "valor":
                tags = ("alerta",)
            elif l.situacao == "ok" and not (l.icms_difere or l.cfop_difere):
                tags = ("bom",)
            tabela.inserir((
                l.situacao_texto, l.numero, l.participante[:40],
                reais(l.valor_a) if l.valor_a is not None else "",
                reais(l.valor_b) if l.valor_b is not None else "",
                reais(l.diferenca_valor) if l.diferenca_valor else "",
                reais(l.icms_a) if l.icms_a is not None else "",
                reais(l.icms_b) if l.icms_b is not None else "",
                "; ".join(l.motivos + ([l.parecida_com] if l.parecida_com else []))), tags)

    # ------------------------------------------------------------ passo 3
    def _gerar(self):
        if self.resultado is None:
            return
        destino = Path(self.var_saida.get().strip() or str(PASTA_SAIDAS))
        arquivo = destino / saida.nome_sugerido(self.resultado)
        try:
            saida.gerar(self.resultado, arquivo)
        except PermissionError:
            messagebox.showerror("Comparador",
                                 f"{arquivo.name} está aberto no Excel. Feche e clique de novo.")
            return
        except Exception as erro:
            messagebox.showerror("Comparador", f"Não consegui gravar o Excel:\n\n{erro}")
            return
        self.bt_abrir.state(["!disabled"])
        self.var_status.set(f"Excel gerado: {arquivo}")
        messagebox.showinfo("Comparador", f"Pronto:\n{arquivo}")


if __name__ == "__main__":
    App().mainloop()
