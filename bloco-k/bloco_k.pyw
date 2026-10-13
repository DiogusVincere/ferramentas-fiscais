"""Bloco K: gera as planilhas de importação de inventário a partir da planilha de estoque.

Abra pelo hub (Ferramentas Fiscais) ou com dois cliques em "Abrir Bloco K.bat".
1. Escolha a empresa (cada uma tem seu leitor na pasta leitores/).
2. Escolha a planilha de estoque (ou várias, se a empresa manda mais de uma) e clique em
   "Ler planilha". Cada arquivo vira uma importação separada.
3. Confira o resultado e os avisos, depois clique em "Gerar planilhas".
"""
from __future__ import annotations

import datetime as dt
import json
import os
import queue
import sys
import threading
import traceback
from dataclasses import dataclass
from pathlib import Path

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

PASTA = Path(os.path.abspath(__file__)).parent  # abspath preserva a letra de unidade mapeada
RAIZ = PASTA.parent
sys.path.insert(0, str(PASTA))
sys.path.insert(0, str(RAIZ / "hub"))

import tema  # noqa: E402
import leitores  # noqa: E402
from blocok import VERSAO, nucleo, saidas  # noqa: E402
from blocok.nucleo import Aviso, ErroLeitura, formatar, moeda  # noqa: E402

TITULO = "Bloco K"
CONFIG = PASTA / "config.json"
PASTA_SAIDAS = PASTA / "saidas"
LOG_ERROS = PASTA / "erros.log"


def ler_config() -> dict:
    try:
        return json.loads(CONFIG.read_text(encoding="utf-8"))
    except Exception:
        return {}


def gravar_config(dados: dict) -> None:
    try:
        CONFIG.write_text(json.dumps(dados, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass


def registrar_erro(detalhe: str) -> None:
    try:
        with LOG_ERROS.open("a", encoding="utf-8") as f:
            f.write(f"\n==== {dt.datetime.now():%d/%m/%Y %H:%M:%S}\n{detalhe}\n")
    except Exception:
        pass


def nome_pasta(texto: str) -> str:
    return "".join(c for c in texto if c not in '<>:"/\\|?*').strip()[:70]


@dataclass
class Lote:
    """Um arquivo de estoque já lido. Cada lote vira uma importação: cadastro + modelo."""
    arquivo: Path
    leitura: nucleo.Leitura
    consolidados: list
    avisos: list

    @property
    def nome(self) -> str:
        return self.arquivo.name


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(TITULO)
        self.geometry("1180x760")
        self.minsize(940, 620)
        self.configure(bg=tema.FUNDO)
        self.config_usuario = ler_config()

        self.leitores, erros = leitores.carregar()
        self.por_rotulo = {leitores.rotulo(m): m for m in self.leitores}

        self.var_empresa = tk.StringVar()
        self.var_arquivo = tk.StringVar()
        self.var_lote = tk.StringVar()
        self.var_competencia = tk.StringVar()
        self.var_saida = tk.StringVar()
        self.var_resumo = tk.StringVar(value="Escolha a empresa e a planilha de estoque, depois clique em “Ler planilha”.")
        self.var_status = tk.StringVar(value=f"{len(self.leitores)} empresa(s) com leitor cadastrado.")

        self.lotes: list[Lote] = []   # resultado da última leitura, um por arquivo
        self.modulo_lido = None
        self.saida_escolhida = False  # True quando a pessoa escolheu ou digitou a pasta de saída
        self._definindo_saida = False
        self.ultima_pasta_gerada = None

        tema.estilos(self)
        self._montar()

        ultimo = self.config_usuario.get("empresa")
        if ultimo in self.por_rotulo:
            self.var_empresa.set(ultimo)
        elif self.leitores:
            self.var_empresa.set(leitores.rotulo(self.leitores[0]))
        self._definir_saida(str(PASTA_SAIDAS))
        self.var_empresa.trace_add("write", lambda *_: self._invalidar())
        self.var_arquivo.trace_add("write", lambda *_: self._invalidar())
        self.var_saida.trace_add("write", self._saida_editada)

        if erros:
            self.after(300, lambda: messagebox.showwarning(
                "Leitores com problema", "Alguns leitores não carregaram:\n\n" + "\n".join(erros), parent=self))
        if not self.leitores:
            self.after(300, lambda: messagebox.showerror(
                "Nenhum leitor", "Não há nenhum leitor na pasta leitores/.", parent=self))

    # ------------------------------------------------------------ montagem
    def _montar(self):
        tema.faixa(self, "Bloco K — Posse de Produtos",
                   "Da planilha de estoque às planilhas de importação do inventário.", f"v{VERSAO}")

        corpo = ttk.Frame(self, padding=(20, 14, 20, 0))
        corpo.pack(fill="both", expand=True)
        corpo.columnconfigure(2, weight=1)

        # Passo 1 e 2: empresa e arquivo(s)
        ttk.Label(corpo, text="1", style="Passo.TLabel").grid(row=0, column=0, sticky="w", padx=(0, 8))
        ttk.Label(corpo, text="Empresa").grid(row=0, column=1, sticky="w", padx=(0, 12))
        self.combo = ttk.Combobox(corpo, textvariable=self.var_empresa, state="readonly",
                                  values=list(self.por_rotulo), width=70)
        self.combo.grid(row=0, column=2, sticky="w", pady=4)

        ttk.Label(corpo, text="2", style="Passo.TLabel").grid(row=1, column=0, sticky="w", padx=(0, 8))
        ttk.Label(corpo, text="Planilha de estoque").grid(row=1, column=1, sticky="w", padx=(0, 12))
        linha_arq = ttk.Frame(corpo)
        linha_arq.grid(row=1, column=2, sticky="ew", pady=4)
        linha_arq.columnconfigure(0, weight=1)
        self.ent_arquivo = ttk.Entry(linha_arq, textvariable=self.var_arquivo)
        self.ent_arquivo.grid(row=0, column=0, sticky="ew")
        ttk.Button(linha_arq, text="Procurar…", command=self._procurar).grid(row=0, column=1, padx=(8, 0))
        self.bt_ler = ttk.Button(linha_arq, text="Ler planilha", style="Primario.TButton", command=self._ler)
        self.bt_ler.grid(row=0, column=2, padx=(8, 0))
        ttk.Label(corpo, text="Dá para escolher mais de um arquivo: cada um vira uma importação separada.",
                  style="Apagado.TLabel").grid(row=2, column=2, sticky="w")

        # Resumo, arquivo exibido e abas de conferência
        linha_resumo = ttk.Frame(corpo)
        linha_resumo.grid(row=3, column=0, columnspan=3, sticky="ew", pady=(12, 6))
        linha_resumo.columnconfigure(0, weight=1)
        ttk.Label(linha_resumo, textvariable=self.var_resumo, style="Resumo.TLabel", wraplength=780,
                  justify="left").grid(row=0, column=0, sticky="w")
        self.lbl_lote = ttk.Label(linha_resumo, text="Ver arquivo")
        self.combo_lote = ttk.Combobox(linha_resumo, textvariable=self.var_lote, state="readonly", width=46)
        self.combo_lote.bind("<<ComboboxSelected>>", lambda _e: self._exibir_lote(self.combo_lote.current()))

        self.abas = ttk.Notebook(corpo)
        self.abas.grid(row=4, column=0, columnspan=3, sticky="nsew")
        corpo.rowconfigure(4, weight=1)

        self.tab_modelo = tema.Tabela(self.abas, [
            ("codigo", "Código", 130, "w", False), ("descricao", "Descrição", 310, "w", True),
            ("un", "Un", 50, "center", False), ("ncm", "NCM", 85, "center", False),
            ("ind", "Ind.", 45, "center", False), ("qtd", "Quantidade", 95, "e", False),
            ("medio", "Valor médio", 100, "e", False), ("total", "Saldo (total)", 125, "e", False),
            ("linhas", "Linhas", 60, "center", False)])
        self.tab_linhas = tema.Tabela(self.abas, [
            ("linha", "Linha", 60, "center", False), ("codigo", "Código", 130, "w", False),
            ("descricao", "Descrição", 330, "w", True), ("un", "Un", 50, "center", False),
            ("ncm", "NCM", 85, "center", False), ("qtd", "Quantidade", 95, "e", False),
            ("vu", "Valor unit.", 100, "e", False), ("total", "Total", 115, "e", False)])
        self.tab_avisos = tema.Tabela(self.abas, [
            ("tipo", "Tipo", 110, "w", False), ("codigo", "Código", 150, "w", False),
            ("aviso", "Aviso (clique para ver inteiro)", 900, "w", True)])
        self.tab_avisos.arvore.bind("<<TreeviewSelect>>", self._mostrar_aviso)
        self.abas.add(self.tab_modelo, text="Posse de Produtos")
        self.abas.add(self.tab_linhas, text="Linhas lidas")
        self.abas.add(self.tab_avisos, text="Avisos")

        legenda = ttk.Frame(corpo)
        legenda.grid(row=5, column=0, columnspan=3, sticky="w", pady=(6, 0))
        tk.Label(legenda, bg=tema.VERMELHO_T, width=2, relief="solid", bd=1, highlightthickness=0).pack(
            side="left", padx=(0, 6))
        ttk.Label(legenda, text="conferir: veja a aba Avisos", style="Apagado.TLabel").pack(side="left", padx=(0, 24))
        ttk.Label(legenda, text="Saldo = valor total do código em estoque (soma das linhas).",
                  style="Apagado.TLabel").pack(side="left")

        # Passo 3: competência, pasta de saída e geração
        rodape = ttk.Frame(corpo, padding=(0, 14, 0, 14))
        rodape.grid(row=6, column=0, columnspan=3, sticky="ew")
        rodape.columnconfigure(4, weight=1)
        ttk.Label(rodape, text="3", style="Passo.TLabel").grid(row=0, column=0, sticky="w", padx=(0, 8))
        ttk.Label(rodape, text="Competência").grid(row=0, column=1, sticky="w", padx=(0, 8))
        ttk.Entry(rodape, textvariable=self.var_competencia, width=12).grid(row=0, column=2, sticky="w")
        ttk.Label(rodape, text="Salvar em").grid(row=0, column=3, sticky="e", padx=(24, 8))
        self.ent_saida = ttk.Entry(rodape, textvariable=self.var_saida)
        self.ent_saida.grid(row=0, column=4, sticky="ew")
        ttk.Button(rodape, text="Escolher…", command=self._escolher_saida).grid(row=0, column=5, padx=(8, 0))
        self.bt_abrir = ttk.Button(rodape, text="Abrir pasta", command=self._abrir_pasta, state="disabled")
        self.bt_abrir.grid(row=0, column=6, padx=(8, 0))
        self.bt_gerar = ttk.Button(rodape, text="Gerar planilhas", style="Primario.TButton",
                                   command=self._gerar, state="disabled")
        self.bt_gerar.grid(row=0, column=7, padx=(8, 0))
        ttk.Label(rodape, text="dd/mm/aaaa, último dia do mês", style="Apagado.TLabel").grid(
            row=1, column=2, columnspan=2, sticky="w", pady=(2, 0))

        # Barra de status
        base = ttk.Frame(self, style="TFrame")
        base.pack(fill="x", side="bottom")
        self.progresso = ttk.Progressbar(base, mode="indeterminate", length=160)
        ttk.Label(base, textvariable=self.var_status, style="Status.TLabel", wraplength=1100,
                  justify="left").pack(side="left", fill="x", expand=True)

        # Caminhos longos: mostrar o fim (nome do arquivo / última pasta)
        for var, entrada in ((self.var_arquivo, self.ent_arquivo), (self.var_saida, self.ent_saida)):
            var.trace_add("write", lambda *_, e=entrada: self.after_idle(lambda: e.xview_moveto(1)))

    # ------------------------------------------------------------ ações
    def _arquivos(self) -> list[Path]:
        """Arquivos escolhidos; vários ficam separados por ';' no campo."""
        return [Path(p.strip().strip('"')) for p in self.var_arquivo.get().split(";") if p.strip().strip('"')]

    def _procurar(self):
        inicial = self.config_usuario.get("pasta_entrada")
        if not inicial or not Path(inicial).exists():
            inicial = str(Path.home())
        caminhos = filedialog.askopenfilenames(
            parent=self, title="Planilha(s) de estoque", initialdir=inicial,
            filetypes=[("Planilhas", "*.xls *.xlsx *.xlsm *.xlsb *.csv"), ("Todos os arquivos", "*.*")])
        if caminhos:
            self.var_arquivo.set("; ".join(os.path.normpath(c) for c in caminhos))

    def _escolher_saida(self):
        atual = self.var_saida.get() or str(PASTA_SAIDAS)
        pasta = filedialog.askdirectory(parent=self, title="Onde salvar as planilhas",
                                        initialdir=atual if Path(atual).exists() else str(PASTA))
        if pasta:
            self.var_saida.set(os.path.normpath(pasta))  # o trace marca como escolhida

    def _mostrar_aviso(self, _evento=None):
        arvore = self.tab_avisos.arvore
        selecao = arvore.selection()
        if selecao:
            tipo, codigo, texto = arvore.item(selecao[0], "values")
            self.var_status.set(f"{tipo} · {codigo}: {texto}" if codigo else f"{tipo}: {texto}")

    def _saida_editada(self, *_):
        if not self._definindo_saida:
            self.saida_escolhida = True

    def _definir_saida(self, valor: str):
        """Troca a pasta de saída sem contar como escolha da pessoa."""
        self._definindo_saida = True
        self.var_saida.set(valor)
        self._definindo_saida = False

    def _abrir_pasta(self):
        if self.ultima_pasta_gerada and Path(self.ultima_pasta_gerada).exists():
            os.startfile(self.ultima_pasta_gerada)

    def _invalidar(self):
        """Empresa ou arquivo mudou: a leitura anterior não vale mais."""
        if self.lotes:
            self.lotes = []
            self.bt_gerar.configure(state="disabled")
            self.var_status.set("Empresa ou arquivo mudou. Clique em “Ler planilha” de novo.")

    def _ocupado(self, sim: bool, mensagem: str = ""):
        estado = "disabled" if sim else "normal"
        self.bt_ler.configure(state=estado)
        self.combo.configure(state="disabled" if sim else "readonly")
        if sim:
            self.bt_gerar.configure(state="disabled")
            self.var_status.set(mensagem)
            self.progresso.pack(side="right", padx=12)
            self.progresso.start(12)
            self.configure(cursor="watch")
        else:
            self.progresso.stop()
            self.progresso.pack_forget()
            self.configure(cursor="")

    def _em_segundo_plano(self, tarefa, ao_terminar, mensagem):
        self._ocupado(True, mensagem)
        fila: queue.Queue = queue.Queue()

        def trabalhar():
            try:
                fila.put(("ok", tarefa(), None))
            except Exception as e:  # devolve para a tela tratar
                fila.put(("erro", e, traceback.format_exc()))

        threading.Thread(target=trabalhar, daemon=True).start()

        def conferir_fila():
            try:
                situacao, resultado, detalhe = fila.get_nowait()
            except queue.Empty:
                self.after(120, conferir_fila)
                return
            self._ocupado(False)
            if situacao == "ok":
                ao_terminar(resultado)
            else:
                self._mostrar_erro("Não foi possível ler a planilha", resultado, detalhe)

        self.after(120, conferir_fila)

    def _mostrar_erro(self, titulo, erro, detalhe=None):
        if isinstance(erro, ErroLeitura):
            messagebox.showerror(titulo, str(erro), parent=self)
        else:
            registrar_erro(detalhe or repr(erro))
            messagebox.showerror(titulo, f"Erro inesperado: {erro}\n\nOs detalhes foram gravados em {LOG_ERROS.name}.",
                                 parent=self)
        self.var_status.set(titulo + ".")

    def _ler(self):
        modulo = self.por_rotulo.get(self.var_empresa.get())
        arquivos = self._arquivos()
        if not modulo:
            messagebox.showwarning("Falta a empresa", "Escolha a empresa.", parent=self)
            return
        if not arquivos:
            messagebox.showwarning("Falta a planilha", "Escolha a planilha de estoque.", parent=self)
            return
        faltando = [str(a) for a in arquivos if not a.is_file()]
        if faltando:
            messagebox.showwarning("Arquivo não encontrado", "Não encontrei:\n\n" + "\n".join(faltando), parent=self)
            return

        def tarefa():
            lotes = []
            for arquivo in arquivos:
                try:
                    leitura = modulo.ler(arquivo)
                except ErroLeitura as e:
                    raise ErroLeitura(f"{arquivo.name}:\n{e}") from e
                consolidados = nucleo.consolidar(leitura.itens, leitores.indicador(modulo))
                avisos = [Aviso("Leitura", "", a) for a in leitura.avisos] + nucleo.conferir(
                    leitura.itens, consolidados, conferir_valores=leitores.tem_valores(modulo),
                    conferir_descricao=leitores.tem_descricao(modulo),
                    conferir_total_calculado=leitores.total_conferivel(modulo))
                lotes.append(Lote(arquivo, leitura, consolidados, avisos))
            return modulo, lotes

        lento = any(a.suffix.lower() not in (".xlsx", ".xlsm", ".csv") for a in arquivos)
        mensagem = "Lendo pelo Excel… (arquivos .xls levam alguns segundos)" if lento else "Lendo…"
        self._em_segundo_plano(tarefa, self._mostrar_leitura, mensagem)

    def _mostrar_leitura(self, resultado):
        modulo, lotes = resultado
        self.modulo_lido, self.lotes = modulo, lotes
        self.config_usuario.update(empresa=leitores.rotulo(modulo), pasta_entrada=str(lotes[0].arquivo.parent))
        gravar_config(self.config_usuario)

        if len(lotes) > 1:
            self.combo_lote.configure(values=[f"{i}. {l.nome}" for i, l in enumerate(lotes, start=1)])
            self.lbl_lote.grid(row=0, column=1, padx=(16, 6))
            self.combo_lote.grid(row=0, column=2)
        else:
            self.lbl_lote.grid_remove()
            self.combo_lote.grid_remove()
        self._exibir_lote(0)

        competencias = [l.leitura.competencia for l in lotes if l.leitura.competencia]
        competencia = competencias[0] if competencias else None
        self.var_competencia.set(competencia.strftime("%d/%m/%Y") if competencia else "")
        if not self.saida_escolhida:
            self._definir_saida(str(self._pasta_padrao(modulo, competencia)))

        # Dois arquivos com o mesmo sufixo gerariam os mesmos nomes (ex.: o .xls e o .csv do mesmo estoque)
        sufixos = [l.leitura.sufixo for l in lotes]
        repetidos = [l.nome for l in lotes if sufixos.count(l.leitura.sufixo) > 1]
        pode_gerar = all(l.leitura.itens for l in lotes) and not repetidos
        self.bt_gerar.configure(state="normal" if pode_gerar else "disabled")

        total_avisos = sum(len(l.avisos) for l in lotes)
        if repetidos:
            self.var_status.set("Há arquivos que gerariam os mesmos nomes de saída. Escolha só um de cada.")
            messagebox.showwarning(
                "Arquivos repetidos",
                "Estes arquivos gerariam planilhas com o mesmo nome:\n\n" + "\n".join(repetidos) +
                "\n\nEscolha só um deles (por exemplo, o .xls ou o .csv do mesmo estoque) e leia de novo.", parent=self)
        elif not competencia:
            self.var_status.set("Não achei a competência no arquivo nem na pasta: digite no campo 3 antes de gerar.")
        elif total_avisos:
            self.var_status.set(f"Leitura concluída: {total_avisos} aviso(s) para conferir. "
                                "Confira e clique em “Gerar planilhas”.")
        else:
            self.var_status.set("Leitura concluída sem avisos. Clique em “Gerar planilhas”.")
        meses = {(c.year, c.month) for c in competencias}
        if len(meses) > 1:
            messagebox.showwarning(
                "Competências diferentes",
                "Os arquivos são de meses diferentes:\n\n" +
                "\n".join(f"{l.nome}: {l.leitura.competencia:%m/%Y}" for l in lotes if l.leitura.competencia) +
                "\n\nTodos serão gerados com a competência do campo 3.", parent=self)

    def _exibir_lote(self, indice: int):
        """Mostra nas abas o arquivo escolhido (quando há mais de um)."""
        if not self.lotes:
            return
        indice = max(0, min(indice, len(self.lotes) - 1))
        lote = self.lotes[indice]
        leitura, consolidados, avisos = lote.leitura, lote.consolidados, lote.avisos
        if len(self.lotes) > 1:
            self.combo_lote.current(indice)

        codigos_alerta = {a.codigo for a in avisos if a.tipo in ("Dados", "Divergência", "Código longo")}

        self.tab_modelo.limpar()
        for c in consolidados:
            tags = ("alerta",) if c.codigo in codigos_alerta else ()
            self.tab_modelo.inserir((c.codigo, c.descricao, c.unidade, c.ncm, c.indicador, formatar(c.quantidade, 4),
                                     formatar(c.valor_medio, 4), moeda(c.total), len(c.linhas)), tags)

        self.tab_linhas.limpar()
        for it in leitura.itens:
            tags = ("alerta",) if it.codigo in codigos_alerta else ()
            self.tab_linhas.inserir((it.linha or "", it.codigo, it.descricao, it.unidade, it.ncm,
                                     formatar(it.quantidade, 4), formatar(it.unitario, 4), moeda(it.total)), tags)

        ordem = {"Leitura": 0, "Coluna trocada": 1, "Dados": 2, "Divergência": 3, "Código longo": 4}
        self.tab_avisos.limpar()
        for a in sorted(avisos, key=lambda a: ordem.get(a.tipo, 9)):
            self.tab_avisos.inserir((a.tipo, a.codigo, a.texto), ("alerta",) if a.tipo == "Coluna trocada" else ())

        self.abas.tab(0, text=f"Posse de Produtos ({len(consolidados)})")
        self.abas.tab(1, text=f"Linhas lidas ({len(leitura.itens)})")
        self.abas.tab(2, text=f"Avisos ({len(avisos)})")

        partes = []
        if len(self.lotes) > 1:
            partes.append(f"Arquivo {indice + 1} de {len(self.lotes)}")
        partes += [f"Aba “{leitura.aba}”", f"{len(leitura.itens)} linhas", f"{len(consolidados)} códigos"]
        partes.append(f"Estoque {moeda(sum(c.total for c in consolidados))}" if leitores.tem_valores(self.modulo_lido)
                      else "Sem valores: Saldo zerado")
        partes.append(f"conta {leitura.conta or leitores.conta(self.modulo_lido)}")
        participantes = sorted({f"{c.participante} {c.cnpj_participante}".strip() for c in consolidados
                                if c.participante or c.cnpj_participante})
        if participantes:
            partes.append("participante " + ", ".join(participantes))
        if leitura.sufixo:
            partes.append(f"saída “…{leitura.sufixo}”")
        self.var_resumo.set("   ·   ".join(partes))

    def _pasta_padrao(self, modulo, competencia):
        pasta = PASTA_SAIDAS / nome_pasta(leitores.rotulo(modulo))
        return pasta / competencia.strftime("%Y-%m") if competencia else pasta

    def _competencia(self):
        texto = self.var_competencia.get().strip()
        try:
            data = dt.datetime.strptime(texto, "%d/%m/%Y").date()
        except ValueError:
            messagebox.showwarning("Competência", "Digite a competência no formato dd/mm/aaaa, ex.: 30/06/2026.",
                                   parent=self)
            return None
        fim_do_mes = nucleo.ultimo_dia(data.year, data.month)
        if data != fim_do_mes and messagebox.askyesno(
                "Competência", f"A competência costuma ser o último dia do mês.\n\nUsar {fim_do_mes:%d/%m/%Y}?",
                parent=self):
            data = fim_do_mes
            self.var_competencia.set(data.strftime("%d/%m/%Y"))
        lida = next((l.leitura.competencia for l in self.lotes if l.leitura.competencia), None)
        if lida and (lida.year, lida.month) != (data.year, data.month) and not messagebox.askyesno(
                "Competência", f"O arquivo lido é de {lida:%m/%Y}, mas a competência digitada é {data:%m/%Y}.\n\n"
                               "Gerar assim mesmo?", parent=self):
            return None
        return data

    def _gerar(self):
        if not self.lotes:
            return
        competencia = self._competencia()
        if competencia is None:
            return
        if not self.saida_escolhida:
            self._definir_saida(str(self._pasta_padrao(self.modulo_lido, competencia)))
        pasta = Path(self.var_saida.get().strip().strip('"') or PASTA_SAIDAS)

        planos = []  # (lote, destino do cadastro, destino do modelo, conta)
        for lote in self.lotes:
            nome_cad, nome_mod = saidas.nomes_arquivos(competencia, lote.leitura.sufixo)
            conta = lote.leitura.conta or leitores.conta(self.modulo_lido)
            planos.append((lote, pasta / nome_cad, pasta / nome_mod, conta))
        existentes = [d.name for _, cad, mod, _ in planos for d in (cad, mod) if d.exists()]
        if existentes and not messagebox.askyesno(
                "Arquivos já existem", "Já existe nesta pasta:\n\n" + "\n".join(existentes) + "\n\nSubstituir?",
                parent=self):
            return
        try:
            pasta.mkdir(parents=True, exist_ok=True)
            for lote, cad, mod, conta in planos:
                saidas.gerar_cadastro(lote.consolidados, conta, cad)
                saidas.gerar_modelo(lote.consolidados, competencia, conta, mod)
        except PermissionError:
            messagebox.showerror("Arquivo em uso", "Não consegui gravar: feche as planilhas no Excel e tente de novo.",
                                 parent=self)
            return
        except Exception as e:
            self._mostrar_erro("Não foi possível gerar as planilhas", e, traceback.format_exc())
            return

        self.ultima_pasta_gerada = str(pasta)
        self.bt_abrir.configure(state="normal")
        self.var_status.set(f"{2 * len(planos)} planilhas geradas em {pasta}")
        linhas = []
        for lote, cad, mod, conta in planos:
            codigos = len({c.codigo for c in lote.consolidados})
            linhas.append(f"{cad.name}\n   {codigos} produtos para cadastrar\n"
                          f"{mod.name}\n   {len(lote.consolidados)} registros de posse, conta {conta}")
        messagebox.showinfo("Planilhas geradas", "\n\n".join(linhas) +
                            f"\n\nCompetência {competencia:%d/%m/%Y}\nPasta: {pasta}", parent=self)


def principal():
    try:  # texto nítido em telas com zoom do Windows
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    App().mainloop()


if __name__ == "__main__":
    principal()
