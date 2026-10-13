"""Um leitor por empresa.

Cada arquivo .py desta pasta vira uma opção na lista "Empresa" da tela. Arquivos que
começam com "_" são ignorados (ex.: _modelo.py, que serve de base para uma empresa nova).
O leitor exemplo.py mostra um leitor completo para um layout simples.

Todo leitor precisa ter:
    EMPRESA          nome que aparece na tela
    ler(caminho)     função que devolve um blocok.nucleo.Leitura
E pode ter:
    CODIGO                   código interno da empresa (aparece antes do nome)
    CONTA_CONTABIL           conta contábil do estoque no sistema contábil (padrão "54")
    INDICADOR_PROPRIEDADE    0, 1 ou 2 (padrão 0 = do informante em seu poder)
    SEM_VALORES              True se a planilha da empresa não tem valores (Saldo sai zerado
                             e a tela não avisa valor zerado item a item)
    SEM_DESCRICAO            True se a planilha não tem descrição (cadastro sai em branco
                             e a tela não avisa item a item)
    SEM_CONFERENCIA_TOTAL    True quando o total informado não é quantidade × unitário
                             (custo médio, por exemplo), para a tela não avisar divergência
                             em toda linha

Na tela dá para escolher vários arquivos de uma vez: cada um é lido com ler() e vira uma
importação separada. O leitor separa os nomes de saída com Leitura.sufixo.
"""
from __future__ import annotations

import importlib
import pkgutil


def carregar():
    """Devolve (lista de módulos de leitor, lista de erros ao carregar)."""
    leitores, erros = [], []
    for info in pkgutil.iter_modules(__path__):
        if info.name.startswith("_"):
            continue
        try:
            modulo = importlib.import_module(f"{__name__}.{info.name}")
        except Exception as e:
            erros.append(f"leitores/{info.name}.py não carregou: {e}")
            continue
        faltando = [nome for nome in ("EMPRESA", "ler") if not hasattr(modulo, nome)]
        if faltando:
            erros.append(f"leitores/{info.name}.py: falta {', '.join(faltando)}.")
            continue
        leitores.append(modulo)
    leitores.sort(key=rotulo)
    return leitores, erros


def rotulo(modulo) -> str:
    codigo = str(getattr(modulo, "CODIGO", "") or "").strip()
    return f"{codigo} - {modulo.EMPRESA}" if codigo else modulo.EMPRESA


def conta(modulo) -> str:
    return str(getattr(modulo, "CONTA_CONTABIL", "54"))


def indicador(modulo) -> int:
    return int(getattr(modulo, "INDICADOR_PROPRIEDADE", 0))


def tem_valores(modulo) -> bool:
    return not getattr(modulo, "SEM_VALORES", False)


def tem_descricao(modulo) -> bool:
    return not getattr(modulo, "SEM_DESCRICAO", False)


def total_conferivel(modulo) -> bool:
    return not getattr(modulo, "SEM_CONFERENCIA_TOTAL", False)
