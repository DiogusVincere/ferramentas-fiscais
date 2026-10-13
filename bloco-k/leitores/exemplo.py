"""Leitor de exemplo: Empresa Exemplo Ltda (código EX1).

Serve de referência para escrever o leitor de uma empresa real e é o que a planilha
exemplos/estoque-exemplo.xlsx (gerada por exemplos/gerar_planilha_exemplo.py) usa.

Layout esperado, na primeira aba, com os títulos em qualquer linha das 15 primeiras:
    Código | Descrição | Unidade | NCM | Quantidade | Valor unitário | Valor total
A coluna "Valor total" é opcional: sem ela, o total é quantidade × valor unitário.
Linhas sem código (em branco, totais, observações) são ignoradas.

Competência: pela pasta do mês em que o arquivo está (ex.: ...\\2026\\07-2026\\arquivo.xlsx),
depois pelo nome do arquivo (ex.: "estoque 07-2026.xlsx") e, por último, por uma data escrita
nas linhas acima dos títulos (ex.: "Inventário em 31/07/2026"). Quando as fontes discordam,
sai um aviso; quando nenhuma traz o mês, a tela pede para digitar.
"""
from blocok.nucleo import (ItemLido, Leitura, achar_cabecalho, celula, competencia_das_pastas,
                           competencia_do_nome, competencia_do_texto, escolher_competencia, ncm,
                           numero, texto)
from blocok.planilha import ler_aba

EMPRESA = "EMPRESA EXEMPLO LTDA"
CODIGO = "EX1"
CONTA_CONTABIL = "54"
INDICADOR_PROPRIEDADE = 0  # 0 = estoque do informante em seu poder

TITULOS = {
    "codigo": ("Código", "Cod", "Cód."),
    "descricao": ("Descrição", "Produto"),
    "unidade": ("Unidade", "Un", "UM"),
    "ncm": "NCM",
    "quantidade": ("Quantidade", "Qtd", "Qtde"),
    "valor_unitario": ("Valor unitário", "Vlr unit", "Preço"),
}
OPCIONAIS = {"valor_total": ("Valor total", "Total")}


def ler(caminho) -> Leitura:
    aba = ler_aba(caminho, 0)
    linha_titulos, col = achar_cabecalho(aba.linhas, TITULOS, OPCIONAIS)

    itens, avisos = [], []
    for i in range(linha_titulos + 1, len(aba.linhas)):
        linha = aba.linhas[i]
        codigo = texto(celula(linha, col["codigo"]))
        if not codigo:
            continue
        n = i + 1
        quantidade = numero(celula(linha, col["quantidade"]))
        unitario = numero(celula(linha, col["valor_unitario"]))
        total = numero(celula(linha, col.get("valor_total")))
        if quantidade is None:
            avisos.append(f"Linha {n} ({codigo}): quantidade não é número. Linha ignorada.")
            continue
        if unitario is None and total is None:
            avisos.append(f"Linha {n} ({codigo}): sem valor unitário nem total. Linha ignorada.")
            continue
        itens.append(ItemLido(
            codigo=codigo,
            descricao=texto(celula(linha, col["descricao"])),
            ncm=ncm(celula(linha, col["ncm"])),
            unidade=texto(celula(linha, col["unidade"])).upper(),
            quantidade=quantidade,
            valor_unitario=unitario,
            valor_total=total,
            linha=n,
        ))

    competencia, aviso = escolher_competencia([
        ("pasta", competencia_das_pastas(caminho)),
        ("nome do arquivo", competencia_do_nome(caminho)),
        ("cabeçalho da planilha", _competencia_do_cabecalho(aba.linhas[:linha_titulos])),
    ])
    if aviso:
        avisos.append(aviso)
    return Leitura(itens=itens, aba=aba.nome, competencia=competencia, avisos=avisos)


def _competencia_do_cabecalho(linhas):
    """Primeira data ou mês/ano escrito nas linhas acima dos títulos."""
    for linha in linhas:
        for valor in linha:
            data = competencia_do_texto(valor)
            if data:
                return data
    return None
