"""MODELO para o leitor de uma empresa nova. Este arquivo não aparece na tela (começa com "_").

Como usar:
1. Copie este arquivo nesta mesma pasta e dê o nome da empresa, em minúsculas e sem
   espaço ou acento: ex. "metalurgica_abc.py".
2. Ajuste EMPRESA, CODIGO e CONTA_CONTABIL.
3. Em TITULOS, escreva os títulos das colunas como estão na planilha (acento, maiúscula
   e ":" no fim não importam). Um campo pode ter títulos alternativos:
   "quantidade": ("Quantidade Final", "Qtde Final").
4. Abra a tela do Bloco K: a empresa nova já aparece na lista "Empresa".

Saldo do modelo de importação = valor total do item. Se a planilha tiver a coluna de
total, informe valor_total; se não tiver, o total é quantidade × valor unitário.

Qual aba ler:
  - ler_aba(caminho, -1) = última aba; ler_aba(caminho, 0) = primeira; ou pelo nome.
  - achar_aba(ler_abas(caminho), TITULOS) = a primeira aba que tiver os títulos.
Se a planilha não tiver linha de títulos, troque achar_cabecalho() por colunas fixas,
ex.: col = {"codigo": 0, "descricao": 1, "ncm": 3, ...} (A = 0, B = 1, C = 2...).

Um leitor pronto e funcional está em exemplo.py.
"""
from blocok.nucleo import ItemLido, Leitura, achar_cabecalho, celula, competencia_do_caminho, competencia_do_texto, ncm, numero, texto
from blocok.planilha import ler_aba

EMPRESA = "NOME DA EMPRESA LTDA"
CODIGO = "000"
CONTA_CONTABIL = "54"
INDICADOR_PROPRIEDADE = 0  # 0 = estoque do informante em seu poder

TITULOS = {
    "codigo": "CÓDIGO",
    "descricao": "DESCRIÇÃO",
    "ncm": "NCM",
    "unidade": "UNIDADE",
    "quantidade": "QUANTIDADE",
    "valor_unitario": "VALOR UNITÁRIO",
}
OPCIONAIS = {"valor_total": ("VALOR TOTAL", "TOTAL")}  # colunas que podem faltar


def ler(caminho) -> Leitura:
    aba = ler_aba(caminho, -1)
    linha_titulos, col = achar_cabecalho(aba.linhas, TITULOS, OPCIONAIS)

    itens, avisos = [], []
    for i in range(linha_titulos + 1, len(aba.linhas)):
        linha = aba.linhas[i]
        codigo = texto(celula(linha, col["codigo"]))
        if not codigo:
            continue
        n = i + 1
        quantidade = numero(celula(linha, col["quantidade"]))
        valor = numero(celula(linha, col["valor_unitario"]))
        if quantidade is None or valor is None:
            avisos.append(f"Linha {n} ({codigo}): quantidade ou valor unitário não é número. Linha ignorada.")
            continue
        itens.append(ItemLido(
            codigo=codigo,
            descricao=texto(celula(linha, col["descricao"])),
            ncm=ncm(celula(linha, col["ncm"])),
            unidade=texto(celula(linha, col["unidade"])),
            quantidade=quantidade,
            valor_unitario=valor,
            valor_total=numero(celula(linha, col.get("valor_total"))),
            linha=n,
        ))

    # Competência: nome da aba, depois nome do arquivo/pasta; se nada der, a tela pede para digitar.
    competencia = competencia_do_texto(aba.nome) or competencia_do_caminho(caminho)
    return Leitura(itens=itens, aba=aba.nome, competencia=competencia, avisos=avisos)
