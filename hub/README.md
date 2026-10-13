# Hub

Tela única que lista as ferramentas da suíte. Para cada uma mostra um cartão com o
estado (**PRONTA**, **FALTAM BIBLIOTECAS** ou **NÃO ENCONTRADA**), um botão **Abrir**, um
**Ver pasta** e, quando falta biblioteca, **Instalar o que falta**, que roda o `pip` numa
janelinha com a saída linha por linha.

Ao abrir uma ferramenta o hub captura a saída do processo e, se a janela morrer nos
primeiros 2,5 s, mostra o que o Python respondeu — sem isso o `pythonw` engole o erro.

![Hub: um cartão por ferramenta, com o estado e os botões](../docs/prints/hub.png)

## Arquivos

- `hub.pyw` — a tela. A lista `FERRAMENTAS` é a fonte da verdade: cada `Ferramenta` tem
  `caminho` (relativo à raiz do repositório) e `requisitos` (`{módulo que se importa:
  pacote do pip}`). Para acrescentar uma ferramenta, basta um item novo.
- `tema.py` — paleta (azul-ardósia como acento sobre cinzas neutros; verde, âmbar e
  vermelho de estado), `estilos()` para o ttk, `faixa()` do topo, `barra_status()`,
  `etiqueta()` e a classe `Tabela` (Treeview com rolagem, zebra e tags `alerta`/`bom`).
  **Toda tela da suíte importa daqui.**

## Abrir

`Abrir Ferramentas.bat` na raiz, ou `pythonw hub\hub.pyw`.
