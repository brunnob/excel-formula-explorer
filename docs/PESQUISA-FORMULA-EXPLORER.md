# Formula Explorer (Spreadsheet Studio): pesquisa e especificação de comportamento

> Objetivo: documentar com o máximo de precisão como o **Formula Explorer** do add-in
> **Spreadsheet Studio** funcionava, para replicá-lo fielmente.
> Evidências em [`docs/referencias/`](referencias/). Pesquisa feita em 2026-10-09.

Legenda de confiança: **[CONFIRMADO]** visto em screenshot/vídeo ou afirmado no site oficial ·
**[INFERIDO]** dedução forte a partir das evidências · **[ABERTO]** não verificado · **[DECIDIR]** precisa de decisão nossa.

Fontes visuais principais:
- **Vídeo oficial "Formula Explorer Walkthrough"** (Vimeo 30356951, out/2011, 3min02s, Excel 2010, pasta "Mine Aquisition.xlsm"), analisado quadro a quadro. Arquivos `2011-oficial-*`.
- **MrExcel Podcast #1627** (YouTube, jan/2013, versão mais nova). Arquivos `2013-mrexcel-*`.
- **3 screenshots do site oficial** (2011). Arquivos `2011-site-*`.

---

## 1. O produto

| Item | Detalhe |
|---|---|
| Nome | Spreadsheet Studio ("Studio for Excel"), add-in de revisão e auditoria de planilhas |
| Empresa | Spreadsheet Studio Pty Ltd, Melbourne, Austrália (ABN 35 152 538 755), sob licença do "Payne Street Trust" |
| Criador | Joseph (Joe) McDaid: engenharia da computação + contabilidade (Univ. Cape Town), ex-banco de investimento e modelagem financeira, autor de muitas questões do ModelOff 2012 |
| Linha do tempo | Site no ar desde fev/2011 · vídeos no Vimeo em out/2011 · edições Lite (grátis) e Pro (A$80) em 2011–2012, sendo o Formula Explorer **exclusivo da Pro** · a partir de 2013 tudo gratuito · MrExcel #1627 (jan/2013) · última versão conhecida **1.3.0** · site retornando 404 em 2017 |
| Tecnologia | COM add-in **VSTO 2010 + .NET 4.0 Client Profile** (WinForms), com atualização automática semanal. Excel 2007/2010, inicialmente 32-bit; em 2013 já funcionava no 64-bit |
| Por que morreu | O site saiu do ar, o auto-update passou a dar erro ao iniciar o Excel e o add-in sumiu do menu (fórum MrExcel). Em 2018 ainda havia gente pedindo um equivalente (Microsoft Tech Community) |

### Ribbon "Studio"

| Grupo | Controles |
|---|---|
| **Navigation** | **Formula Explorer** (botão grande, ícone de globo azul) |
| Collaboration | New Adv. Comment · Comment Browser · Comment Report |
| Visualisation | Restore · ☑ Backup · Formula · Dependents · Function · Heatmap |
| Support | Settings · About · Help |

Supertip do botão **[CONFIRMADO]** (`2011-oficial-00`): título "Launch Formula Explorer". O texto
diz que a ferramenta serve para navegar com facilidade por fórmulas complexas.

---

## 2. Funcionalidades oficiais (tabela de edições do site)

Hyperlink to Ranges · Hyperlink to Named Ranges · Hyperlink to Range Functions · Unhide Hidden Sheets ·
Expand Grouped Ranges · Modal Operation · Formula Fragment Highlighting · Formula Fragment Solving.

O item "Modal Operation" engana: no vídeo a janela é **não modal** (ver 3.6). O nome provavelmente
quer dizer que a janela opera em "modo" de exploração.

---

## 3. Comportamento observado

### 3.1 A janela **[CONFIRMADO]**
- Título "Formula Explorer"; na barra de título só o **X**, sem minimizar nem maximizar.
- Área de texto somente leitura, fundo branco, com a fórmula **incluindo o `=`**. Fórmulas longas
  **quebram linha** (`2011-oficial-03`), com cerca de 4 linhas visíveis.
- Botões **Back** e **Close** no canto inferior direito. O Close tem o estilo de botão padrão.
- Tooltips dos botões: no Back, "Retraces your steps…" (refaz o caminho de volta); no Close, "Exits Formula Explorer…".
- Tamanho aproximado de **550 × 115 px** a 96 DPI.
- Ao abrir pela primeira vez, a janela aparece no **canto superior direito da janela do Excel**, sobre
  o ribbon (`2011-oficial-01`). No vídeo de 2013 ela aparece no canto superior esquerdo, o que sugere
  que a última posição é lembrada. **[INFERIDO]**
- Ao trocar de fórmula, a área de texto **fica em branco por um instante, com cursor de espera**, e
  então a nova fórmula aparece (`2011-oficial-12`). Ou seja, ela pré-processa algo ao carregar.

### 3.2 Coloração **[CONFIRMADO, cores medidas em pixel]**
É a paleta de 7 cores do Range Finder do Excel, aplicada **por ordem de aparição** e de forma cíclica
(no vídeo de 2013, a 8ª referência volta ao azul):

| # | Cor | RGB |
|---|---|---|
| 1 | Azul | `#0000FF` |
| 2 | Verde | `#008000` |
| 3 | Roxo | `#9900CC` |
| 4 | Bordô | `#800000` |
| 5 | Verde claro | `#00CC33` |
| 6 | Laranja | `#FF6600` |
| 7 | Magenta | `#CC0099` |

- Recebem cor: referências (`K3`, `$I27`, `I$19:$I$19`), referências a outra aba
  (`Sheet8!E7`, `'Inventory Forcast'!C$6`, com o prefixo na mesma cor) e **nomes definidos**
  (`Hundred`, `Aqui_Per`). Exemplo do vídeo: `=SUM(I$19:$I$19)+J$19*$I$28/Hundred-SUM(OFFSET($I$19,0,0,1,Aqui_Per))`
  fica azul, verde, roxo, bordô, verde claro e laranja (`2011-oficial-13`).
- Ficam em **preto**: `=`, funções, operadores, números, textos, parênteses e vírgulas.
- Referência repetida ganha a mesma cor da primeira ocorrência. **[INFERIDO, comportamento do Excel]**

### 3.3 Bordas nos precedentes da planilha **[CONFIRMADO]**
Enquanto a janela está aberta, as células referenciadas que ficam **na aba da fórmula** ganham uma
borda fina na **mesma cor** da referência, no estilo do Range Finder do Excel (`2011-oficial-01`,
`2013-mrexcel-02`). As bordas são atualizadas quando se desce para outra fórmula e **somem ao fechar**
(`2011-oficial-21`). O Excel não está em modo de edição nesse momento, então quem desenha é o add-in.

### 3.4 Hover: destaque e resultado parcial **[CONFIRMADO]**
O mouse sobre um caractere destaca o **nó da árvore sintática** que aquele caractere representa:

| Sob o mouse | Fragmento destacado | Evidência |
|---|---|---|
| Referência / nome | a referência, **sublinhada** na cor dela, com cursor de mão | `2011-site-01`, `2011-oficial-04`, `2013-mrexcel-03` |
| Operador binário | a expressão binária inteira, respeitando precedência e associatividade | `^` → `(($I27/Hundred)+1)^(K$31)` (`2011-oficial-06`); `/` → `E30/C4`; o último `+` → a fórmula toda |
| Parêntese `(` ou `)` | o grupo inteiro, de parêntese a parêntese ("Bracket Forest") | `2011-oficial-05`, `2011-site-03` |
| Nome de função | a chamada inteira `FUNC(...)` | `2011-oficial-14` |
| Função que **retorna intervalo** (OFFSET…) | a chamada inteira, **sublinhada** como um link | `2011-oficial-14` |
| Espaço vazio, abaixo ou à direita do texto | nada | 1m11–1m15 do vídeo |

- Fundo do destaque: **`#DCDCDC`** (Color.Gainsboro), medido em pixel.
- Tooltip padrão do Windows com o título **"Returns"** em negrito e o valor em cinza. Ele aparece
  depois de um pequeno atraso e **fica visível enquanto o mouse estiver parado** (mais de 20 s no vídeo).
- Formato do valor:
  - número com até 15 dígitos significativos: `1.00760985118373`, `0.461538461538462`, `1.062`, `100`, `0.126`;
  - célula vazia aparece como `[Blank]`;
  - **intervalo**: endereço sem `$`, a palavra "containing", a dimensão e os valores, por exemplo
    `I19:J19 containing [1x2] , 0.2520547…`. I19 está vazia, por isso o primeiro valor sai em branco (`2011-oficial-14`).

### 3.5 Clique, "descer" e Back **[CONFIRMADO]**
**Clique simples** numa referência, nome ou função de intervalo:
- O Excel **ativa a aba e seleciona o destino**, rolando até ele: a mesma aba, outra aba
  (`Sheet8!E7`), a aba do nome (`Hundred`, `2011-oficial-07`) ou o intervalo resolvido pela função
  (OFFSET → `I19:J19`, `2011-oficial-15`).
- O elemento clicado ganha um **fundo dourado** (aprox. `#FFD700`, Color.Gold) que marca onde o
  usuário está. Só um elemento por vez fica dourado.
- **A fórmula exibida não muda.** Dá para pular entre as partes da mesma fórmula.

**Dois cliques seguidos** na mesma referência, ou seja, clicar de novo na que já está dourada:
- A janela **"desce" para a fórmula da célula de destino**. Exemplo: `K$31` → mostra `=J31+K19`
  (`2011-oficial-08` e `-09`); depois `J31` → mostra a fórmula de J31 (`-10` a `-13`).
- As bordas coloridas são redesenhadas para a nova fórmula.
- Se isso é um duplo-clique de verdade ou "clicar de novo na dourada" é **[DECIDIR]**: os dois cliques
  ficaram de 0,3 a 0,8 s de distância no vídeo.

**Back** é uma pilha de passos ("retraces your steps"). Sequência observada:
`K33 → (clica K$31, Hundred, $I27, K$31) → desce para K31 → (clica K19, J31) → desce para J31 → clica OFFSET`.
Os três Back seguidos produziram:
1. continua na fórmula de **J31**, seleção volta para J31 e o dourado some (`2011-oficial-18`);
2. fórmula de **K31**, seleciona K31 (`-19`);
3. fórmula de **K33**, seleciona K33 (`-20`).

O modelo que explica isso é o seguinte:
- **Cliques simples não entram na pilha.** Um Back, estando com algo dourado, apenas volta a
  seleção para a célula "raiz", aquela cuja fórmula está exibida.
- **Descer** empilha a raiz anterior.
- O Back sem nada dourado **desempilha**, mostrando a fórmula anterior e selecionando a célula dela.
- O Back fica desabilitado na abertura e é habilitado quando há algo dourado ou a pilha não está vazia.

**Close** (ou X) fecha a janela, **deixa a seleção onde está** e remove as bordas
(`2013-mrexcel-12`, `2011-oficial-21`).

### 3.6 A janela não é modal e acompanha a seleção **[CONFIRMADO, novo]**
Com o Formula Explorer aberto, o usuário clicou em células da planilha (L36 e depois L37), e a janela
**passou a mostrar a fórmula da célula selecionada** (`=Notes_FO!L150` etc., `2011-oficial-02` e `-03`).
O Back ficou habilitado nesse momento, então **seleção manual também empilha** a raiz anterior.
Navegações feitas pelo próprio Formula Explorer não disparam essa troca.

---

## 4. Modelo de estados (para implementar)

```
Estado: raiz (célula cuja fórmula está exibida), alvoDourado (nó ou nulo), pilha (raízes anteriores)

Abrir          : raiz := ActiveCell; pilha := []; alvoDourado := nulo; desenha bordas
Seleção manual : pilha.push(raiz); raiz := ActiveCell; alvoDourado := nulo; re-renderiza
Hover(nó)      : destaca nó (#DCDCDC); se navegável, sublinha; tooltip "Returns" com o valor do nó
Clique(nó nav.): seleciona o destino (reexibe aba / expande grupo se preciso); alvoDourado := nó
Descer(nó nav.): pilha.push(raiz); raiz := 1ª célula do destino; alvoDourado := nulo; seleciona raiz; re-renderiza
Back           : se alvoDourado ≠ nulo -> seleciona raiz; alvoDourado := nulo
                 senão se pilha não vazia -> raiz := pilha.pop(); seleciona raiz; re-renderiza
Close          : remove bordas; fecha (a seleção fica onde está)
Back habilitado = alvoDourado ≠ nulo OU pilha não vazia
```

"Nó navegável" é qualquer referência, nome ou chamada de função cujo resultado avaliado seja um
**objeto Range**. Chamar `Evaluate` em um texto como `OFFSET($I$19,0,0,1,2)` retorna um Range, e era
quase certamente assim que o original detectava as "Range Functions" e montava o texto "I19:J19 containing…".

---

## 5. Implementação provável do original [INFERIDO]
- `Form` WinForms não modal, pertencente à janela do Excel, com `RichTextBox` somente leitura:
  `SelectionColor` para as cores, `SelectionBackColor` para Gainsboro e Gold e `SelectionFont` para o sublinhado.
- `ToolTip` com `ToolTipTitle = "Returns"`; os botões têm o próprio tooltip.
- Um parser de fórmulas gera a árvore com o intervalo de caracteres de cada nó. O hit-test usa
  `GetCharIndexFromPosition` mais a verificação de que o ponto está dentro do caractere.
- Os valores vêm de `Worksheet.Evaluate(fragmento)` na aba da raiz; se o resultado for um `Range`, o nó é navegável.
- Os eventos `SheetSelectionChange` são ignorados durante as navegações iniciadas pela própria janela.

---

## 6. Casos de borda

### Decisões de projeto
1. **Descer para o precedente = um único clique** na referência quando o destino é uma célula única com fórmula (decisão alterada em 2026-10-23). O clique navega, empilha a raiz antiga no Back e mostra a fórmula nova, sem dourado. Para intervalos, constantes e células vazias, o clique só navega e o nó fica dourado. Um duplo-clique habitual desce apenas um nível: o segundo clique é ignorado. O original usava dois cliques (o primeiro navegava e deixava dourado, o segundo descia).
2. **Tudo em inglês**: textos da UI, tooltips e fórmula como em `Range.Formula`, com números em formato invariante (`1.062`). Português "não existe".
3. **Bordas dos precedentes via camada transparente** sobreposta ao Excel, sem alterar a pasta, feita direito desde o início.
4. **Abas ocultas e grupos recolhidos**: reexibir/expandir ao navegar e **reocultar/recolher ao sair** do destino. Exceção: se o usuário fechar a janela parado no destino, ele fica visível. Linhas e colunas ocultas manualmente são apenas selecionadas.
5. **Célula sem fórmula**: mostrar o valor em preto, sem interação.
6. **Tecnologia**: a melhor disponível hoje para o Excel e o Windows atuais, não a do original. Ver [ARQUITETURA.md](ARQUITETURA.md).
7. Fora do escopo: **fórmulas spill e matrizes dinâmicas** (`#`, `@`, `LET`/`LAMBDA`); não é preciso cobrir 100% dos casos.

### Tabela de casos (os itens [DECIDIR] foram resolvidos pelas decisões acima)

| # | Caso | Proposta |
|---|---|---|
| 1 | Célula ativa sem fórmula (constante ou vazia) | **[DECIDIR]** mostrar o valor em preto, sem interação |
| 2 | Seleção de várias células | usar a ActiveCell |
| 3 | Fórmula matricial antiga `{=...}` | mostrar sem chaves; a avaliação parcial pode divergir |
| 4 | **Excel em português**: `=SOMA(A1;B1)`, decimal com vírgula | **[DECIDIR]** exibir como na barra de fórmulas (local) e avaliar internamente em inglês, mapeando token a token |
| 5 | Tooltip num Windows pt-BR | número na cultura local (`1,062`), como o .NET original faria |
| 6 | Mesma referência repetida | mesma cor |
| 7 | Coluna/linha inteira (`A:A`, `1:1`), intervalos, união e interseção | coloridos e navegáveis |
| 8 | Referência 3D `Plan1:Plan3!A1` | colorida; navega para a primeira aba |
| 9 | Referência a outra pasta `[Livro.xlsx]Aba!A1` | navega se a pasta estiver aberta; senão não é clicável |
| 10 | Referência estruturada `Tabela1[Coluna]` | colorida e navegável (resolve como Range); `[@Col]` pode não avaliar fora do contexto da linha |
| 11 | Nome que não é intervalo (`=0,1`, fórmula) | sem cor, sem link, tooltip com o valor |
| 12 | `INDIRECT`/`INDEX`/`CHOOSE`/`IF` que retornam intervalo | navegáveis pela mesma regra do Range |
| 13 | Funções que dependem da célula (`ROW()`, `COLUMN()` sem argumento, `CELL`) | o valor parcial pode sair errado; aceitar |
| 14 | Fragmento com mais de 255 caracteres (limite do `Evaluate`) | tooltip "não foi possível avaliar"; o nó raiz usa o valor da própria célula |
| 15 | Funções voláteis (`RAND`, `NOW`) | o tooltip recalcula e pode diferir da célula; aceitar |
| 16 | Fragmento lento (SUMPRODUCT em intervalo enorme) | avaliar sob demanda no hover, com cache por fórmula |
| 17 | Resultado de erro, texto ou booleano | erro aparece como `#DIV/0!`; texto entre aspas (para diferenciar de `[Blank]`); booleano como VERDADEIRO/FALSO |
| 18 | Resultado matriz (`A1:A3*2`) | `[3x1] v1, v2, v3`, limitado a N valores |
| 19 | Hover no `=`, em vírgula ou em argumento vazio | `=` → fórmula inteira; vírgula → nada |
| 20 | Destino em aba oculta ou muito oculta | **[DECIDIR]** reexibir (como o original) e talvez reocultar ao sair |
| 21 | Destino em linhas/colunas agrupadas recolhidas | expandir o grupo (como o original) |
| 22 | Destino em linhas/colunas ocultas manualmente | **[DECIDIR]** reexibir ou só selecionar |
| 23 | Pasta com estrutura protegida (não dá para reexibir aba) | avisar e não navegar |
| 24 | "Descer" para célula sem fórmula ou para um intervalo | só navegar; não desce |
| 25 | Seleção de gráfico ou forma com a janela aberta | ignorar |
| 26 | Edição de célula com a janela aberta | o Excel bloqueia COM durante a edição; ignorar eventos até sair do modo de edição |
| 27 | Raiz alterada ou recalculada com a janela aberta | re-renderizar no SheetChange e limpar o cache de valores |
| 28 | Cálculo manual | o tooltip usa os valores atuais das células; aceitar |
| 29 | **Bordas dos precedentes**: desenhar com Shapes "suja" a pasta e apaga o desfazer | **[DECIDIR]** janela transparente sobreposta (não altera a pasta) ou adiar |
| 30 | Reexibir abas ou expandir grupos via código apaga a pilha de desfazer do Excel | aceitar (o original fazia o mesmo) |
| 31 | Monitores com DPI alto | a janela precisa ser DPI-aware |
| 32 | Posição e tamanho da janela | abrir no canto superior direito do Excel e lembrar a última posição; **[DECIDIR]** permitir redimensionar |
| 33 | Teclado | Esc fecha; Enter aciona o Close (botão padrão) |
| 34 | Aba com aspas no nome (`'Jo''s Sheet'!A1`) | o tokenizer trata `''` |
| 35 | Estilo de referência R1C1 ligado | **[DECIDIR]** exibir em A1 (mais simples) ou em R1C1 |

---

## 7. Fontes
- Site oficial arquivado (Wayback Machine): [Navigation Tools](http://web.archive.org/web/20120313120615/http://www.spreadsheetstudio.com/tools/navigation-tools/) ·
  [Editions](http://web.archive.org/web/20120313121647/http://www.spreadsheetstudio.com/editions/) ·
  [FAQ](http://web.archive.org/web/20120313120526/http://www.spreadsheetstudio.com/faq/) ·
  [Download/requisitos (2013)](http://web.archive.org/web/20130823170136/http://www.spreadsheetstudio.com/download/) ·
  [Excel Index Function (blog)](http://web.archive.org/web/20120330200546/http://www.spreadsheetstudio.com/excel-index-function/) ·
  [About](http://web.archive.org/web/20120313053748/http://www.spreadsheetstudio.com/contact/aboutus/)
- Vídeos oficiais (Vimeo, canal "Spreadsheet Studio"): [Formula Explorer Walkthrough](https://vimeo.com/30356951) · [Trailer](https://vimeo.com/30935661) · [Visualisation](https://vimeo.com/30474250) · [Collaboration](https://vimeo.com/30712701)
- [MrExcel Podcast #1627 – Spreadsheet Studio (YouTube, jan/2013)](https://www.youtube.com/watch?v=AZgz2XfRC_c)
- [Microsoft Tech Community (2018)](https://techcommunity.microsoft.com/t5/excel/spreadsheet-studio-formula-explorer-for-excel-2013/td-p/283727) ·
  [Fórum MrExcel](https://www.mrexcel.com/board/threads/spreadsheet-studio.938918/) ·
  [Contextures Blog (2014)](https://contexturesblog.com/archives/2014/06/19/excel-tools-for-complex-workbooks/) ·
  [Software Informer – v1.3.0](https://spreadsheet-studio.software.informer.com/)
