# Arquitetura: Formula Explorer para o Excel moderno

Alvo: **Microsoft 365 Excel para desktop (64-bit, Current Channel) no Windows 10/11**, com telas de DPI alto.
O comportamento a replicar está em [PESQUISA-FORMULA-EXPLORER.md](PESQUISA-FORMULA-EXPLORER.md).

## 1. Escolha da tecnologia

| Opção | Dá conta do que foi mapeado? | Situação em 2026 | Veredito |
|---|---|---|---|
| **Office.js (Web Add-in)** | **Não.** Não existe API para avaliar o texto de uma fórmula: seria preciso escrever o trecho numa célula, o que altera a pasta e apaga o desfazer. Não há como desenhar sobre a grade (sem camada transparente). Não dá para descobrir que `OFFSET(...)` devolve um intervalo. A UI fica presa a um painel lateral ou a uma janela de navegador separada | É a plataforma estratégica da Microsoft e a mais documentada | Descartada: não faz o principal |
| VSTO | Sim | Suportado, mas congelado no .NET Framework 4.8, sem novos investimentos. Instalação pesada (runtime VSTO, ClickOnce/MSI, certificado) | Descartada: tecnologia legada |
| C++ nativo (COM/ATL) | Sim | Estável | Descartada: muito mais trabalho sem ganho real aqui |
| **Excel-DNA + .NET 10 (C#)** | **Sim, tudo**: modelo de objetos COM completo (`Evaluate`, `Goto`, eventos de seleção, abas e grupos), janelas Win32 para a camada transparente e WPF com DPI por monitor | Excel-DNA 1.9 (nov/2025) suporta oficialmente .NET 10, que é LTS com suporte até nov/2028. Projeto ativo, open source, com comunidade grande | **Escolhida** |

Por que é a que tem mais material para resolver problemas: tudo o que o add-in faz é **modelo de
objetos COM do Excel** (a API do Excel mais documentada que existe, com 25 anos de exemplos),
**.NET/WPF** e **Win32**, todos documentados pela Microsoft. O Excel-DNA é só a "cola" fina que
carrega o .NET dentro do Excel, gera o ribbon e empacota tudo num `.xll`.

Custos aceitos:
- a máquina precisa do **.NET 10 Desktop Runtime x64** (o SDK de desenvolvimento já inclui);
- só uma versão do .NET moderno pode ser carregada por processo do Excel, então há risco de conflito
  se outro add-in usar outro .NET (raro).

Instalação para o usuário final: um arquivo `.xll` de 64 bits. Sem admin, sem instalador; basta
adicionar em *File → Options → Add-ins → Excel Add-ins*.

## 2. Componentes

```
src/
  FormulaExplorer.Core/        net10.0, sem dependência de Excel, 100% testável
    Tokenizer      fórmula → tokens com posição (refs, nomes, funções, operadores, strings, erros, arrays)
    Parser         tokens → árvore (AST) com o intervalo de caracteres de cada nó; precedência real do Excel
    RefColorizer   ordem de aparição → paleta de 7 cores do Range Finder; ref repetida = mesma cor
    HitMap         índice do caractere → menor nó "destacável" (ref, operação binária, grupo (), chamada)
    ValueFormatter resultado do Evaluate → texto do tooltip (G15 invariante, [Blank], "texto", #ERRO, "A1:B2 containing [2x2] …")
  FormulaExplorer.AddIn/       net10.0-windows, Excel-DNA, WPF
    Ribbon         aba "Studio" › grupo "Navigation" › botão grande "Formula Explorer" (supertip igual ao original)
    ExcelGateway   todo o acesso COM: ler fórmula/valor, Evaluate na aba da raiz, navegar,
                   reexibir/reocultar abas, expandir/recolher grupos, eventos de seleção
    ExplorerSession máquina de estados: raiz, alvo dourado, pilha do Back, seleção manual, duplo-clique
    UI/ExplorerWindow  janela WPF não modal pertencente ao Excel (título, área da fórmula, Back, Close)
    UI/FormulaView     controle próprio: FormattedText com quebra de linha, cores por faixa,
                       destaque (#DCDCDC) e dourado (#FFD700) via BuildHighlightGeometry,
                       sublinhado nos links e hit-test por caractere
    UI/ReturnsToolTip  tooltip "Returns" em negrito + valor em cinza, que fica aberto enquanto o mouse estiver parado
    Overlay/OverlayWindow   janela Win32 em camadas, transparente a cliques (WS_EX_LAYERED|TRANSPARENT|NOACTIVATE),
                            pertencente ao Excel; desenha as bordas coloridas dos precedentes
    Overlay/GridGeometry    intervalo → retângulos em pixels de tela por painel (painéis congelados e divididos,
                            zoom, rolagem), via Pane.PointsToScreenPixelsX/Y, recortados à área visível
    Overlay/ViewTracker     enquanto a janela está aberta, verifica ~20×/s a "assinatura da vista"
                            (aba, rolagem, zoom, posição da janela) e redesenha quando ela muda
tests/
  FormulaExplorer.Core.Tests/  xUnit: tokenizer, parser (precedência e spans), cores, hit-test, formatação
  integration/                 scripts PowerShell que abrem o Excel via COM, carregam o .xll, executam
                               cenários e tiram screenshots para conferir contra os quadros do vídeo original
```

## 3. Decisões técnicas importantes
- **Avaliação de trechos**: `Worksheet.Evaluate(texto)` na aba da célula raiz. Se o resultado for um `Range`,
  o nó é navegável; isso cobre refs, nomes, `OFFSET`, `INDEX`, `INDIRECT`, `CHOOSE` e `IF` que devolvem
  ref. A avaliação é feita sob demanda no hover, com cache por fórmula. Fragmentos com mais de
  255 caracteres não são avaliados; o nó raiz usa o valor da própria célula.
- **Threads**: tudo roda na thread principal do Excel, como no original. Chamadas COM são diretas e
  não há marshaling. Quando o Excel está ocupado (por exemplo, em modo de edição de célula), as chamadas
  COM falham com `VBA_E_IGNORE`; a sessão ignora e tenta de novo no próximo tick.
- **Seleção manual**: `Application.SheetSelectionChange` + `SheetActivate` + `WindowActivate`. Navegações feitas pela
  própria sessão ligam um flag que suprime o evento resultante.
- **Fórmula exibida**: `Range.Formula` (sempre em inglês), sem o `{}` das fórmulas matriciais antigas.
- **DPI**: a camada transparente trabalha em pixels físicos, os mesmos do COM do Excel. A janela WPF
  é DPI-aware por monitor.
- **Restauração de abas e grupos**: a sessão guarda o que reexibiu ou expandiu e desfaz ao sair do
  destino (novo clique, Back ou troca de raiz). No Close, o destino atual fica como está.

## 4. Fases
1. **Base**: SDK, solução, `.xll` carregando no Excel, botão no ribbon abrindo uma janela vazia.
2. **Core**: tokenizer, parser, cores e hit-test, com testes.
3. **Janela**: renderização colorida, hover, tooltip "Returns" com `Evaluate`.
4. **Navegação**: clique, dourado, duplo-clique para descer, pilha do Back, seguir a seleção, Close.
5. **Camada transparente**: bordas dos precedentes acompanhando rolagem, zoom e painéis.
6. **Abas e grupos ocultos**, casos de borda, posição e tamanho da janela, teclado (Esc/Enter).
7. **Empacotamento** do `.xll`, instruções de instalação e roteiro de testes de integração.
