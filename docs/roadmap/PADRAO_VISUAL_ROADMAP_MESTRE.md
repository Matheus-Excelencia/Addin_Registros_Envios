# Padrão visual obrigatório — Roadmap Mestre Frm Cadastro

## Regra principal
O arquivo de referência visual fornecido pelo usuário, **“ultima tentativa.pdf”**, define o padrão oficial do painel **“Roadmap Mestre Interativo — Frm Cadastro”**. Em toda consulta, atualização, revisão ou avanço do roadmap, preservar a mesma estrutura, hierarquia visual, organização dos cartões, filtros, painéis e disposição dos elementos. Não propor nem aplicar redesign por iniciativa própria.

## Elementos que devem permanecer
1. Título: “Roadmap Mestre Interativo — Frm Cadastro”.
2. Subtítulo de acompanhamento com versão/data e contexto de planejamento, diagnóstico e evolução controlada.
3. Seção “Painel do roadmap” com cartões-resumo:
   - Fases de evolução (F0–F9)
   - Proposta documental (DOC)
   - Próxima ação
   - Implementação autorizada
4. Seção “Filtrar fases”, com os filtros: Todas, Diagnóstico, Pendentes, Propostas e Histórico.
5. Cartões das fases F0–F9 e cartões separados para DOC e HIST.
6. Seção “Detalhes da fase selecionada”, incluindo objetivo, entrega esperada, próxima ação recomendada, critérios de conclusão e estado do diagnóstico.
7. Ação para preparar prompt de produção para a fase selecionada.
8. Seção “Ordem de execução recomendada”.
9. Registro da proposta DOC-REG / GitBook no Roadmap Mestre.
10. Seção “Regras permanentes do roadmap”.

## Comportamento
- Atualizar apenas conteúdo, estados, progresso, evidências, pendências e próxima ação com base em fatos confirmados; preservar o padrão visual.
- Manter a interatividade existente: filtros, seleção de fase e botões que preparam prompts.
- Distinguir claramente fatos comprovados, hipóteses, propostas e decisões pendentes.
- Não marcar uma implementação como autorizada sem decisão explícita.
- Se algum detalhe visual não puder ser reproduzido, preservar o restante do padrão e explicar a limitação; não substituir por outro layout.
- Só alterar estilo ou estrutura visual mediante pedido explícito do usuário.

## Instrução de continuidade para o ChatGPT
Sempre que o usuário mencionar roadmap — por exemplo, consultar, atualizar, revisar, avançar ou planejar — usar este padrão como referência obrigatória. O roadmap pode evoluir em conteúdo; seu estilo não deve mudar sem autorização explícita.

## Limite de persistência
Este documento registra a regra no repositório, na branch `portal-web-evolut`. Isso não equivale a configurar a memória interna ou as instruções globais da conta ChatGPT. Para máxima continuidade entre conversas, manter este arquivo como referência do projeto e, quando possível, incluí-lo nas instruções do projeto ChatGPT.
