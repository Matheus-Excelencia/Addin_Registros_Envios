# F0 — Evidências de CI e testes Office.js isolados

**Data:** 10/10/2026  
**Branch:** `fix/f0-canonical-headers-regressions`  
**Escopo:** cabeçalhos canônicos, gravação Email/SMS e tratamento conservador de falha incerta.

## Execução confirmada do commit solicitado

- Commit verificado: [`b5f59dffb16519206a590ae92b9063fd96b6de3a`](https://github.com/Matheus-Excelencia/Addin_Registros_Envios/commit/b5f59dffb16519206a590ae92b9063fd96b6de3a)
- Workflow: **F0 header contract regressions**
- Run: [38018447448](https://github.com/Matheus-Excelencia/Addin_Registros_Envios/actions/runs/38018447448)
- Resultado: **success**
- Job: `f0-regressions`, ID `114113895875`
- Etapas verificadas: checkout, setup Node 20, setup Python, verificação de sintaxe JavaScript, regressões estáticas F0 e testes de comportamento JavaScript isolados — todas concluídas com sucesso.
- O workflow está configurado para executar `node --check` nos quatro scripts, `python -m unittest discover -s tests -p "test_f0_header_contract.py" -v` e `node --test tests/test_f0_functional_isolated.js`.

## Testes adicionados após o commit solicitado

O commit `788ae2489ef51593c7731c43c00cd0d773bf57ab` adiciona mocks da superfície Office.js que executam as funções reais `adicionarRegistroEmail` e `adicionarRegistroSMS` carregadas dos arquivos do projeto:

1. **Falha antes da persistência:** o mock rejeita o `sync()` antes de gravar. A função falha, não surge nova linha e não há segunda tentativa.
2. **Falha de confirmação após persistência:** o mock persiste a linha e depois rejeita `sync()`. A função rejeita, a linha com o ID permanece e não há segunda gravação.
3. **Gravação confirmada:** o mock confirma a linha com o ID esperado, com uma tentativa.
4. **Reconciliação por ID (simulação separada):** os testes `SIMULATION ONLY` modelam ausência, registro existente e IDs duplicados. Eles não executam a persistência Office.js nem provam que a aplicação reconcilia automaticamente.

## Limite funcional importante

Os testes com mocks são isolados e não conectam ao Excel/Office.js real. Eles mostram que o código de Email/SMS atual não faz reconciliação automática por ID quando o `context.sync()` falha depois de a linha já ter sido persistida: a função propaga o erro e não tenta novamente. O protocolo de reconciliação continua sendo apenas simulação. A reconciliação real exigiria uma leitura independente/segura após falha, pesquisa exata do ID e tratamento de zero, uma ou múltiplas correspondências, sem reexecutar a escrita.

## Salvaguardas mantidas

- Office.js continua sendo o fluxo principal.
- Nenhuma alteração de SSO, permissões, manifest, Graph ou proteção SSO do Vercel.
- Nenhum merge e nenhuma publicação em produção.
- A ausência/ambiguidade de cabeçalhos bloqueia a gravação; não são criadas colunas automaticamente.
- Não há retry cego após resultado incerto.

## CI confirmado após correção do mock

- Commit: [`f7999d94500f59b0478d687c16b6b28246f96ccc`](https://github.com/Matheus-Excelencia/Addin_Registros_Envios/commit/f7999d94500f59b0478d687c16b6b28246f96ccc)
- Run: [38018616034](https://github.com/Matheus-Excelencia/Addin_Registros_Envios/actions/runs/38018616034)
- Resultado: **success**; job `f0-regressions` concluído.
- Evidências dos logs: sintaxe JavaScript aprovada; 5 testes de contrato estáticos aprovados; 11 testes isolados aprovados (11/11).
- A primeira execução dos mocks falhou porque o mock não expunha `getCell().format.wrapText`; a correção acrescentou essa superfície ao mock, sem alterar o código de produção. A execução seguinte passou.
- Os testes com funções reais + mock confirmam os caminhos simulados de falha antes da persistência, falha de confirmação após persistência e gravação confirmada. Isso não é teste contra Excel real.

## Reconciliação por ID — implementação inicial em validação

Commits na branch isolada:
- `03bd2b7`: adiciona leitura independente e reconciliação por ID ao fluxo Email.
- `e0d6917`: aplica a mesma proteção ao fluxo SMS.
- `1a57cf7`: ajusta testes das funções reais com mocks para cobrir os novos resultados.
- `1c619b5`: garante que o mock de reconciliação usa contexto separado e lê o estado compartilhado da planilha simulada.

Comportamento implementado:
1. A falha no `context.sync()` final é tratada como resultado de escrita incerto; a função não repete a escrita.
2. É iniciada uma leitura via `Excel.run` independente para procurar o cabeçalho e o ID exato.
3. Uma única correspondência considera o registro reconciliado; nenhuma correspondência gera resultado não resolvido; múltiplas correspondências geram conflito explícito.
4. Cabeçalho canônico ausente/ambíguo ou falha na leitura independente bloqueiam a conclusão e não disparam nova gravação.

**Estado de validação:** CI do commit `1c619b5` passou no run [38018765949](https://github.com/Matheus-Excelencia/Addin_Registros_Envios/actions/runs/38018765949): sintaxe JS, 5 testes estáticos e 11 testes isolados aprovados. Depois disso, foram acrescentados testes diretos para zero/uma/múltiplas correspondências e falha de leitura no commit `7151ed2`; o CI específico desses novos testes foi aprovado no run [38018823659](https://github.com/Matheus-Excelencia/Addin_Registros_Envios/actions/runs/38018823659): sintaxe JavaScript, 5 testes estáticos e 13 testes isolados aprovados (incluindo zero/uma/múltiplas correspondências e falha de leitura). A execução foi disparada para o merge ref do PR #7 incluindo o commit `7151ed2`; o log identifica o merge ref `af1b24f`. Isso comprova a suíte automatizada desse snapshot, não validação no Excel Web real. Os mocks não equivalem a um teste no Excel Web real; validar o comportamento no host Office.js permanece obrigatório antes de considerar F0 concluída. A chamada aninhada de `Excel.run` e o comportamento de leitura após falha de sincronização devem ser verificados no ambiente Excel suportado.

## Próxima etapa — validação manual no Excel Web (reconciliação implementada; host real pendente)

1. **Não repetir a escrita** quando `context.sync()` falhar com resultado de persistência incerto.
2. Depois da falha, tentar uma **leitura independente** da faixa usada em novo `Excel.run`/contexto, sem reutilizar comandos pendentes do contexto que falhou.
3. Resolver o cabeçalho canônico `ID REGISTRO` com as regras atuais; se ausente ou ambíguo, não prosseguir.
4. Contar correspondências exatas para o ID tentado: zero = resultado não resolvido, sem regravação automática; uma = registro encontrado e reconciliado; mais de uma = conflito de IDs, bloquear e exigir análise.
5. Se a leitura independente falhar, preservar o estado indeterminado e orientar consulta ao histórico antes de qualquer nova tentativa.
6. Os testes com mocks já cobrem leitura independente, zero/uma/múltiplas correspondências, falha de leitura e ausência de segunda escrita. Próximo: executar o roteiro manual no Excel Web antes de alegar validação no host real.
7. Roteiro manual: (a) abrir uma cópia descartável do workbook no Excel para Web com o suplemento na branch/preview de teste; (b) confirmar cabeçalhos completos e exclusivos em Email/SMS, incluindo `ID REGISTRO`; (c) fazer uma gravação controlada normal e confirmar ID único; (d) simular falha de confirmação somente em ambiente de teste, sem remover conexão/rede durante gravação de dados reais; (e) confirmar que a reconciliação abre leitura independente e classifica exatamente um ID como reconciliado; (f) em cópias de teste separadas, testar ID ausente, ID duplicado e leitura independente indisponível; (g) confirmar que nenhum cenário realiza segunda escrita automaticamente; (h) guardar horário, host, navegador, ID sintético e resultado sem incluir dados pessoais. Não provoque falhas em workbook de produção.

A lógica de reconciliação por ID está implementada na branch isolada. O que falta é validar no Excel Web real, com cópia descartável do workbook e IDs de teste exclusivos, sem merge nem publicação em produção.
