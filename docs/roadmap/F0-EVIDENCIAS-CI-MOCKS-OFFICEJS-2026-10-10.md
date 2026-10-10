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

## Próxima ação

Aguardar o CI do commit `788ae2489ef51593c7731c43c00cd0d773bf57ab`. Se estiver verde, avaliar a implementação de reconciliação real por ID em alteração isolada, com testes de leitura independente e IDs duplicados, antes de considerar a fase F0 concluída.
