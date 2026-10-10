# F0 — Registro de validação manual no Excel para Web

**Status:** PENDENTE — nenhum resultado de host real registrado  
**Branch de trabalho:** `fix/f0-canonical-headers-regressions`  
**Escopo:** confirmar leitura independente e reconciliação por `ID REGISTRO` após falha de `context.sync()`.  
**Regra:** não marcar como validado até preencher as evidências observadas no Excel para Web. Não executar em workbook de produção.

## 1. Identificação da execução

- Data/hora (incluindo fuso):
- Responsável pela execução:
- URL/identificador do ambiente de teste (não incluir tokens ou segredos):
- Excel para Web / navegador / versão, se disponível:
- Versão/commit do suplemento testado:
- Identificador do workbook descartável (não compartilhar dados pessoais):
- Aba testada: `Email` / `SMS`
- ID sintético exclusivo utilizado:
- Resultado geral: **PENDENTE**

## 2. Pré-condições e segurança

- [ ] Usar uma cópia descartável do workbook, nunca o arquivo de produção.
- [ ] Confirmar que a aba de teste contém os cabeçalhos esperados, sem duplicatas ou ambiguidades.
- [ ] Confirmar a coluna `ID REGISTRO` e a sua posição real.
- [ ] Garantir que o ID usado é sintético, exclusivo e não pertence a outro registro.
- [ ] Registrar o estado inicial da aba e o número de linhas relevantes.
- [ ] Confirmar que a falha será induzida apenas por um mecanismo de teste controlado e aprovado. Não desligar a rede nem interromper gravações reais para forçar a falha.
- [ ] Não incluir tokens, e-mails pessoais, conteúdo de mensagens ou dados sensíveis nas evidências.

## 3. Matriz de cenários

Preencher **resultado observado**, **evidência** e **observações** após cada execução. Não presumir o resultado esperado como resultado real.

| ID | Cenário | Resultado esperado | Resultado observado | Evidência/arquivo | Observações |
|---|---|---|---|---|---|
| M01 | Gravação normal de um registro sintético | Um registro persistido e um ID único | PENDENTE | | |
| M02 | Falha de confirmação após persistência, seguida de leitura independente | Leitura em contexto independente encontra exatamente um ID; status reconciliado; sem segunda escrita | PENDENTE | | |
| M03 | Reconciliação sem correspondência do ID | Resultado não resolvido; nenhuma nova escrita automática | PENDENTE | | |
| M04 | Reconciliação com duas correspondências do mesmo ID em workbook descartável preparado para o teste | Conflito explícito; nenhuma nova escrita automática | PENDENTE | | |
| M05 | Leitura independente indisponível/falha | Resultado indeterminado ou erro explícito; nenhuma nova escrita automática | PENDENTE | | |
| M06 | Cabeçalho `ID REGISTRO` ausente, duplicado ou ambíguo em workbook de teste | Reconciliação bloqueada; nenhuma nova escrita automática | PENDENTE | | |

**Nota de segurança:** M02–M06 devem usar exclusivamente workbook descartável e dados sintéticos. Para testar falha de confirmação após persistência, prefira um ponto de injeção de falha controlado em build de teste; não provoque indisponibilidade em produção.

## 4. Evidências mínimas a anexar

- Captura do Excel para Web mostrando que o ambiente é de teste, ocultando dados não necessários.
- Log do suplemento com mensagem/resultado e ID sintético.
- Estado da linha antes/depois ou captura da coluna `ID REGISTRO`, mostrando que não houve gravação duplicada.
- Registro do cenário, data/hora e commit testado.
- Se houver falha, erro completo sem tokens, dados pessoais ou segredos.

Não anexar credenciais, access tokens, cookies, links privados contendo tokens ou conteúdo real de clientes.

## 5. Critérios de aprovação

- [ ] A leitura independente realmente executa no host Excel para Web após a falha de confirmação.
- [ ] Uma correspondência exata resulta em reconciliação positiva.
- [ ] Zero correspondências mantém o resultado não resolvido, sem retry automático.
- [ ] Mais de uma correspondência gera conflito explícito, sem retry automático.
- [ ] Falha da leitura ou cabeçalho ausente/ambíguo mantém estado indeterminado e não grava novamente.
- [ ] Não existe segunda escrita automática em nenhum dos cenários de resultado incerto.
- [ ] Email e SMS foram avaliados separadamente, ou foi registrada justificativa clara para qualquer cenário ainda não executado.
- [ ] Evidências e desvios foram registrados no relatório F0.

## 6. Resultado final

- Estado final: **PENDENTE — aguarda execução e evidências do Excel para Web**.
- Aprovado por:
- Data/hora:
- Desvios conhecidos:
- Ação corretiva necessária:
- Link do relatório/evidências:

**Importante:** CI e mocks aprovados não equivalem a validação no host real. Não declarar F0 concluída, não fazer merge e não publicar em produção com este registro pendente.
