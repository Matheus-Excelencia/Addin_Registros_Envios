# F0 — Auditoria de evidências do contrato canônico

- **Projeto:** Frm Cadastro
- **Data da auditoria:** 2026-10-10
- **Branch auditada:** `portal-web-evolut`
- **Commit auditado:** `f10204cdda917faace174cacb4f66be530b6d5fc`
- **Documento-base:** [F0 — Contrato canônico, decisões e testes](F0-CONTRATO-CANONICO-DECISOES-E-TESTES.md)
- **Resultado:** **F0 aberta — divergências técnicas identificadas; testes de workbook não executados.**
- **Escopo executado:** leitura do repositório, manifesto, scripts, workflow do GitHub Actions, PR aberto e metadados de deployments Vercel. Não houve alteração de código funcional, planilhas, permissões ou produção.

## 1. Evidências verificadas

| Área | Evidência observada | O que prova | O que não prova |
|---|---|---|---|
| Contrato F0 | O documento-base registra D1–D8, E1–E4 e T01–T12; declara que os testes estão planejados e não executados. | Há regras funcionais e critérios documentados. | Não comprova implementação, compatibilidade de workbooks nem aceite final. |
| Manifesto Office | `manifest.xml` declara `ReadWriteDocument`, ExcelApi 1.4, IdentityAPI 1.3 e `WebApplicationInfo` com escopos `openid`, `profile`, `User.Read`. | São declarações presentes no manifesto versionado. | Não comprova registro Entra ativo, redirect URI, consentimento, tokens válidos ou autorização efetiva. |
| Módulos de dados | `app/scripts/email.js`, `sms.js`, `historico.js`, `resumo.js` e `config.js` usam `Excel.run`; Email/SMS também chamam `Office.auth.getAccessToken`. | O caminho principal atual é Office.js no contexto Excel. | Não comprova que o SSO funcione em todos os ambientes nem que exista backend/Graph funcional. |
| CI | O GitHub Actions run [38017114804](https://github.com/Matheus-Excelencia/Addin_Registros_Envios/actions/runs/38017114804) concluiu com `success` para o commit `2263bf254007c11eb996177fec192f880ef59fc0` do PR #5. | As verificações configuradas passaram para aquele commit. | Não é teste funcional do add-in, autenticação, concorrência ou gravação em Excel. |
| PR #5 | [PR #5](https://github.com/Matheus-Excelencia/Addin_Registros_Envios/pull/5) está aberto como rascunho, de `automation/ci-baseline` para `portal-web-evolut`. | A proposta de CI ainda está em revisão. | Não está integrada à branch-alvo. |
| Vercel | Foram observados deployments de Preview `READY` de diferentes branches, incluindo `automation/ci-baseline` e `fix/f4-header-ambiguity-guards`, com `target: null`. | Existem artefatos de Preview associados aos commits listados pelo Vercel. | `READY` não comprova fluxo funcional, acesso sem SSO, autenticação ou produção. |

## 2. Divergências de contrato encontradas no código da branch auditada

### F0-A01 — Normalização de cabeçalhos remove acentos

Em `app/scripts/email.js`, `sms.js`, `historico.js` e `config.js`, as funções de normalização usam decomposição Unicode e removem marcas diacríticas. Isso faz com que variações acentuadas sejam tratadas como equivalentes, enquanto D3 do contrato diz para não remover acentos implicitamente.

**Risco:** aliases ou cabeçalhos distintos podem colidir após a normalização.

**Tratamento exigido:** catálogo explícito de aliases e normalização autorizada; detectar colisões antes de mapear/gravar. Não ampliar compatibilidade silenciosamente.

### F0-A02 — Índices de cabeçalho podem ser sobrescritos

Nos mapeamentos locais de cabeçalho de Email e SMS na branch auditada, a atribuição `mapa[normalizar(valor)] = indice` substitui um índice anterior se houver duas colunas que normalizem para a mesma chave. O código pode, portanto, perder a evidência da duplicidade antes da validação.

**Risco:** escrita na coluna errada quando há cabeçalho repetido ou aliases conflitantes.

**Tratamento exigido:** coletar todas as ocorrências por chave; rejeitar estruturas ambíguas antes de qualquer escrita. A validação posterior de campos ausentes não substitui a detecção de duplicidade.

### F0-A03 — Criação automática de cabeçalho e ID

Em `email.js` e `sms.js`, quando não se reconhece uma estrutura de cabeçalho, o código pode escrever cabeçalhos padrão na linha inicial. Quando `ID REGISTRO` não existe, o código cria automaticamente a coluna seguinte.

**Conflito:** D6/D7 e as regras de integridade do F0 determinam bloqueio em estrutura não reconhecida/ambígua e vedam alterações estruturais implícitas nesta fase.

**Risco:** modificar uma planilha não reconhecida ou deslocar/introduzir cabeçalhos sem inventário e autorização.

**Tratamento exigido:** bloquear escrita em estrutura não reconhecida; não criar coluna nem cabeçalho implicitamente. Qualquer migração precisa de tarefa separada, backup, validação e autorização específica.

### F0-A04 — Garantia de unicidade do ID não demonstrada

Email/SMS geram IDs com data, trecho de `Date.now()` e parte de `Math.random()`. O formato pode reduzir colisões, mas não demonstra unicidade global nem coordenação entre gravações concorrentes.

**Tratamento exigido:** manter ID imutável após criação; definir escopo de unicidade; verificar colisão no conjunto de dados relevante; reconciliar concorrência e resultado incerto. Não declarar unicidade garantida antes de testes.

### F0-A05 — Repetição de escrita e confirmação

O código de SMS tenta gravar até três vezes quando a confirmação do ID não corresponde. Cada tentativa calcula novamente a próxima linha e tenta gravar o registro. Isso merece validação específica para o cenário em que a primeira gravação tenha sido persistida, mas a confirmação não tenha sido observada.

**Risco:** possível duplicação em resultado incerto, dependendo do comportamento do Excel/Office.js e do erro observado.

**Tratamento exigido:** consultar/reconciliar pelo ID antes de repetir; só permitir retry quando o estado anterior for conhecido ou a operação for idempotente comprovada. O cenário ainda precisa de teste controlado.

### F0-A06 — Histórico depende de posição de linha para navegação

O Histórico guarda `linhaPlanilha` para navegar até o registro e também carrega `ID REGISTRO` como campo de dados. A navegação por linha pode ficar desatualizada após ordenação, inserção ou exclusão de linhas.

**Tratamento exigido:** usar ID para identidade lógica e resolver a linha atual antes de navegar quando possível; tratar ID ausente/duplicado como diagnóstico, não como associação confiável. T11 permanece pendente.

### F0-A07 — Configuração pode prosseguir com aliases conflitantes

A leitura de Config cria um mapa simples com normalização e escolhe o índice encontrado para listas e modelos. Cabeçalhos repetidos ou nomes que colidam após remover acentos podem mascarar conflito.

**Tratamento exigido:** detectar duplicidade/conflito antes de usar listas, modelos ou limite SMS; não assumir que uma configuração parcial é válida.

## 3. Branches e correções existentes fora da branch auditada

A branch `fix/f4-header-ambiguity-guards` existe e há Preview Vercel `READY` associado a ela. A inspeção desse branch encontrou detecção explícita de cabeçalhos duplicados em SMS e Histórico. Porém:
- esse código está fora da branch `portal-web-evolut` auditada;
- a evidência inspecionada não demonstra integração/merge;
- Email ainda contém lógica de criação automática de `ID REGISTRO` no código inspecionado;
- os guards não substituem a validação completa de F0 nem a execução de T01–T12.

**Decisão:** registrar essa branch como trabalho candidato a revisão; não assumir que as correções já fazem parte da versão-base e não fazer merge automático nesta auditoria.

## 4. Decisões canônicas reafirmadas

1. A planilha Excel continua sendo a fonte de verdade até decisão arquitetural formal.
2. Office.js é o caminho comprovado no código auditado; Graph permanece condicional e não comprovado como integração funcional.
3. Campo/cabeçalho ambíguo, ausente ou incompatível bloqueia escrita quando comprometer o mapeamento seguro.
4. Nenhum cabeçalho ou coluna é criado automaticamente em uma estrutura não reconhecida nesta fase.
5. Aliases só são aceitos se constarem de catálogo explícito e versionado, com regras de normalização definidas.
6. IDs existentes não são regenerados; geração, escopo de unicidade, colisão e concorrência exigem contrato verificável.
7. Resultado incerto de gravação exige reconciliação antes de retry; sem reconciliação, interromper e informar estado desconhecido.
8. Leitura parcial, permissão negada ou sessão expirada não pode ser apresentada como sucesso/completude.
9. Email/SMS e seus esquemas distintos, incluindo `Obs` versus `Obs.`, devem ser preservados até que a mudança seja validada.
10. Nenhum teste de escrita deve usar workbook operacional ou dados reais nesta fase.

## 5. Plano de testes e estado

| Teste | Estado atual | Pré-condição/bloqueio |
|---|---|---|
| T01 — Cabeçalhos canônicos | Planejado; não executado | Workbooks sintéticos e harness de reconhecimento. |
| T02 — Campo essencial ausente | Planejado; não executado | Cópia sintética; confirmar bloqueio sem escrita. |
| T03 — Cabeçalhos duplicados/conflitantes | Planejado; não executado | Cobrir aliases iguais após normalização. |
| T04 — Aliases aprovados/desconhecidos | Planejado; não executado | Catálogo de aliases versionado. |
| T05 — Campos e tipos | Planejado; não executado | Matriz de obrigatoriedade e entradas sintéticas. |
| T06 — ID ausente/duplicado | Planejado; não executado | Escopo de unicidade e dados isolados. |
| T07 — Concorrência | Planejado; não executado | Harness que reproduza duas gravações concorrentes. |
| T08 — Sessão expirada | Planejado; não executado | Identidade de teste autorizada. |
| T09 — Falha e confirmação incerta | Planejado; não executado | Injeção controlada de falha em cópia de teste. |
| T10 — Fallback | Planejado; não executado | Matriz de caminhos e permissões aprovadas. |
| T11 — Histórico/Resumo | Planejado; não executado | Dataset sintético com IDs, linhas alteradas e dados parciais. |
| T12 — Permissão negada | Planejado; não executado | Identidade de teste sem permissão necessária. |

Não foi executado teste funcional em Excel neste trabalho. O CI verde do PR #5 é uma verificação estática separada e não altera o estado destes testes.

## 6. Dependências externas E1–E4

- **E1 Vercel:** há metadados de deployments Preview; ainda não foi provado o comportamento HTTP/funcional nem o acesso autorizado ao Preview. Não alterar SSO nem aliases.
- **E2 Entra ID:** manifesto inspecionado; configuração efetiva, redirect URIs, consentimento e fluxo real permanecem não verificados.
- **E3 Graph:** decisão de adoção e teste de ponta a ponta permanecem pendentes. Não adicionar Graph como caminho principal por inferência.
- **E4 Workbooks:** faltam cópias representativas autorizadas para avaliar variações reais de cabeçalhos, fórmulas e Config. Não acessar ou modificar workbooks operacionais.

## 7. Critério de saída da F0

**F0 não está concluída.** Para fechar:
- resolver ou aceitar formalmente F0-A01 a F0-A07, com referência a commit/PR e justificativa;
- alinhar o código à política de cabeçalhos/aliases sem mudanças destrutivas;
- definir escopo de ID e comportamento de retry/reconciliação;
- executar T01–T12 em ambiente isolado ou registrar cada bloqueio com responsável/dependência;
- fechar E1–E4 com evidência ou bloqueio formalmente aceito, sem alegar capacidade não verificada;
- registrar aceite e referências no fluxo DOC-REG autorizado.

## 8. Registro DOC-REG

- **Título proposto:** F0 — Auditoria de evidências e divergências do contrato canônico
- **Tipo:** relatório de auditoria técnica/documental
- **Fonte versionada:** este arquivo no GitHub
- **Estado GitBook:** não publicado/sincronizado; destino e acesso não comprovados
- **Estado dos testes:** não executados
- **Estado de implementação:** divergências registradas; nenhuma correção funcional foi aplicada por este relatório
- **Estado de publicação:** nenhuma publicação de produção realizada

## 9. Referências verificáveis

- [Contrato F0 na branch auditada](https://github.com/Matheus-Excelencia/Addin_Registros_Envios/blob/portal-web-evolut/docs/roadmap/F0-CONTRATO-CANONICO-DECISOES-E-TESTES.md)
- [Manifesto Office](https://github.com/Matheus-Excelencia/Addin_Registros_Envios/blob/portal-web-evolut/manifest.xml)
- [Email.js](https://github.com/Matheus-Excelencia/Addin_Registros_Envios/blob/portal-web-evolut/app/scripts/email.js)
- [SMS.js](https://github.com/Matheus-Excelencia/Addin_Registros_Envios/blob/portal-web-evolut/app/scripts/sms.js)
- [Histórico.js](https://github.com/Matheus-Excelencia/Addin_Registros_Envios/blob/portal-web-evolut/app/scripts/historico.js)
- [Config.js](https://github.com/Matheus-Excelencia/Addin_Registros_Envios/blob/portal-web-evolut/app/scripts/config.js)
- [Mapeamento técnico Alpha 2.5](https://github.com/Matheus-Excelencia/Addin_Registros_Envios/blob/portal-web-evolut/docs/ALPHA-2.5-MAPEAMENTO-TECNICO.md)
- [Workflow CI — execução do PR #5](https://github.com/Matheus-Excelencia/Addin_Registros_Envios/actions/runs/38017114804)
- [PR #5 — baseline de qualidade](https://github.com/Matheus-Excelencia/Addin_Registros_Envios/pull/5)
- [Branch candidata com guards de ambiguidade](https://github.com/Matheus-Excelencia/Addin_Registros_Envios/tree/fix/f4-header-ambiguity-guards)
