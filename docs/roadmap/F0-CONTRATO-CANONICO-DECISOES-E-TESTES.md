# F0 — Consolidação funcional do contrato canônico pós-F0-C.4

- **Projeto:** Frm Cadastro
- **Repositório:** `Matheus-Excelencia/Addin_Registros_Envios`
- **Branch-alvo:** `portal-web-evolut`
- **Tipo:** decisão funcional/documentação; não é implementação
- **Estado:** decisões funcionais consolidadas para orientar a construção; compatibilidade e validação técnica pendentes
- **Data do registro:** 2026-10-09

## 1. Escopo e limites

O responsável pelo projeto autorizou consolidar as decisões funcionais necessárias para construir o Frm Cadastro, priorizando segurança dos registros, compatibilidade, clareza de erros e preservação das funcionalidades existentes.

Esta aprovação autoriza a consolidação documental das regras recomendadas. Não comprova implementação, execução de testes, compatibilidade com workbooks reais, publicação, nem aprovação de permissões externas.

Nesta atividade não se executam testes, não se implementa Microsoft Graph, não se alteram workbooks operacionais, registros, IDs, fórmulas, dados, permissões, configurações de Vercel/Entra ID, deployments ou funcionalidades de Email/SMS. Não se publica nem sincroniza conteúdo com GitBook.

## 2. Base documental e evidência disponível

F0-C.1 a F0-C.4 registraram inspeção documental de arquivos na branch `portal-web-evolut`. Os relatórios anteriores identificaram:

- `Email`: 11 campos propostos/observados, incluindo `ID REGISTRO`.
- `SMS`: 9 campos propostos/observados, incluindo `ID REGISTRO`.
- `Config`: listas e valores de configuração para empresas, supervisores, responsáveis, mensagens e limite SMS.
- Os scripts reconhecem variações de cabeçalhos e procuram cabeçalhos em uma janela de até 30 linhas.
- O formato de ID observado é `ENV-AAAAMMDD-...`; unicidade global, concorrência e recuperação não foram comprovadas.
- O Histórico usa dados das abas e referências de linha/índice; foi identificada dependência de `localStorage` para o fluxo de registro duplicado.
- Os módulos atuais de dados dependem de Office.js/Excel no levantamento; Graph aparece como possibilidade arquitetural, não como integração funcional comprovada.
- A existência de um deployment com estado `READY` não prova que o fluxo funcional esteja disponível ou autenticado.

Essas são evidências de inspeção de código/documentação, não evidências de testes executados nem de que todos os workbooks reais seguem o contrato.

## 3. Matriz de decisões funcionais D1–D8

As recomendações abaixo consolidam a orientação funcional autorizada. Detalhes de implementação podem variar, desde que satisfaçam os critérios e não violem as restrições de segurança.

### D1 — Cabeçalhos canônicos, tipos e formatos

**Regra recomendada**
- Manter como esquema canônico documentado os campos registrados em F0-C.4 para `Email`, `SMS` e `Config`.
- Tratar os nomes canônicos como destino normalizado; aceitar apenas aliases que estejam em catálogo explícito e versionado.
- Definir tipos e formatos em contrato: datas normalizadas, quantidades numéricas inteiras quando aplicável, texto como texto, limite SMS numérico e validação de campos de configuração.
- Não inferir tipo ou destino de coluna a partir de uma correspondência ambígua.
- Preservar os cabeçalhos e os dados existentes durante a fase documental; não renomear colunas automaticamente.

**Justificativa/impacto:** previsibilidade e integridade, com compatibilidade legada controlada. Workbooks não conformes podem precisar de análise ou plano de migração antes de qualquer mudança.

**Critério de aceite:** tabela aprovada de campo, tipo, formato, normalização, exemplos válidos/inválidos e mapeamentos de aliases; comparação posterior com amostras reais.

### D2 — Obrigatoriedade e validação por operação

**Regra recomendada**
- Distinguir campos estruturais necessários para reconhecer uma aba, campos obrigatórios para gravar e campos opcionais de registro.
- Bloquear gravação quando faltarem campos essenciais, houver tipos inválidos ou não for possível mapear a coluna com segurança.
- Não inventar valores padrão para dados de negócio. Um padrão só pode ser usado quando estiver explicitamente definido no contrato e for visível/consistente para o usuário.
- Validar a configuração necessária à operação antes de iniciar a gravação.

**Justificativa/impacto:** reduz registros incompletos e evita que a ausência de informação seja mascarada. Pode bloquear arquivos antigos até que sua compatibilidade seja verificada.

**Critério de aceite:** matriz de obrigatoriedade para registrar Email, registrar SMS, consultar Histórico, montar Resumo e editar Config; cada falha tem mensagem e resultado esperados.

### D3 — Aliases legados

**Regra recomendada**
- Manter compatibilidade apenas com aliases conhecidos, documentados, sem conflito e mapeados para um único campo canônico.
- Comparações podem normalizar espaços e caixa quando a regra for explícita; não remover acentos ou reinterpretar nomes de forma implícita.
- Aliases desconhecidos ou que mapeiem a mais de um campo devem produzir diagnóstico, não uma escolha automática.
- Prever catálogo versionado e política de descontinuação, sem migração automática nesta fase.

**Justificativa/impacto:** mantém compatibilidade sem tornar o reconhecimento permissivo e imprevisível.

**Critério de aceite:** catálogo versionado com alias, campo canônico, normalização permitida e exemplos de conflito; teste de cada alias aprovado.

### D4 — Geração e unicidade de ID

**Regra recomendada**
- Preservar IDs existentes e tratá-los como identidade imutável dos registros.
- Até que um mecanismo mais forte seja aprovado e comprovado, não declarar que o formato observado garante unicidade global.
- A geração futura deve ocorrer uma única vez por novo registro, com verificação de colisão no escopo definido e reconciliação de operações concorrentes.
- Não regenerar ID automaticamente em edição, repetição de operação ou recuperação de falha.
- A regra de unicidade deve ser explícita sobre escopo (workbook, conjunto de dados ou global); a recomendação funcional é exigir unicidade no conjunto de dados compartilhado em que o Histórico e a deduplicação operam.

**Justificativa/impacto:** preserva referências e evita duplicação ou quebra de vínculos. Garantias reais dependem da arquitetura de gravação e da concorrência.

**Critério de aceite:** especificação de geração, escopo, colisão e imutabilidade aprovada; testes de ID ausente/duplicado e concorrência executados em cópias autorizadas.

### D5 — Recuperação após gravação incerta

**Regra recomendada**
- Se o resultado de uma gravação for incerto, não repetir cegamente.
- Reconsultar pelo ID estável e reconciliar o estado antes de decidir se uma nova tentativa é segura.
- Se a consulta também falhar, interromper a operação, informar que o resultado é desconhecido e orientar conferência; não afirmar sucesso ou fracasso sem evidência.
- Repetição automática só será permitida quando houver mecanismo de idempotência demonstrável.

**Justificativa/impacto:** reduz registros duplicados. Pode exigir intervenção do usuário quando não houver confirmação suficiente.

**Critério de aceite:** cenários de gravação concluída sem confirmação, gravação não realizada e consulta indisponível resultam em estados distintos, sem duplicidade silenciosa.

### D6 — Ambiguidades e estruturas inválidas

**Regra recomendada**
- Bloquear escrita quando houver cabeçalhos duplicados, alias conflitante, campo essencial ausente, tipo incompatível ou estrutura cuja interpretação não seja determinística.
- Permitir leitura diagnóstica limitada apenas se o resultado for rotulado como parcial e não puder ser confundido com uma consulta completa.
- Não escolher silenciosamente a primeira coluna, linha ou alias que coincidir.
- Informar o nome do campo e a ação corretiva, sem expor credenciais ou dados sensíveis.

**Justificativa/impacto:** prioriza integridade dos registros e clareza dos erros. Arquivos ambíguos podem ficar temporariamente indisponíveis para escrita.

**Critério de aceite:** cada classe de ambiguidade tem comportamento definido; nenhuma escrita é direcionada a uma coluna ambígua.

### D7 — Compatibilidade e migração

**Regra recomendada**
- Adotar contrato canônico com compatibilidade legada controlada por catálogo e matriz de versões.
- Não migrar, renomear, excluir ou reformatar dados existentes automaticamente.
- Antes de qualquer migração futura, exigir inventário de arquivos, cópia de segurança, validação prévia, plano de reversão e autorização específica.
- Arquivos não reconhecidos devem ser diagnosticados e bloqueados para escrita até avaliação.

**Justificativa/impacto:** permite evolução sem comprometer o histórico e os fluxos de Email/SMS. Exige inventário e testes representativos antes de declarar compatibilidade.

**Critério de aceite:** matriz de compatibilidade aprovada e cada versão suportada associada a testes; migração, se necessária, fica em tarefa e autorização separadas.

### D8 — Fallback, acesso, leitura parcial e concorrência

**Regra recomendada**
- Falta de acesso, sessão expirada ou permissão negada bloqueia a operação protegida; não contornar SSO nem permissões.
- Usar Office.js apenas em contexto compatível e autorizado. Graph permanece opção arquitetural condicionada a decisão de escopo e validação externa; não presumir que está implementado.
- Se a fonte ou leitura estiver incompleta, sinalizar parcialidade e não apresentar Resumo como completo.
- Em conflito concorrente, reler e reconciliar antes de gravar novamente; se não for possível, interromper com erro claro.
- Preservar os caminhos atuais de Email e SMS até que qualquer alteração seja autorizada, testada e aprovada.

**Justificativa/impacto:** prioriza segurança e continuidade sem confundir fallback com autenticação ou consistência.

**Critério de aceite:** matriz de falhas e fallbacks aprovada e testes demonstram bloqueio seguro, sinalização de parcialidade e ausência de repetição insegura.

## 4. Contrato documental de referência

### Email (11 campos)

`Data`, `Realizado por`, `Empresa`, `Qtde`, `Email Resposta`, `Supervisor`, `Obs`, `Historico Externo`, `Mensagem`, `Assunto`, `ID REGISTRO`.

### SMS (9 campos)

`Data`, `Realizado por`, `Empresa`, `Qtde`, `Supervisor`, `Obs.`, `Historico Externo`, `Mensagem`, `ID REGISTRO`.

### Config

Proposta documental: `EMPRESAS`, `SUPERVISORES`, `EMAILS_RESPOSTA`, `REALIZADO POR`, `NOME MENSAGEM EMAIL`, `TEXTO MENSAGEM EMAIL`, `NOME MENSAGEM SMS`, `TEXTO MENSAGEM SMS`, `LIMITE SMS`.

**Ressalva:** nomes, aliases, tipos e obrigatoriedade acima constituem base documental consolidada; a compatibilidade com workbooks reais permanece pendente. Não tratar esta tabela como prova de esquema universalmente implantado.

## 5. Pendências externas E1–E4

### E1 — Vercel
**Falta:** configuração efetiva de build/output, respostas HTTP e conteúdo servido, acesso normal aos Previews e comportamento sob proteções existentes. Estado `READY` não é critério suficiente.

**Evidência de encerramento:** capturas/saídas datadas de configuração e respostas reais, obtidas por responsável autorizado, com URL/ambiente identificados.

### E2 — Microsoft Entra ID
**Falta:** confirmação do aplicativo registrado, redirect URIs, fluxo de autenticação, consentimento e escopos efetivamente concedidos.

**Evidência de encerramento:** configuração e resultado de fluxo autorizado confirmados pelo responsável do aplicativo; nenhum dado secreto deve ser registrado no documento.

### E3 — Microsoft Graph (condicional)
**Falta:** decisão documentada de adoção. Se adotado, confirmar permissões mínimas, acesso autorizado ao arquivo, leitura/escrita e comportamento de sessão expirada/falhas.

**Evidência de encerramento:** teste autorizado de ponta a ponta com resultado observável e escopos confirmados. Até lá, Graph é hipótese arquitetural, não integração comprovada.

### E4 — Workbooks representativos
**Falta:** amostras autorizadas que representem versões reais, aliases, cabeçalhos, fórmulas e configurações em uso.

**Evidência de encerramento:** inventário sem dados pessoais desnecessários, relatório de conformidade por versão e resultados dos testes em cópias de teste. Nenhum workbook operacional será alterado nesta tarefa.

## 6. Plano de testes T01–T12

**Estado geral:** PLANEJADOS — NÃO EXECUTADOS nesta entrega.

Pré-condição geral: contrato funcional consolidado, ambiente de teste autorizado, cópias de workbook isoladas, dados sintéticos, versão/commit identificado e procedimento de restauração. Não executar testes de escrita em dados reais.

Para cada teste registrar: identificador, data, executor autorizado, ambiente, commit/build, dados de entrada (sintéticos), passos executados, resultado esperado, resultado observado, evidência anexada, aprovado/reprovado/bloqueado e referência do defeito.

### T01 — Cabeçalhos canônicos
- **Pré-condição:** cópias de teste com abas e cabeçalhos conforme contrato.
- **Passos:** carregar/inspecionar Email, SMS e Config; executar reconhecimento de estrutura no ambiente de teste.
- **Esperado:** todos os campos reconhecidos sem deslocamento; estrutura classificada conforme o contrato.
- **Evidência:** workbook sintético usado, saída/log de reconhecimento e comparação campo a campo.

### T02 — Cabeçalho ausente
- **Pré-condição:** cópia de teste com um campo essencial removido.
- **Passos:** tentar reconhecer e executar a operação afetada.
- **Esperado:** diagnóstico explícito e bloqueio da escrita; nenhuma coluna errada é usada.
- **Evidência:** entrada reproduzível, mensagem/log e verificação de que não houve gravação.

### T03 — Cabeçalhos duplicados ou conflitantes
- **Pré-condição:** cópia com cabeçalho duplicado e cópia com dois aliases apontando ao mesmo campo.
- **Passos:** reconhecer estrutura e tentar operação de escrita.
- **Esperado:** ambiguidade detectada; escrita bloqueada salvo regra determinística previamente aprovada.
- **Evidência:** cópias sintéticas, diagnóstico e estado dos dados após tentativa.

### T04 — Aliases aprovados
- **Pré-condição:** catálogo de aliases aprovado e cópias cobrindo cada alias.
- **Passos:** testar cada alias permitido, um por cenário, e um alias desconhecido.
- **Esperado:** aliases aprovados mapeiam corretamente; desconhecido/conflitante é rejeitado sem inferência silenciosa.
- **Evidência:** catálogo versionado, resultados por alias e log de mapeamento.

### T05 — Campos e tipos
- **Pré-condição:** matriz de obrigatoriedade/tipos aprovada e conjunto sintético válido/inválido.
- **Passos:** testar datas, quantidades, campos obrigatórios/vazios, limite SMS e configuração inválida.
- **Esperado:** entradas válidas aceitas; inválidas recebem erro definido sem gravação parcial indevida.
- **Evidência:** tabela de casos, valores de entrada, mensagens e comparação antes/depois.

### T06 — ID ausente ou duplicado
- **Pré-condição:** política de ID aprovada e cópias com ID ausente, único e duplicado.
- **Passos:** exercitar geração de novo ID em ambiente isolado e detecção de duplicidade; verificar edição e repetição sem alterar IDs existentes.
- **Esperado:** geração conforme contrato, colisão tratada, ID existente preservado e nenhuma duplicidade silenciosa.
- **Evidência:** IDs sintéticos, logs, estado anterior/posterior e resultado de unicidade no escopo definido.

### T07 — Concorrência
- **Pré-condição:** ambiente de teste autorizado com mecanismo de concorrência representativo e dados sintéticos.
- **Passos:** iniciar duas operações controladas sobre o mesmo conjunto de dados e observar conflito/recuperação.
- **Esperado:** conflito detectado/reconciliado; sem perda ou duplicação silenciosa.
- **Evidência:** sequência temporal das operações, IDs, logs e comparação final dos registros.

### T08 — Sessão expirada
- **Pré-condição:** ambiente autorizado que permita expirar/revogar a sessão de teste de forma segura.
- **Passos:** tentar leitura e escrita protegidas após expiração.
- **Esperado:** operação interrompida, erro claro, sem sucesso falso e sem tentativa de contornar autenticação.
- **Evidência:** estado de sessão, resposta/erro e verificação de ausência de gravação.

### T09 — Falha de gravação e confirmação incerta
- **Pré-condição:** mecanismo de falha controlada em ambiente de teste, sem atingir dados reais.
- **Passos:** simular falha antes da gravação, após persistência sem confirmação e indisponibilidade da consulta de reconciliação.
- **Esperado:** diferenciar os estados; consultar por ID antes de repetir; se não for possível reconciliar, informar resultado desconhecido e parar.
- **Evidência:** logs da falha injetada, busca por ID, estado final e prova de ausência de duplicidade.

### T10 — Fallback
- **Pré-condição:** caminhos primário/fallback disponíveis em ambiente autorizado e matriz de compatibilidade aprovada.
- **Passos:** tornar o caminho primário indisponível em teste e observar a seleção do fallback.
- **Esperado:** fallback somente em contexto compatível e autorizado; origem identificável; nenhum bypass ou mistura silenciosa de fontes.
- **Evidência:** configuração de teste, logs da seleção e resultado funcional.

### T11 — Histórico e Resumo
- **Pré-condição:** dados sintéticos conhecidos em Email e SMS, incluindo datas, quantidades, IDs e cenários parciais.
- **Passos:** consultar Histórico, selecionar registro, conferir associação por ID e comparar agregações do Resumo com cálculo independente.
- **Esperado:** associação correta; totais/filtros corretos; dados parciais identificados; nenhuma dependência exclusiva de posição de linha que cause associação incorreta.
- **Evidência:** conjunto sintético, cálculo esperado, saída do app e comparação reproduzível.

### T12 — Permissão negada
- **Pré-condição:** ambiente autorizado com identidade de teste sem a permissão necessária.
- **Passos:** tentar leitura e escrita sem permissão.
- **Esperado:** acesso negado com mensagem clara, sem exposição indevida, fallback não autorizado ou tentativa de contornar a proteção.
- **Evidência:** configuração/identidade de teste descrita sem segredos, resposta de autorização e verificação de ausência de alteração.

## 7. Critério de encerramento da F0

F0 só pode ser marcada como concluída quando:
1. D1–D8 estiverem documentadas como decisões funcionais aprovadas e houver especificação verificável.
2. A compatibilidade do contrato com workbooks representativos tiver sido avaliada e os casos incompatíveis tiverem plano explícito.
3. E1–E4 tiverem evidências verificáveis ou forem formalmente registradas como bloqueios aceitos, sem declarar a capacidade técnica correspondente como validada.
4. T01–T12 tiverem resultado registrado; testes bloqueados devem indicar a dependência e não contam como aprovados.
5. Defeitos críticos que ameacem perda, corrupção, duplicação ou exposição de registros estiverem resolvidos e retestados.
6. O aceite final e as referências das evidências estiverem registrados no fluxo DOC-REG.

A aprovação funcional atual, isoladamente, não satisfaz estes critérios.

## 8. Preparação DOC-REG / GitBook (não publicada)

**Título proposto:** F0 — Contrato canônico do Frm Cadastro: decisões funcionais e plano de validação  
**Tipo:** proposta documental para revisão/registro  
**Referências:** F0-C.1, F0-C.2, F0-C.3, F0-C.4; este documento; evidências E1–E4 e resultados T01–T12 quando existirem.

**Resumo para DOC-REG:** consolida as regras funcionais recomendadas D1–D8, mantém a compatibilidade com workbooks reais como pendência verificável, separa as dependências externas E1–E4 e prepara os critérios e evidências dos testes T01–T12. Não representa implementação, execução, publicação ou sincronização.

**Estado DOC-REG/GitBook:** preparado para o fluxo existente; não publicado, não sincronizado e sem ID de registro atribuído nesta entrega. A publicação requer seguir o processo autorizado e confirmar o resultado.

## 9. Registro desta entrega

- Conteúdo desta proposta: decisões funcionais recomendadas, contrato de referência, pendências externas, plano de testes e critérios de encerramento da F0.
- Nenhum teste executado.
- Nenhuma implementação ou migração realizada.
- Nenhuma alteração em workbooks, IDs existentes, fórmulas, dados, permissões, Vercel, Entra ID, deployments ou código funcional.
- Nenhuma publicação/sincronização GitBook declarada.
- A existência deste documento no repositório, caso commitado, prova apenas a presença da documentação, não a execução dos testes nem a validação externa.
