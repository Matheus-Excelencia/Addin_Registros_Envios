# Automação, qualidade e publicação controlada

**Escopo inicial:** repositório `Matheus-Excelencia/Addin_Registros_Envios`, branch de referência `portal-web-evolut`.

## Inventário verificado em 09/10/2026

| Área | Evidência observada | Limite atual |
|---|---|---|
| GitHub | Repositório acessível por integração conectada; leitura e escrita de branches/arquivos disponíveis. | Branches listadas como não protegidas; não foi possível confirmar regras de proteção ativas. |
| Branch de referência | `portal-web-evolut` no commit `f10204cdda917faace174cacb4f66be530b6d5fc`. | Não equivale a autorização para publicar em produção. |
| CI existente | A API do GitHub retornou zero execuções de Actions recentes e a pasta `.github/workflows` não existia na branch consultada. | Não havia evidência de pipeline de qualidade executando nessa branch. |
| Vercel | Projeto `addin-registros-envios` conectado ao repositório GitHub. | Plano Hobby; Preview mais recente READY, mas isso não prova teste funcional nem integração autenticada. |
| Preview/SSO | Proteção SSO habilitada para deployments não servidos por domínios customizados. | Testes externos sem sessão autorizada não conseguem validar telas protegidas; não contornar a proteção. |
| Produção | Domínio `addin-registros-envios.vercel.app` aparece no projeto e no manifesto do suplemento. | Não publicar nem trocar alias de produção neste fluxo de baseline. |
| Manifesto Office | `manifest.xml` declara `ReadWriteDocument`, `ExcelApi 1.4`, `IdentityAPI 1.3` e `WebApplicationInfo`. | Declarações no manifesto não provam que o registro Entra está ativo, que há consentimento ou que autenticação/gravação funcionam. |
| Dependências | Estrutura contém HTML/CSS/JavaScript e não foi encontrado `package.json` na raiz da branch consultada. | Pipeline inicial usa Node e Python nativos; não presume build de framework nem instala dependências de aplicação. |
| GitBook / DOC-REG | Documentos técnicos versionados no GitHub. | Conexão/publicação no GitBook ainda não comprovada; não afirmar sincronização. |
| Microsoft Entra / Graph | O manifesto contém um App ID e escopos declarados. | Não há acesso administrativo ao tenant confirmado; Graph permanece condicional e não é considerado implementado. |
| Amostras de planilhas | Contratos e scripts documentados no repositório. | Compatibilidade real com arquivos representativos permanece pendente de amostras autorizadas e testes no Excel. |

## Automação inicial

O workflow `.github/workflows/quality-gate.yml` faz verificações estáticas sem segredos e sem publicar:

1. Verifica sintaxe dos arquivos JavaScript em `app/scripts/` com `node --check`.
2. Valida o XML do manifesto do Office.
3. Confere declarações de capacidades e referências de assets locais.
4. Confirma a presença das páginas centrais do suplemento.

O workflow roda em pull requests direcionados a `portal-web-evolut` ou `main`, em pushes para a branch de trabalho, por acionamento manual e semanalmente quando estiver disponível na branch padrão do GitHub. A agenda do GitHub Actions só é executada a partir da branch padrão; portanto, a verificação semanal não deve ser considerada ativa antes da incorporação controlada do workflow nessa branch.

## O que este pipeline não testa

- Excel/Office runtime e APIs Office.js em uma sessão real.
- Leitura/gravação em uma pasta de trabalho real ou compatibilidade de cabeçalhos.
- Login Entra, tenant, redirect URIs, consentimento ou autorização.
- Microsoft Graph, que não está validado como caminho de dados.
- Timeout, concorrência, persistência, IDs duplicados ou recuperação após gravação incerta.
- Proteção SSO/HTTP do Preview em uma sessão autorizada.
- Deploy, alias ou mudança no domínio de produção.

Esses cenários exigem testes de integração separados. Não simular sucesso nem transformar uma checagem estática em evidência de funcionamento.

## Regra de publicação controlada

- Usar branch e pull request para cada conjunto de mudanças.
- Exigir verificações estáticas antes de considerar uma proposta pronta para revisão.
- Usar Preview e dados de teste antes de promover.
- Não fazer deploy direto, não alterar aliases/domínio de produção e não instalar o manifesto do Preview no Excel sem validação específica.
- Bloquear gravações quando o esquema estiver ausente ou ambíguo; nunca criar colunas ou repetir gravação incerta automaticamente sem reconciliação.
- Não incluir tokens, segredos, dados pessoais ou planilhas reais em logs e artefatos.
- Registrar o resultado, commit, evidências e pendências no Roadmap Mestre e DOC-REG.

## Próximos passos automáticos

1. Executar este gate no pull request e corrigir falhas objetivas.
2. Inspecionar os módulos Email/SMS/Histórico/Configuração e adicionar testes unitários sem modificar comportamento funcional não validado.
3. Planejar testes de integração para Office.js e esquema da planilha com dados autorizados.
4. Validar Preview com sessão SSO autorizada e manter produção inalterada.
5. Confirmar a possibilidade de publicação complementar no GitBook; até lá, o GitHub é a fonte versionada.
