# Registra Alpha 2.1.5
Preferências de ordem e visibilidade são isoladas por hash de identidade Microsoft (tenant/oid ou issuer/sub) e identidade do arquivo. Não são autenticação nem controle de acesso, somente personalização. Persistência no navegador, sem sincronização entre dispositivos; sem SSO usa sessionStorage por arquivo (fechar aba ao trocar conta), sem identidade ou storage usa memória. Não persiste token, nome ou URL bruta. A demonstração é separada.
As três colunas legadas ORDEM FORMULARIOS, EXIBIR HISTORICO, EXIBIR RESUMO não são excluídas e não determinam preferências pessoais. A criação, edição, tipo, obrigatoriedade, coluna e atividade de perguntas continuam compartilhados. Valores padrão iniciais preservam Valor no Histórico e opcional no Resumo; demais começam ocultos. Config log persistente na página; mensagens operacionais continuam temporárias/locais. Os cenários de dois usuários, dois arquivos, persistência, armazenamento restrito e SSO foram aprovados pelo responsável na homologação de 27/11/2026. Sem alteração do manifesto ou credenciais.

## Organização dos cadastros na Config

A função compartilhada `gravarCadastroOrdenado` em `public/registro-security.js` é chamada por Configurações, E-mail e SMS. A lista afetada é lida, acrescida do valor novo e ordenada com Intl.Collator pt-BR (sensibilidade base, ordem numérica natural). As colunas são gravadas separadamente, com deslocamentos reais do intervalo usado, sem movimentar as demais listas. Mensagens são ordenadas pelo nome com texto associado, inclusive quando as colunas não são adjacentes. A gravação utiliza RegistroSecurity.cell para neutralizar prefixos de fórmula; wrapText=false é aplicado ao intervalo gravado. Não remove quebras explícitas de linha. O cadastro duplicado é bloqueado/ignorado conforme a tela. Não há evento de ordenação para edição manual no Excel.

A rotina lê e regrava a lista afetada; não oferece transação atômica entre operadores. A aprovação de DEV-002 pelo responsável encerra a pendência funcional, não elimina o risco residual de concorrência. Manter backup e procedimento de recuperação.

## Evidências de encerramento — 27/11/2026

- Baseline funcional: [`ba1b10c`](https://github.com/Matheus-Excelencia/Addin_Registros_Envios/commit/ba1b10c3530e18ad36e5a5da5f534569472308d0), branch main.
- [GitHub Actions 46027246753](https://github.com/Matheus-Excelencia/Addin_Registros_Envios/actions/runs/46027246753): 132 aprovados, 0 falhas, 0 ignorados, conforme logs consultados.
- [Vercel — produção](https://vercel.com/acme-vs-coyote/addin-registros-envios/8DCD3Z5yzYQAQrzavEeNPTN4pdVw): baseline funcional READY na consulta anterior ao encerramento.
- Homologação manual: 31/31 itens aprovados por declaração do responsável técnico em 27/11/2026. Não houve execução independente pelo assistente no Excel Web nem coleta de capturas ou logs individuais dos testes manuais. O aceite funcional não substitui auditoria ou certificação.
- Manifesto 2.0.0.0, produto Alpha 2.1.5, identidade Office, SSO e código funcional preservados.

Consulte [matriz final](ALPHA-215-HOMOLOGACAO-FINAL.md). Pendências externas: aprovação e publicação documental no SharePoint, DOC-REG-007 e eventual autorização de distribuição/reabertura. Preenchimento automático de perguntas personalizadas pertence à Alpha 3.0 e não integra esta entrega.
