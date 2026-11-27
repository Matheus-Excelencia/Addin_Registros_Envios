# Registra — homologação final Alpha 2.1.5

**Data:** 27/11/2026 · **Responsável técnico:** Matheus Passos · **Resultado funcional:** aprovado conforme declaração do responsável.

## Evidências de encerramento — 27/11/2026

- Baseline funcional: [`ba1b10c`](https://github.com/Matheus-Excelencia/Addin_Registros_Envios/commit/ba1b10c3530e18ad36e5a5da5f534569472308d0), branch main.
- [GitHub Actions 46027246753](https://github.com/Matheus-Excelencia/Addin_Registros_Envios/actions/runs/46027246753): 132 aprovados, 0 falhas, 0 ignorados, conforme logs consultados.
- [Vercel — produção](https://vercel.com/acme-vs-coyote/addin-registros-envios/8DCD3Z5yzYQAQrzavEeNPTN4pdVw): baseline funcional READY na consulta anterior ao encerramento.
- Homologação manual: 31/31 itens aprovados por declaração do responsável técnico em 27/11/2026. Não houve execução independente pelo assistente no Excel Web nem coleta de capturas ou logs individuais dos testes manuais. O aceite funcional não substitui auditoria ou certificação.
- Manifesto 2.0.0.0, produto Alpha 2.1.5, identidade Office, SSO e código funcional preservados.

## Matriz de aceite

### Config e formulários — 6/6

- [x] Cadastrar empresa pelo formulário de E-mail
- [x] Cadastrar empresa pelo formulário de SMS
- [x] Cadastrar supervisor e realizado por pelos formulários
- [x] Cadastrar e-mail de resposta pelo formulário
- [x] Confirmar ordenação dos nomes e textos das mensagens
- [x] Confirmar ausência de quebra automática de texto

### Preferências pessoais — 6/6

- [x] Testar dois usuários Microsoft diferentes
- [x] Testar dois arquivos Excel diferentes
- [x] Confirmar persistência após fechar e abrir o suplemento
- [x] Confirmar ordem individual dos formulários
- [x] Confirmar visibilidade individual no Histórico e Resumo
- [x] Testar armazenamento indisponível e identificação manual

### Site oficial — 5/5

- [x] Confirmar botão de voltar ao topo no index
- [x] Confirmar botão de voltar ao topo no artigo
- [x] Verificar imagem do Brad Pitt no Dark Mode
- [x] Verificar layout restaurado do backup
- [x] Confirmar funcionamento em telas menores

### Segurança e autenticação — 7/7

- [x] Confirmar identificação automática Microsoft SSO
- [x] Confirmar fallback manual quando o SSO estiver indisponível
- [x] Verificar gravação de textos semelhantes a fórmulas
- [x] Verificar números negativos, zero, datas e campos vazios
- [x] Testar isolamento da duplicação entre arquivos
- [x] Testar expiração e consumo único da duplicação
- [x] Inspecionar logs de produção

### Demonstração — 4/4

- [x] Comparar funcionalidades com o suplemento real
- [x] Testar novos cadastros e ordenação
- [x] Testar reinicialização dos dados fictícios
- [x] Confirmar isolamento da autenticação real

### Pendências anteriores — 3/3

- [x] UX-005 — efeitos visuais da abertura
- [x] UX-006 — fluidez da animação de boas-vindas
- [x] DEV-002 — comportamento das gravações simultâneas por diferentes operadores

## Encerramento e limites

Não há pendência funcional conhecida no escopo homologado. A aprovação de DEV-002 não transforma gravações do Excel em transações atômicas; mantém-se o risco residual de concorrência e a necessidade de backup e recuperação. A aprovação de SSO não substitui permissões do Microsoft 365. A matriz anterior da Alpha 2.0 permanece como histórico; não se presumem evidências individuais adicionais além dos itens declarados acima.

A documentação técnica no GitHub está consolidada. A aprovação formal do DOC-REG-007, a revisão documental corporativa, a publicação dos documentos na pasta oficial SharePoint e a autorização explícita de distribuição/reabertura operacional (caso ainda não registrada externamente) dependem do responsável autorizado. Não houve acesso ou alteração ao SharePoint neste encerramento.

Preenchimento automático de perguntas personalizadas pertence à Alpha 3.0 e está fora desta entrega. A proposta de carro de F1/som foi descartada. Não implementar nem alterar o roadmap nesta atualização.
