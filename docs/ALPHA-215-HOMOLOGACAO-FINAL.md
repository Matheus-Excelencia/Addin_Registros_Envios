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

A documentação técnica no GitHub está consolidada. A aprovação documental e a autorização explícita de distribuição/reabertura foram concedidas pelo responsável em 27/11/2026. Publicação dos Word na pasta SharePoint pendente de acesso e confirmação; não houve alteração no SharePoint.

Preenchimento automático de perguntas personalizadas pertence à Alpha 3.0 e está fora desta entrega. A proposta de carro de F1/som foi descartada. Não implementar nem alterar o roadmap nesta atualização.


## Autorização do responsável — 27/11/2026

O responsável técnico confirmou sua autorização para concluir as pendências, incluindo aprovação documental do DOC-REG-008 e MAN-REG-002 e distribuição/reabertura operacional (REL-001). Aceite funcional de 31/31 itens e baseline de 132/132 testes preservados. Os documentos Word foram revisados para Alpha 2.1.5; a publicação na pasta corporativa do SharePoint ainda depende de acesso autorizado e confirmação de envio. Não alterar manifesto, SSO, versão, código funcional ou escopo da Alpha 3.0. Em cada atualização futura, revisar README, documentação técnica, manual de instalação/uso, versões Word e necessidade de republicação no SharePoint, preservando histórico e evidências.

## Numeração das emissões corporativas — Alpha 2.1.5

- DOC-REG-008 — Documentação Técnica Profissional, sucessora da DOC-REG-007 (Alpha 1.06).
- MAN-REG-002 — Manual de Instalação, sucessor do MAN-REG-001 (Alpha 1.06).
- Ambas as novas emissões estão preparadas em DOCX e PDF. As anteriores permanecem preservadas. Publicação no SharePoint ainda não confirmada.
- O caminho histórico deste manual Markdown no GitHub permanece estável para preservar links existentes; a numeração corporativa vigente é MAN-REG-002.

## Termo de liberação operacional — REL-REG-001

- Emissão corporativa de 27/11/2026, Alpha 2.1.5, em Word e PDF, vinculada ao [registro de liberação](REL-REG-001.md).
- Formaliza REL-001 e a autorização já concedida pelo responsável, sem constituir assinatura digital certificada.
- Arquivamento no SharePoint dos três documentos corporativos (DOC-REG-008, MAN-REG-002 e REL-REG-001) ainda não confirmado.
