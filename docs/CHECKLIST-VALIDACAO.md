# Checklist de validação — Sistema Evolut

## Organização e Preview
- [x] CSS e JavaScript da aplicação organizados em `app/styles/` e `app/scripts/`.
- [x] Referências locais verificadas por requisição HTTP no Preview.
- [x] Ícones 16, 32 e 80 px respondem HTTP 200 no Preview.
- [x] Usuário confirmou que todas as telas abriram normalmente.
- [ ] Investigar/definir o endpoint raiz `/`, que responde HTTP 404 no Preview.
- [ ] Validar formalmente a sintaxe XML do manifesto.
- [ ] Criar manifesto de teste com ID e URLs independentes da produção.
- [ ] Testar Office.js e navegação dentro do Excel.
- [ ] Testar leitura e gravação em cópia de planilha.
- [ ] Testar filtros, duplicação, mensagens, configurações e validações.

## Portal WEB Evolut
- [ ] Página inicial web com identidade Portal WEB Evolut.
- [ ] Escolha explícita do modelo/planilha.
- [ ] Campo opcional para link da planilha e ação para abri-la.
- [ ] Navegação para E-mail, SMS, histórico, resumo e configurações conforme o modelo escolhido.
- [ ] Definir autenticação e autorização reais antes de dados reais.
- [ ] Definir integração segura para leitura/gravação (por exemplo, Microsoft Graph com consentimento apropriado ou backend autorizado).
- [ ] Testar permissões, erros, sessão e acesso indevido.
- [ ] Confirmar que o suplemento Excel continua funcionando e disponível.

## Branches e segurança
- [x] Criada branch `sistema-evolut` a partir do estado de `alpha-2.0-dev`.
- [x] Criada branch `portal-web-evolut` a partir do estado de `alpha-2.5-dev`.
- [ ] Atualizar Vercel e demais integrações para usar as novas branches.
- [ ] Verificar PRs, links e automações antes de excluir branches antigas.
- [ ] Excluir `alpha-2.0-dev` e `alpha-2.5-dev` somente depois da migração dos vínculos e confirmação do usuário.
- [x] Não alterar `main`, `alpha-1.05` nem produção.

## Estado atual
A abertura visual das telas foi confirmada pelo usuário. Isso não comprova, por si só, que a gravação, leitura de dados, Office.js, autenticação ou integração com planilhas estejam funcionais. O manifesto servido no Preview ainda aponta para GUID e URLs de produção; não instalá-lo como manifesto de teste.
