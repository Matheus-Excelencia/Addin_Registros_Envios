# Alpha 2.5 — Portal Web

## Objetivo
Criar uma entrada web para selecionar a planilha/modelo de trabalho e, na etapa seguinte, registrar E-mail/SMS, consultar Histórico, Resumo e Configurações sem precisar abrir o suplemento dentro do Excel Online.

## Estado desta entrega
- A branch `alpha-2.5-dev` parte de `alpha-2.0-dev`.
- `index.html` cria uma primeira tela de portal responsiva.
- O usuário pode escolher a entrada padrão do Excel ou informar um link HTTPS de planilha.
- A abertura da planilha ocorre em uma nova aba.
- Não há autenticação nem gravação de dados nesta primeira tela.

## Limitação técnica importante
As páginas atuais do suplemento usam APIs do Office/Excel para ler e gravar tabelas dentro da pasta de trabalho. Essas APIs dependem de um contexto Office e não funcionam automaticamente em um site comum. Portanto, simplesmente apontar os botões do portal para `taskpane.html`, `email.html` ou outras páginas atuais não entrega uma aplicação web funcional.

## Etapas propostas
1. Validar visualmente a página inicial da Alpha 2.5.
2. Definir como os modelos serão cadastrados e como identificar a planilha selecionada.
3. Escolher a arquitetura de dados:
   - Microsoft Graph/Excel API, com autenticação Microsoft Entra e permissões adequadas, para continuar usando arquivos Excel; ou
   - banco de dados da aplicação, caso os registros devam funcionar independentemente da planilha.
4. Implementar login real (não um formulário de login fictício) e autorização por usuário.
5. Adaptar os módulos de E-mail, SMS, Histórico, Resumo e Configurações para a arquitetura escolhida.
6. Validar em Preview com dados de teste antes de qualquer mudança na produção.

## Segurança
- Não armazenar senhas no frontend.
- Não tratar um link de planilha como autorização de acesso; a permissão precisa ser validada pelo Microsoft 365/SharePoint ou pela API escolhida.
- Não alterar `main`, `alpha-1.05` ou o domínio de produção durante o desenvolvimento.
