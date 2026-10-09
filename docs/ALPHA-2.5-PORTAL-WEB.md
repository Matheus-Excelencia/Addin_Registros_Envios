# Portal WEB Evolut

## Objetivo
Criar uma entrada web complementar ao suplemento Excel, permitindo selecionar o modelo/planilha de trabalho e, na etapa seguinte, registrar E-mail/SMS, consultar Histórico, Resumo e Configurações sem precisar abrir o suplemento dentro do Excel Online.

## Branches e nomes
- Nome do produto: **Portal WEB Evolut**.
- Branch técnica atual: `portal-web-evolut`.
- Branch legada: `alpha-2.5-dev`, mantida temporariamente até atualizar deploys, links e integrações.
- O Sistema Evolut (suplemento) segue em `sistema-evolut`; as duas experiências devem coexistir.

## Estado desta entrega
- `index.html` cria uma primeira tela de portal responsiva.
- O usuário pode escolher a entrada padrão do Excel ou informar um link HTTPS de planilha.
- A abertura da planilha ocorre em uma nova aba.
- Não há autenticação nem gravação de dados nesta primeira tela.
- A abertura visual da aplicação foi validada manualmente pelo usuário, mas não equivale a validação de persistência nem de acesso a dados.

## Limitação técnica importante
As páginas atuais do suplemento usam APIs do Office/Excel para ler e gravar tabelas dentro da pasta de trabalho. Essas APIs dependem de um contexto Office e não funcionam automaticamente em um site comum. Portanto, simplesmente apontar os botões do portal para `taskpane.html`, `email.html` ou outras páginas atuais não entrega uma aplicação web funcional.

## Etapas propostas
1. Evoluir a página inicial com identidade visual Portal WEB Evolut.
2. Permitir escolher e identificar o modelo de trabalho.
3. Manter um campo opcional para colar o link da planilha e abri-la.
4. Definir como os modelos serão cadastrados e como identificar a planilha selecionada.
5. Escolher a arquitetura de dados:
   - Microsoft Graph/Excel API, com autenticação Microsoft Entra e permissões adequadas, para continuar usando arquivos Excel; ou
   - banco de dados da aplicação, caso os registros devam funcionar independentemente da planilha.
6. Implementar login real (não um formulário de login fictício) e autorização por usuário.
7. Adaptar os módulos de E-mail, SMS, Histórico, Resumo e Configurações para a arquitetura escolhida.
8. Preservar o suplemento Excel como opção paralela de uso; não remover nem substituir sua entrada.
9. Validar em Preview com dados de teste antes de qualquer mudança na produção.

## Segurança
- Não armazenar senhas no frontend.
- Não tratar um link de planilha como autorização de acesso; a permissão precisa ser validada pelo Microsoft 365/SharePoint ou pela API escolhida.
- Não alterar `main`, `alpha-1.05` ou o domínio de produção durante o desenvolvimento.
