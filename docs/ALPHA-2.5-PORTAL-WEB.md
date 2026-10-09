# Alpha 2.5 — Portal web + suplemento

## Estado desta etapa
- A página inicial do portal está em `index.html`.
- Os estilos e o comportamento da tela estão em `app/styles/portal.css` e `app/scripts/portal.js`.
- O usuário pode escolher um módulo, filtrar atalhos, colar um link HTTPS permitido e abrir a planilha em outra aba.
- O suplemento Office continua preservado e usa suas páginas existentes.

## Limitações importantes
- Este portal ainda não possui autenticação real; não simular login nem armazenar senhas.
- A seleção de planilha é manual por link. A lista automática exige Microsoft Entra ID/MSAL e Microsoft Graph.
- Os formulários atuais foram construídos com Office.js e não são garantidos como independentes do Excel.
- Os botões do portal apenas navegam para os módulos existentes. Não considerar gravação via navegador validada até adaptar a camada de dados.
- Não usar dados pessoais reais nos testes.

## Próximos passos técnicos
1. Mapear as tabelas, colunas e operações que os módulos atuais usam.
2. Definir o esquema de cada modelo de planilha suportado.
3. Configurar aplicativo Microsoft Entra ID e redirect URI de Preview; nunca incluir segredo no frontend.
4. Implementar login MSAL e seleção/listagem de arquivos via Graph com permissões mínimas.
5. Implementar leitura e gravação via Graph somente para operações suportadas pelo esquema validado.
6. Oferecer fallback explícito para abrir no Excel quando a operação exigir Office.js.
7. Testar links inválidos, acesso negado, arquivo incompatível, sessão expirada, conflitos e gravação em cópia de teste.
8. Manter produção, `main` e `alpha-1.05` intocados até aprovação explícita.

## Critério de aceite
O portal só poderá ser anunciado como alternativa completa ao suplemento quando o usuário autenticado puder selecionar um arquivo autorizado, validar seu modelo e concluir as operações suportadas sem Excel aberto, com mensagens de erro e confirmação confiáveis.
