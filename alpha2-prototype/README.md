# Alpha 2.0 — Ambiente de desenvolvimento

Este diretório contém exclusivamente o protótipo e a organização inicial da Alpha 2.0. O desenvolvimento ocorre na branch `alpha-2.0-dev`.

## Estrutura

- `index.html` — protótipo navegável com dados fictícios; ainda não grava no Excel nem em serviços externos.
- `assets/` — recursos visuais exclusivos da Alpha 2.0, incluindo o ícone de desenvolvimento.
- `docs/` — arquitetura, plano de desenvolvimento e documentação do protótipo.
- `tests/` — roteiros e verificações com dados fictícios.
- `manifest/` — manifestos de teste separados, a serem criados somente após confirmar uma URL de Preview própria.

## Regras de isolamento

- Não alterar `main`, `alpha-1.05` nem o deploy de produção para desenvolver a Alpha 2.0.
- Não apontar manifestos da Alpha 2.0 para a URL de produção da Alpha 1.05.
- Não usar dados pessoais reais no protótipo.
- Não gravar na planilha operacional oficial durante os testes iniciais.
- Não criar banco, serviço pago ou credenciais sem confirmar acesso, necessidade e custo.
- Não depender de privilégios administrativos do SharePoint.

## Estado atual

O protótipo visual e a documentação inicial estão no repositório. A implantação independente na Vercel ainda depende de permissão de criação/implantação no time Vercel; as tentativas atuais foram recusadas com erro 403. Nenhuma proteção foi desativada e nenhum ajuste foi feito no projeto de produção.