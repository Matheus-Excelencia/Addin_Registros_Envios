# Sistema Evolut — suplemento e portal web

Este repositório reúne o suplemento do Excel e a evolução de um portal web complementar.

## Branches de trabalho
- **Sistema Evolut:** `sistema-evolut` — organização e evolução do suplemento Office.
- **Portal WEB Evolut:** `portal-web-evolut` — portal web complementar para selecionar modelo/planilha, acessar os módulos e abrir a planilha opcionalmente.
- **Branches legadas:** `alpha-2.0-dev` e `alpha-2.5-dev` permanecem temporariamente para preservar deploys e links até atualizar as integrações.

O GitHub não aceita espaços nos nomes técnicos das branches; por isso usamos hífens. Os nomes apresentados ao usuário continuam “Sistema Evolut” e “Portal WEB Evolut”.

## Aplicação do suplemento
As páginas HTML e o manifesto permanecem na raiz. CSS fica em `app/styles/` e JavaScript em `app/scripts/`. Não mova páginas nem altere rotas do manifesto sem validar hospedagem e navegação.

## Portal WEB Evolut
O portal é uma experiência complementar, não uma substituição do suplemento. O objetivo é oferecer:
1. Seleção do modelo/planilha de trabalho.
2. Campo opcional para informar o link da planilha e abri-la.
3. Acesso aos módulos de registro de E-mail e SMS, histórico, resumo e configurações, respeitando o modelo selecionado.
4. Autenticação e integração segura com os dados antes de disponibilizar operações reais.

## Documentação
- [Roadmap e mapa de dependências](docs/PLANO-ORGANIZACAO.md)
- [Checklist de validação](docs/CHECKLIST-VALIDACAO.md)
- [Plano do Portal WEB Evolut](docs/ALPHA-2.5-PORTAL-WEB.md)
- [Protótipo isolado](alpha2-prototype/README.md)

## Segurança
- Não modificar `main`, `alpha-1.05` ou produção sem autorização explícita.
- Não instalar no Excel o manifesto do Preview enquanto ele usar GUID/URLs de produção.
- Não usar dados reais até autenticação, autorização, leitura/gravação e validação serem testadas.
