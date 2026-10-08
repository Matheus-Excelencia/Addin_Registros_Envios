# Mapa de dependências e plano de reorganização

## Objetivo
Organizar o código sem alterar o comportamento funcional do suplemento. A migração de CSS/JavaScript foi feita exclusivamente em `alpha-2.0-dev`; `main` e `alpha-1.05` continuam fora desta etapa.

## Estrutura atual verificada em alpha-2.0-dev
- `manifest.xml`, páginas HTML e `assets/` permanecem na raiz para manter os endpoints usados pelo manifesto.
- Os CSS da aplicação estão em `app/styles/`.
- Os scripts JavaScript da aplicação estão em `app/scripts/`.
- Os sete HTMLs que carregam CSS/JS locais foram atualizados para os novos caminhos.
- `commands.html` contém apenas a inclusão de `office.js`; não referencia um `commands.js` local.
- Os scripts consultados navegam para páginas HTML na raiz (por exemplo, `taskpane.html`, `email.html` e `historico.html`), então essas páginas não devem ser movidas até haver uma estratégia de rotas testada.
- O manifesto consultado continua com o GUID estável e a URL de produção `https://addin-registros-envios.vercel.app/`. Não usar esse manifesto para validar a Alpha 2.0.
- O protótipo isolado está em `alpha2-prototype/`.

## Estrutura atual resumida
```text
/
├── README.md
├── manifest.xml
├── taskpane.html
├── novo.html
├── email.html
├── sms.html
├── historico.html
├── resumo.html
├── config.html
├── commands.html
├── assets/
├── app/
│   ├── README.md
│   ├── styles/
│   └── scripts/
├── alpha2-prototype/
└── docs/
    ├── PLANO-ORGANIZACAO.md
    └── CHECKLIST-VALIDACAO.md
```

## Etapas concluídas
1. Documentar a estrutura e os riscos de migrar páginas e URLs.
2. Mover CSS/JavaScript para `app/styles/` e `app/scripts/` na branch `alpha-2.0-dev`.
3. Atualizar os caminhos locais nos sete HTMLs principais.
4. Fazer conferência estática dos caminhos locais de CSS/JS e dos caminhos de navegação consultados.
5. Atualizar o checklist com o que foi conferido e o que permanece pendente.

## Próximas etapas seguras
1. Validar os recursos do manifesto e a sintaxe XML.
2. Conseguir acesso ao projeto Vercel correto ou permissão para criar um projeto Preview independente.
3. Testar a branch em Preview HTTPS: abrir páginas diretamente, conferir CSS/JS/Office.js e navegação.
4. Testar com manifesto independente e cópia de uma pasta de trabalho, sem dados reais.
5. Comparar individualmente as branches `alpha-1.01` a `alpha-1.05` antes de replicar a organização, preservando particularidades de cada versão.
6. Tratar `main` e `alpha-1.05` por último, somente após validação e autorização explícita.

## Bloqueio atual
A tentativa de criar o projeto Vercel independente para Alpha 2.0 retornou erro 403 por falta de permissão. A conexão disponível também não listou deployments do projeto esperado. Não existe URL Preview independente confirmada nesta etapa; por isso, a validação de hospedagem/Excel está pendente.

## Limites desta verificação
A conferência de caminhos é estática: ela não comprova que a hospedagem serve os arquivos corretamente nem que o suplemento funciona no Excel. Ainda faltam a validação formal do XML, a conferência integral de recursos e os testes funcionais em Preview.
