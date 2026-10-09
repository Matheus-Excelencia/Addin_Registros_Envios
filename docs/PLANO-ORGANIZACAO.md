# Sistema Evolut — mapa de dependências e roadmap

## Nomenclatura oficial
- **Nome do produto/módulo:** Sistema Evolut.
- **Branch de desenvolvimento:** `sistema-evolut`.
- **Branch anterior/legada:** `alpha-2.0-dev` (mantida temporariamente como referência até a troca oficial de links e integrações).
- **Portal web complementar:** Portal WEB Evolut, na branch `portal-web-evolut`.
- **Branch anterior/legada do portal:** `alpha-2.5-dev`.

> O GitHub não permite espaços em nomes de branches. Por isso, os nomes técnicos usam hífens; os nomes exibidos no produto/documentação preservam “Sistema Evolut” e “Portal WEB Evolut”.

## Objetivo
Organizar e evoluir o suplemento Office sem alterar o comportamento existente, mantendo a opção de usar o sistema dentro do Excel e preparando um portal web para usuários que prefiram começar pelo navegador.

## Estrutura atual do suplemento
- `manifest.xml`, páginas HTML e `assets/` permanecem na raiz para manter os endpoints usados pelo manifesto.
- CSS em `app/styles/`.
- JavaScript em `app/scripts/`.
- As páginas HTML carregam os caminhos novos e os scripts navegam entre páginas na raiz.
- `commands.html` contém a inclusão de `office.js`; não referencia `commands.js` local.
- O manifesto publicado no Preview ainda aponta para o GUID e domínio de produção. **Não instalar esse manifesto do Preview no Excel para validar mudanças.**
- O protótipo isolado permanece em `alpha2-prototype/`.

## Roadmap
### Fase 1 — Organização e Preview (em andamento)
1. Documentar a estrutura e os riscos de rotas.
2. Organizar CSS/JS em `app/styles/` e `app/scripts/`.
3. Conferir referências locais e caminhos de navegação.
4. Verificar os recursos publicados no Preview.
5. Registrar o resultado de teste manual de abertura das telas: usuário confirmou que todas abriram normalmente.

### Fase 2 — Validação funcional do Sistema Evolut
1. Validar formalmente o XML do manifesto.
2. Criar e validar manifesto de teste com ID e URLs separados da produção.
3. Testar Office.js dentro do Excel com cópia de uma pasta de trabalho.
4. Validar gravação, leitura, filtros, duplicação, mensagens, configurações e navegação.
5. Corrigir e documentar falhas encontradas.
6. Não promover para a linha estável até aprovação explícita.

### Fase 3 — Portal WEB Evolut (desenvolvimento complementar)
1. Criar uma página principal web com nome e identidade do Portal WEB Evolut.
2. Permitir escolher o modelo/planilha que será usado.
3. Oferecer campo opcional para colar o link da planilha e botão para abri-la.
4. Apresentar os módulos de registro (E-mail, SMS, histórico, resumo e configurações) conforme o modelo escolhido.
5. Manter o suplemento Excel e o portal web como duas formas de acesso ao mesmo produto, sem remover o suplemento.
6. Definir autenticação real e autorização de acesso antes de usar dados reais; não tratar um login visual como segurança.
7. Decidir como o portal lerá/gravarará os dados: integração Microsoft Graph/Excel com permissões consentidas ou backend autorizado. Não simular persistência como se fosse real.
8. Prototipar com dados fictícios até a estratégia de autenticação, armazenamento e custos estar aprovada.

### Fase 4 — Integração e consistência
1. Definir modelos/planilhas suportados e validar suas colunas.
2. Compartilhar regras de validação entre o suplemento e o portal quando possível.
3. Testar ambos os caminhos de uso separadamente e em conjunto.
4. Documentar instalação do suplemento, acesso ao portal e requisitos de permissões.
5. Avaliar migração de branches antigas individualmente, sem sobrescrever diferenças de versões.

### Fase 5 — Promoção segura
1. Comparar `alpha-1.01` a `alpha-1.05` individualmente.
2. Preservar as particularidades funcionais de cada versão.
3. Tratar `main` e `alpha-1.05` por último, apenas após validação e autorização explícita.
4. Não alterar produção, manifesto estável ou deploy de produção durante as fases anteriores.

## Estrutura resumida
```text
/
├── manifest.xml
├── taskpane.html, novo.html, email.html, sms.html
├── historico.html, resumo.html, config.html, commands.html
├── assets/
├── app/
│   ├── styles/
│   └── scripts/
├── alpha2-prototype/
└── docs/
    ├── PLANO-ORGANIZACAO.md
    ├── CHECKLIST-VALIDACAO.md
    └── ALPHA-2.5-PORTAL-WEB.md
```

## Situação e bloqueios
- Preview da branch antiga `alpha-2.0-dev` responde para `taskpane.html`, `manifest.xml`, CSS/JS consultados e ícones.
- A raiz `/` responde HTTP 404 no Preview; páginas HTML diretas funcionam. O usuário confirmou que todas as telas abriram normalmente.
- Ainda faltam validação funcional no Excel e comprovação de leitura/gravação em uma cópia de planilha.
- O manifesto servido pelo Preview contém URLs e GUID de produção; não usá-lo para instalação de teste.
- O Portal WEB Evolut é um fluxo complementar planejado; login real, autorização, integração de dados e persistência precisam ser implementados e testados.

## Política de branches e rastreabilidade
- O novo nome técnico de desenvolvimento é `sistema-evolut`; o portal é `portal-web-evolut`.
- As branches legadas `alpha-2.0-dev` e `alpha-2.5-dev` permanecem por enquanto para evitar quebrar links, deploys e referências.
- Antes de excluir as branches antigas, atualizar configurações da Vercel, links da documentação, automações e PRs; depois confirmar que os novos nomes estão funcionando.
- `main`, `alpha-1.05` e produção não foram modificadas por esta renomeação.
