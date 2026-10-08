# Checklist de validação da reorganização

Marcar cada item somente depois de executar e registrar o resultado.

## Antes da migração
- [x] Confirmar branch e commit de trabalho: `alpha-2.0-dev`.
- [x] Confirmar que `main` e `alpha-1.05` não foram alteradas por esta etapa.
- [x] Registrar os caminhos principais antes/depois da etapa de movimentação.
- [x] Buscar referências principais a `.html`, `.css`, `.js`, `window.location` e `location.href`.

## Verificação estática e de hospedagem
- [x] HTMLs da aplicação apontam para CSS/JS existentes em `app/styles/` e `app/scripts/`.
- [x] Caminhos de navegação nos scripts consultados apontam para páginas HTML mantidas na raiz.
- [x] CSS e JavaScript locais consultados respondem HTTP 200 no Preview.
- [x] Ícones `assets/icon-16.png`, `assets/icon-32.png` e `assets/icon-80.png` respondem HTTP 200 no Preview.
- [ ] Validar formalmente a sintaxe XML do manifesto.
- [ ] Confirmar que existe manifesto de teste independente, com ID e URLs independentes da produção.
- [x] Confirmado que o manifesto publicado neste deploy ainda usa o GUID estável e URLs de produção; não usar para instalar/testar a Alpha 2.0 no Excel.
- [ ] Fazer revisão funcional completa de gravação, planilhas e validações.
- [ ] Investigar o endpoint raiz `/`, que responde HTTP 404 no Preview; as páginas HTML diretas funcionam.

## Deploy e Excel
- [x] URL Preview HTTPS disponível: https://addin-registros-envios-nipm8pr6r-acme-vs-coyote.vercel.app/
- [x] Abrir diretamente as páginas principais no Preview: o usuário informou que todas as telas abriram normalmente.
- [x] Confirmar por requisição HTTP que `taskpane.html`, `manifest.xml`, os CSS/JS locais consultados e os ícones respondem corretamente.
- [ ] Confirmar carregamento de Office.js no Excel/navegador.
- [ ] Importar somente manifesto de teste independente.
- [x] Teste manual informado pelo usuário: todas as telas abriram normalmente.
- [ ] Testar fluxo completo de navegação dentro do Excel.
- [ ] Testar leitura e gravação em cópia de uma pasta de trabalho, nunca em dados reais durante a validação.
- [ ] Testar filtros, duplicação, retorno, mensagens e validações.
- [ ] Registrar erros do console e corrigir antes de avançar.

## Promoção às outras branches
- [ ] Comparar arquivos de cada versão antes de aplicar mudanças.
- [ ] Não substituir funcionalidades antigas por versões mais novas sem análise.
- [ ] Aplicar apenas a reorganização e correções de caminho necessárias.
- [ ] Validar cada branch individualmente.
- [ ] Atualizar `main` e `alpha-1.05` por último, somente após aprovação explícita e plano de reversão.

## Estado atual
- [x] Inventário inicial das branches `main`, `alpha-1.01` a `alpha-1.05` e `alpha-2.0-dev`.
- [x] Documentação de organização e mapa inicial de dependências.
- [x] Migração física dos CSS/JS para `app/styles/` e `app/scripts/` na branch `alpha-2.0-dev`.
- [x] Atualização das referências nos sete HTMLs da aplicação que carregam CSS/JS locais.
- [x] Deploy Preview acessível e teste manual de abertura de todas as telas informado pelo usuário.
- [ ] Validação funcional no Excel e confirmação de persistência dos dados.
- [ ] Validar manifesto de teste separado antes de instalar no Excel.

## Bloqueios e observações
- A tentativa anterior de criar um projeto Vercel separado para Alpha 2.0 retornou erro 403. A integração agora permite consultar deployments do projeto existente e foi localizado um Preview da branch `alpha-2.0-dev`.
- O Preview consultado serve `taskpane.html`, `manifest.xml`, CSS/JS locais e ícones, mas `/` retorna HTTP 404.
- O manifesto publicado no Preview ainda aponta para o domínio de produção e mantém o GUID estável. Não importar esse manifesto no Excel para validar a Alpha 2.0.
- A abertura visual das telas não comprova que salvar, ler dados, Office.js, filtros ou validações funcionem.
