# Checklist de validação da reorganização

Marcar cada item somente depois de executar e registrar o resultado.

## Antes da migração
- [x] Confirmar branch e commit de trabalho: `alpha-2.0-dev`.
- [x] Confirmar que `main` e `alpha-1.05` não foram alteradas por esta etapa.
- [x] Registrar os caminhos principais antes/depois da etapa de movimentação.
- [x] Buscar referências principais a `.html`, `.css`, `.js`, `window.location` e `location.href`.

## Verificação estática
- [x] HTMLs da aplicação apontam para CSS/JS existentes em `app/styles/` e `app/scripts/`.
- [x] Caminhos de navegação verificados nos scripts consultados apontam para páginas HTML mantidas na raiz.
- [ ] Conferir todos os ícones e recursos referenciados no manifesto.
- [ ] Validar formalmente a sintaxe do XML do manifesto.
- [ ] Confirmar que o manifesto de teste usa ID e URLs independentes da produção.
- [x] O manifesto estável consultado continua apontando para a URL de produção; ele não deve ser usado para testar a Alpha 2.0.
- [ ] Fazer uma revisão funcional completa para confirmar que a movimentação não alterou gravação, planilhas ou validações.

## Deploy e Excel
- [ ] Existe URL Preview exclusiva e acessível por HTTPS.
- [ ] Abrir cada página diretamente no Preview sem erro 404.
- [ ] Confirmar carregamento de CSS, JS e Office.js no navegador.
- [ ] Importar somente o manifesto de teste independente.
- [ ] Testar navegação: painel → novo registro → e-mail/SMS; painel → histórico/resumo/configurações.
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
- [ ] Validação em Preview e no Excel.

## Bloqueio registrado
A criação de um projeto Vercel separado para a Alpha 2.0 retornou erro 403 (sem permissão para criar projeto). Não foi possível confirmar um Preview independente. Até resolver isso, não importar o manifesto estável no Excel para validar a Alpha 2.0 e não alterar produção.
