# Checklist de validação da reorganização

Marcar cada item somente depois de executar e registrar o resultado.

## Antes da migração
- [ ] Confirmar branch e commit de trabalho: `alpha-2.0-dev`.
- [ ] Confirmar que `main` e `alpha-1.05` não foram alteradas.
- [ ] Registrar o conjunto de arquivos antes de cada etapa.
- [ ] Buscar referências a `.html`, `.css`, `.js`, `assets/`, `window.location`, `location.href` e URLs absolutas.

## Verificação estática
- [ ] Todo HTML aponta para CSS e JS existentes.
- [ ] Todos os caminhos de navegação levam a páginas existentes.
- [ ] Todos os ícones e recursos referenciados existem.
- [ ] O manifesto XML é bem-formado.
- [ ] IDs e URLs do manifesto de teste são independentes da produção.
- [ ] Nenhum recurso de teste aponta acidentalmente para a URL de produção.
- [ ] Não há alterações em lógica de gravação, planilhas ou validações apenas por causa da movimentação de arquivos.

## Deploy e Excel
- [ ] Existe URL Preview exclusiva e acessível por HTTPS.
- [ ] Abrir cada página diretamente no Preview sem erro 404.
- [ ] Confirmar carregamento de CSS, JS e Office.js.
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

## Estado inicial
- [x] Inventário de arquivos das branches `main`, `alpha-1.01` a `alpha-1.05` e `alpha-2.0-dev`.
- [x] Mapeamento inicial dos caminhos principais em `alpha-2.0-dev`.
- [ ] Migração física dos arquivos.
- [ ] Validação no preview e no Excel.
