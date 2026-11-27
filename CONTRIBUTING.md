# Desenvolvimento e contribuição

## Fluxo de trabalho
- Desenvolver em branch separada. Não publicar ou mesclar na main sem aprovação.
- Usar Node.js 22. Não é necessário instalar dependências para os testes atuais.
- Editar o suplemento em `public/`. Não modificar HTML gerado de `public/demo/` manualmente; executar `node scripts/sync-demo.cjs` na raiz.
- Validar `node scripts/sync-demo.cjs --check` e `node --test tests/*.test.cjs` antes de abrir PR.
- Não incluir dados reais de clientes, tokens ou credenciais. Usar dados fictícios.
- Preservar as URLs públicas, os identificadores Office e SSO. `public/manifest.xml` é publicado como `/manifest.xml`; o diretório `public` não integra a URL.
- Validar a matriz de homologação no Excel Web e o preview da Vercel antes de liberar produção.

## Releases
O botão da landing page aponta para o manifesto do domínio oficial. Não vincular a branch experimental ao download de produção. O cache HTTP do XML deve permanecer desabilitado. A atualização web não exige reinstalação do XML se o contrato Office não mudar. Consulte [política de atualização](docs/POLITICA-ATUALIZACAO.md).
