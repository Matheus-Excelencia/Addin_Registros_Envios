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


## Autorização do responsável — 27/11/2026

O responsável técnico confirmou sua autorização para concluir as pendências, incluindo aprovação documental do DOC-REG-008 e MAN-REG-002 e distribuição/reabertura operacional (REL-001). Aceite funcional de 31/31 itens e baseline de 132/132 testes preservados. Os documentos Word foram revisados para Alpha 2.1.5; a publicação na pasta corporativa do SharePoint ainda depende de acesso autorizado e confirmação de envio. Não alterar manifesto, SSO, versão, código funcional ou escopo da Alpha 3.0. Em cada atualização futura, revisar README, documentação técnica, manual de instalação/uso, versões Word e necessidade de republicação no SharePoint, preservando histórico e evidências.
