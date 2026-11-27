# Política de atualização do suplemento

Manifesto de produção preparado: 2.0.0.0 (Alpha 2.0). Este número representa o contrato do manifesto, não a versão de cada atualização web. A publicação da Alpha 2.0 continua sujeita a autorização.

## Contrato estável
- ID do suplemento: 8f4a2e7d-5c71-4b1a-9f36-3b8f0c7d2146.
- Origem HTTPS: https://addin-registros-envios.vercel.app.
- Entradas: /taskpane.html e /commands.html. Não apontar para URLs temporárias de deploy ou para a demo.
- Ícones: /assets/icon-16.png, /assets/icon-32.png, /assets/icon-80.png.
- Preservar ReadWriteDocument, identidade do aplicativo Microsoft, escopos SSO e requisitos Office até revisão explícita.

## Releases web
1. Desenvolver e testar em branch; validar CI, segurança e homologação.
2. Após aprovação, publicar a versão no domínio permanente. Não alterar o manifesto apenas para acompanhar a versão do produto.
3. O Excel carrega os recursos web atualizados; reabrir o painel/atualizar Excel Web pode ser necessário. Não interromper sessões em uso ou forçar recarregamento durante gravações.
4. Validar produção e manter possibilidade de rollback do deploy. Ícones podem ficar no cache do Office.

## Alterações no XML
Incrementar a versão do manifesto e distribuir o XML atualizado quando houver mudança no contrato, como rótulos de comandos, permissões, escopos, requisitos ou endereços. A implantação centralizada pode exigir consentimento administrativo. Não prometer atualização automática do próprio XML instalado manualmente. Não modificar ID ou Resource SSO sem planejamento de migração.

## Homologação
O domínio permanente ainda aponta para a produção atual até publicação autorizada. Um ambiente de teste real exige origem HTTPS e configuração Microsoft próprias; a demo não valida SSO.

Referência: https://learn.microsoft.com/office/dev/add-ins/publish/maintain-breaking-changes

## Estrutura e distribuição
O código web está em `public/`; a Vercel publica seu conteúdo na raiz do domínio, sem alterar URLs existentes. O download oficial aponta para https://addin-registros-envios.vercel.app/manifest.xml, servido sem cache após publicação autorizada. Testes e documentos ficam fora da pasta publicada. Validar preview e produção separadamente; não confundir o XML da branch com o publicado.

## Nomenclatura aprovada
Suplemento: Registra. Guia Início, grupo Excelência, botão Novo registro. O workbook `Registros de Envios.xlsx`, os identificadores do Office, as URLs públicas e a configuração SSO permanecem inalterados. A versão 2.0.0.0 do XML continua reservada para a primeira publicação da Alpha 2.0.
