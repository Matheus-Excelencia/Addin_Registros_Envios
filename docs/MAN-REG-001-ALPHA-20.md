# Manual de instalação — Alpha 2.0 (rascunho)
**Não publicado.** O manifesto 2.0.0.0 está preparado para a publicação autorizada da Alpha 2.0. O endereço permanente ainda serve a produção Alpha 1.06. Não instalar este XML para homologar o código da branch.

## Após publicação autorizada
1. Abra uma cópia de teste do workbook no Excel Web com conta Microsoft 365 autorizada.
2. Baixe o manifesto XML 2.0.0.0 da versão publicada; salve como XML.
3. Em Início > Suplementos > Avançado (ou Inserir > Suplementos), escolha Carregar meu suplemento e selecione o XML. Se bloqueado, solicite apoio ao administrador.
4. Na guia Início, grupo Registros, clique Novo Registro. Aguarde a abertura ou clique Continuar.
5. Confirme identificação Microsoft ou preencha Realizado por manualmente. Cadastre E-mail e SMS fictícios, verifique Histórico, Resumo e Configurações.
6. Confirme temas, perguntas, filtros e duplicação. Não use dados reais durante homologação.

## Desenvolvimento
Node.js 22: execute node scripts/sync-demo.cjs e node --test tests/*.test.cjs na raiz. A demo usa dados fictícios e não chama SSO. A instalação de homologação exige endpoint e configuração SSO próprios aprovados; ainda não disponibilizados.

## Identidade visual
Os ícones PNG de 16, 32 e 80 px foram atualizados na branch, seguindo a geometria e paleta do SVG da marca. Validar nitidez na faixa de opções após disponibilização de ambiente de homologação. A main permanece sem alteração nesta etapa.

## Atualizações futuras
Após a instalação do manifesto 2.0.0.0, melhorias publicadas no mesmo endereço HTTPS (HTML, CSS, JS e dados de configuração) não exigem nova instalação. Reabra o painel ou atualize o Excel Web se necessário. Ícones podem permanecer em cache; reabra o Office ou limpe o cache conforme orientação Microsoft. Alterações no próprio XML exigem incremento da versão e atualização/reimplantação conforme o método de instalação. Preserve o ID do suplemento, domínio, URLs de entrada e configuração SSO.

## Estrutura e distribuição
O código web está em `public/`; a Vercel publica seu conteúdo na raiz do domínio, sem alterar URLs existentes. O download oficial aponta para https://addin-registros-envios.vercel.app/manifest.xml, servido sem cache após publicação autorizada. Testes e documentos ficam fora da pasta publicada. Validar preview e produção separadamente; não confundir o XML da branch com o publicado.
