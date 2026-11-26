# Manual de instalação — Alpha 2.0 (rascunho)
**Não publicado.** O manifesto atual aponta para Alpha 1.06 em produção. Não usar o XML atual para homologar código da branch.

## Após publicação autorizada
1. Abra uma cópia de teste do workbook no Excel Web com conta Microsoft 365 autorizada.
2. Baixe o manifesto XML da versão publicada; salve como XML.
3. Em Início > Suplementos > Avançado (ou Inserir > Suplementos), escolha Carregar meu suplemento e selecione o XML. Se bloqueado, solicite apoio ao administrador.
4. Na guia Início, grupo Registros, clique Novo Registro. Aguarde a abertura ou clique Continuar.
5. Confirme identificação Microsoft ou preencha Realizado por manualmente. Cadastre E-mail e SMS fictícios, verifique Histórico, Resumo e Configurações.
6. Confirme temas, perguntas, filtros e duplicação. Não use dados reais durante homologação.

## Desenvolvimento
Node.js 22: execute node scripts/sync-demo.cjs e node --test tests/*.test.cjs na raiz. A demo usa dados fictícios e não chama SSO. A instalação de homologação exige endpoint e configuração SSO próprios aprovados; ainda não disponibilizados.
