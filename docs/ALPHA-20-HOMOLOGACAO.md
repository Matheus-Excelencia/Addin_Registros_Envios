# Alpha 2.0 — homologação
Identidade visual aprovada pelo usuário. Produção permanece Alpha 1.06; não mesclar sem autorização.

## Desenvolvimento
- Sincronizar HTML demonstrativo: `node scripts/sync-demo.cjs`. CSS e JS compartilhados. CI verifica divergências; não publica nem escreve na main.
- Testes: `node --test tests/*.test.cjs`.

## Checklist manual
- Excel Web: instalar manifesto em ambiente de teste separado; confirmar SSO, fallback e botão Continuar.
- Email/SMS: listas, campos obrigatórios, cadastro, mensagens e valor zero/vazio.
- Histórico: filtros, duplicação sem alterar original, valor e perguntas personalizadas.
- Resumo: três modos, proporções, filtros, período vazio, volume e preferências.
- Config: criar/editar perguntas, visibilidade e persistência.
- Demo: reiniciar, restaurar dados, repetir abertura, não chamar SSO.
- Ambos os temas: contraste, responsividade e navegação.

## Publicação (bloqueada)
Backup, PR, aprovação, merge, deploy e retorno da demo à main. Não usar manifesto de produção para validar código de branch.
