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

## Segurança — evidência obrigatória antes do merge
| Caso | Esperado | Evidência |
|---|---|---|
| Texto =1+1, +5511999999999, -texto, @texto, espaços e apóstrofo em registros/Config | Literal, sem fórmula; confirmar valores e fórmulas | Pendente |
| Valor -123,45, zero, data, booleano e vazio | Tipo/formato preservado | Pendente |
| Duplicar no mesmo workbook | Uma vez, revisão e novo ID | Pendente |
| Arquivos mesmo nome, URLs diferentes | Sem transferência de dados | Pendente |
| Sem identidade, sessão expirada, canal errado, demo/real | Bloqueio | Pendente |
| SSO com/sem consentimento e token | Identificação ou fallback; sem log bruto | Pendente |
| Ícones, manifesto, abertura, temas | Validação Excel Web | Pendente |

Usar dados fictícios, ambiente HTTPS isolado e SSO próprio autorizado. Registrar navegador, commit e evidência. Liberar apenas após casos críticos aprovados e autorização. O manifesto de produção não valida a branch.
