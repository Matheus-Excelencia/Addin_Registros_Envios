# Revisão de segurança — Alpha 2.0

> **Atualização de situação — 27/11/2026:** este documento preserva o histórico da Alpha 2.0. O aceite funcional mais recente é o da [Alpha 2.1.5](ALPHA-215-HOMOLOGACAO-FINAL.md), com 31/31 itens declarados aprovados e 132/132 testes automatizados. SSO, fallback, segurança, interface e concorrência foram aprovados no escopo declarado. Registros antigos de pendência abaixo são históricos; não equivalem ao estado atual. Aprovação documental corporativa e autorização de distribuição/reabertura, se ainda não registrada, permanecem separadas. Riscos residuais e procedimentos de segurança continuam aplicáveis.


- Neutralização centralizada na gravação Range.values de E-mail, SMS, Config, listas, mensagens e cabeçalhos dinâmicos. Prefixos de fórmula inclusive após espaços/controles recebem apóstrofo; números finitos, datas seriais e booleanos permanecem tipados. Removido escape anterior da coleta para evitar duplicidade. Não altera dados antigos.
- Duplicação: hash SHA-256 da URL exata retornada por getFilePropertiesAsync, conferido na carga do histórico, ao duplicar e ao consumir; validade de 2 minutos, consumo único, canal correto e escopo demo separado. Falha de identidade bloqueia. URL original não persistida. O payload permanece temporariamente na sessão; hash não é criptografia de conteúdo ou autorização. Expiração impede recuperação pelo fluxo, não garante eliminação física sem nova leitura.
- Removidos objetos de erro do console nos formulários. Identificação do operador permanece na interface. Revisar mensagens de erro do Office no ambiente real. Escape HTML existente no histórico preservado. CI com permissões de leitura; manifesto e SSO mantidos.

## Bloqueios de lançamento
Testes unitários e mock não comprovam comportamento real de Range.values. Confirmar texto literal, ausência de fórmulas, apóstrofos, espaços, formatos @, números negativos, telefone e datas. Validar dois arquivos de mesmo nome, cópias, renomeações, arquivo sem URL, demo/real e sessão restaurada. Se o host não fornecer URL estável, duplicação fica indisponível; não usar nome como identidade. Preferências visuais continuam compartilhadas por origem; nomes de colunas em chaves de preferência requerem revisão separada. SSO identifica operador, não substitui autorização Microsoft 365. Validar autenticação e fallback. Não constitui pentest ou certificação.

Referências: https://learn.microsoft.com/javascript/api/excel/excel.range ; https://learn.microsoft.com/javascript/api/office/office.document

## Revisão pré-release
O manifesto mantém ID, Resource, escopos e URLs da produção. O código novo será carregado também por instalações antigas; mudança de rótulos exige atualizar XML, mas não é barreira de acesso à nova versão. Riscos não resolvidos: sem validação real de Range.values/formulas, sem prova de isolamento no host e de SSO, concorrência sem transação e backup de dados operacional não confirmado. Exigir suspensão da operação e aceite do responsável antes de testar no domínio oficial. Rollback de código não reverte dados. Não tratar a revisão como certificação.
