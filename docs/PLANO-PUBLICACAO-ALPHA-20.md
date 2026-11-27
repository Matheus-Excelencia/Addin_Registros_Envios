# Registra Alpha 2.0 — Plano de publicação e rollback

> **Atualização de situação — 27/11/2026:** este documento preserva o histórico da Alpha 2.0. O aceite funcional mais recente é o da [Alpha 2.1.5](ALPHA-215-HOMOLOGACAO-FINAL.md), com 31/31 itens declarados aprovados e 132/132 testes automatizados. SSO, fallback, segurança, interface e concorrência foram aprovados no escopo declarado. Registros antigos de pendência abaixo são históricos; não equivalem ao estado atual. Aprovação documental corporativa e autorização de distribuição/reabertura, se ainda não registrada, permanecem separadas. Riscos residuais e procedimentos de segurança continuam aplicáveis.


**Status: publicação controlada autorizada; uso operacional ainda bloqueado.** Marco original de preparação: NÃO AUTORIZADO PARA PUBLICAÇÃO. O usuário confirmou backup de dados e suspensão de uso. Deploy de lançamento 02097fe confirmado READY. Homologação real permanece pendente; nenhum rollback executado.

## Evidências e baseline
- Produção Alpha 1.06: commit `aa438e4dafdc6d02cd3abe9c39aaa474cf12d91c`.
- Backup Git confirmado: [backup/alpha-1.06-pre-registra](https://github.com/Matheus-Excelencia/Addin_Registros_Envios/tree/backup/alpha-1.06-pre-registra), mesmo SHA.
- Deployment Vercel anterior: `dpl_C9eoJ42UFFW6RVw4jKi9pK3Epj2x`, READY, target production; [dashboard](https://vercel.com/acme-vs-coyote/addin-registros-envios/C9eoJ42UFFW6RVw4jKi9pK3Epj2x).
- A branch Alpha 2.0 estava 76 commits à frente e 18 atrás da main. Reconciliação do histórico com o commit de produção como segundo pai, preservando a árvore revisada da branch. Alterações paralelas: README, artigo, logo, demo e fallback SSO demo já incorporados ou substituídos pela implementação mais recente; restaurado script de cabeçalho aprovado; antiga prévia substituída por redirecionamento temporário. Conferir CI e deploy do commit final.
- Último marco anterior: 105 testes aprovados (run 45921195043). Não confundir testes de mock com homologação Office.

## Critérios de entrada — antes da publicação
- [ ] Autorização explícita do responsável, definir janela fora de uso, comunicar operadores para salvar e fechar painéis, suspender gravações durante toda a validação. Não prometer isolamento de usuários no mesmo domínio: publicação troca código para todos.
- [ ] Responsável com acesso autorizado confirma cópia/versionamento recuperável do workbook SharePoint, incluindo abas Email, SMS, Config, valores, fórmulas, cabeçalhos e IDs. Não baixar dados reais para GitHub. Backup Git não protege a planilha.
- [ ] Confirmar acesso de rollback ao Vercel e GitHub, SHA e deployment de produção, sem remover referência anterior; conferir se deployment anterior continua elegível a rollback.
- [ ] PR de feat/alpha-2.0 para main revisado, sem conflitos; CI verde no exato SHA; confirmar URLs permanentes, Id do Office, permissões, Resource SSO, assets e configuração public/.
- [ ] Preparar planilha fictícia salva no SharePoint/OneDrive e segunda cópia com nome igual em pasta distinta. Não testar no workbook operacional.
- [ ] Aprovar expressamente exceção de risco de homologar no domínio oficial, com janela de manutenção. Se não houver como suspender o uso, não publicar; retomar ambiente isolado autorizado.

## Execução (somente após autorização)
1. Registrar commit candidato e backup, conferir CI e PR; mesclar sem force-push, aguardar deploy READY da main e confirmar domínio oficial no SHA esperado. Nenhuma alteração nas credenciais Entra.
2. Verificar HTTP 200 em /, /taskpane.html, /commands.html, /manifest.xml, /demo/index.html, ícones e JS. Manifesto XML 2.0.0.0, no-store; docs/tests fora do site. Verificar link da logo e redirecionamento antigo /preview-alpha20/demo/index.html.
3. Enquanto operação suspensa, testar primeiro o manifesto antigo já instalado (mesmo ID e URLs), depois instalar/atualizar 2.0.0.0 em conta de teste autorizada conforme manual; conferir rótulos e ícones. Não distribuir o XML novo aos demais antes do aceite.
4. Executar [matriz funcional e de segurança](ALPHA-20-HOMOLOGACAO.md). Exigir confirmação de gravação sem fórmula, valores negativos/zero/vazio, SSO/fallback, isolamento de duplicação, IDs únicos, ausência de perda de dados, tema e navegação. Registrar evidência sem dados pessoais.
5. Se qualquer caso crítico falhar, manter operação suspensa e reverter. Se tudo aprovado, responsável autoriza reabertura e distribuição do XML, registra SHA, hora, responsável, resultado e acompanhamento inicial.

## Rollback — critérios de acionamento
Falha de carregamento, SSO incompatível sem contingência, perda/corrupção de dados, fórmula inesperada, duplicação entre arquivos, erro de gravação, queda geral ou incompatibilidade do manifesto. Não insistir em gravações repetidas.

## Rollback — execução autorizada
1. Suspender uso, registrar evidência/IDs sem segredos; não repetir gravação. No painel Vercel, localizar deployment anterior `dpl_C9eoJ42UFFW6RVw4jKi9pK3Epj2x`, confirmar projeto e SHA, usar Instant Rollback se disponível; verificar o domínio oficial volta ao SHA de backup. Não executar rollback agora.
2. Se indisponível, criar PR de reversão do merge em main, sem reset/force-push. Verificar que a árvore da versão revertida corresponde ao backup e que Vercel publica novamente. Se necessário, reimplantar o snapshot de backup pelo fluxo autorizado. Após rollback Vercel, alinhar main por PR de reversão antes de reativar auto-assign/promover qualquer deploy, para evitar republicar código defeituoso.
3. Confirmar HTTP, abertura do painel e registro fictício com manifesto antigo. Se o XML 2.0.0.0 já foi instalado, os rótulos/ícones poderão continuar em cache; remover apenas a instalação afetada e reinstalar o XML 1.06 do backup conforme método e política do Office. A redução de versão pode exigir remoção e nova instalação. Não remover suplemento de outros usuários sem autorização.
4. Rollback do código NÃO desfaz dados do Excel. Comparar registros e histórico de versões, recuperar apenas dados comprovadamente afetados sob responsabilidade do gestor do workbook, preservando edições posteriores. Nunca restaurar automaticamente toda a planilha.
5. Reabrir operação somente após responsável aprovar a verificação; registrar incidente, preservar logs e replanejar nova release.

Referências: https://vercel.com/docs/deployments/instant-rollback ; https://learn.microsoft.com/office/dev/add-ins/publish/maintain-breaking-changes

## Registro da janela controlada — 26/11/2026
O usuário confirmou backup da planilha operacional, suspensão do uso e autorização de publicação controlada e rollback. O estado do GitHub/Vercel já mostrava o merge 02097fee6cf4ccc8a7a69908fc01772306fc8232 e deploy dpl_PtWbm6zpff9wxSfrHRMTNQmZDgbR READY. Não repetir merge; revisar o candidato atualizado antes de disponibilizar aos operadores. A confirmação de backup é declarada pelo responsável, não verificada por acesso ao SharePoint.


## Autorização do responsável — 27/11/2026

O responsável técnico confirmou sua autorização para concluir as pendências, incluindo aprovação documental do DOC-REG-008 e MAN-REG-002 e distribuição/reabertura operacional (REL-001). Aceite funcional de 31/31 itens e baseline de 132/132 testes preservados. Os documentos Word foram revisados para Alpha 2.1.5; a publicação na pasta corporativa do SharePoint ainda depende de acesso autorizado e confirmação de envio. Não alterar manifesto, SSO, versão, código funcional ou escopo da Alpha 3.0. Em cada atualização futura, revisar README, documentação técnica, manual de instalação/uso, versões Word e necessidade de republicação no SharePoint, preservando histórico e evidências.
