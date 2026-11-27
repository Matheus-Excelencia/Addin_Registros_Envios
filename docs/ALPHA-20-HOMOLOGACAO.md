# Alpha 2.0 — homologação

> **Atualização de situação — 27/11/2026:** este documento preserva o histórico da Alpha 2.0. O aceite funcional mais recente é o da [Alpha 2.1.5](ALPHA-215-HOMOLOGACAO-FINAL.md), com 31/31 itens declarados aprovados e 132/132 testes automatizados. SSO, fallback, segurança, interface e concorrência foram aprovados no escopo declarado. Registros antigos de pendência abaixo são históricos; não equivalem ao estado atual. Aprovação documental corporativa e autorização de distribuição/reabertura, se ainda não registrada, permanecem separadas. Riscos residuais e procedimentos de segurança continuam aplicáveis.

Identidade visual aprovada pelo usuário. Alpha 2.0 publicada em janela controlada; liberação operacional ainda pendente.

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

## Liberação operacional (bloqueada)
Backup, PR, aprovação, merge, deploy e retorno da demo à main. Não usar manifesto de produção para validar código de branch.

## Segurança — evidência obrigatória antes de liberar uso operacional
| Caso | Esperado | Evidência |
|---|---|---|
| Texto =1+1, +5511999999999, -texto, @texto, espaços e apóstrofo em registros/Config | Literal, sem fórmula; confirmar valores e fórmulas | Pendente |
| Valor -123,45, zero, data, booleano e vazio | Tipo/formato preservado | Pendente |
| Duplicar no mesmo workbook | Uma vez, revisão e novo ID | Pendente |
| Arquivos mesmo nome, URLs diferentes | Sem transferência de dados | Pendente |
| Sem identidade, sessão expirada, canal errado, demo/real | Bloqueio | Pendente |
| SSO com/sem consentimento e token | Identificação ou fallback; sem log bruto | Pendente |
| Ícones, manifesto, abertura, temas | Validação Excel Web | Pendente |

Usar dados fictícios e registrar navegador, commit e evidência. A homologação isolada foi bloqueada; a opção de teste no domínio oficial exige autorização explícita e janela sem uso operacional. Não liberar aos operadores antes dos casos críticos aprovados.

## Estrutura e distribuição
O código web está em `public/`; a Vercel publica seu conteúdo na raiz do domínio, sem alterar URLs existentes. O download oficial aponta para https://addin-registros-envios.vercel.app/manifest.xml, servido sem cache após publicação autorizada. Testes e documentos ficam fora da pasta publicada. Validar preview e produção separadamente; não confundir o XML da branch com o publicado.

## Nomenclatura aprovada
Suplemento: Registra. Guia Início, grupo Excelência, botão Novo registro. O workbook `Registros de Envios.xlsx`, os identificadores do Office, as URLs públicas e a configuração SSO permanecem inalterados. A versão 2.0.0.0 do XML continua reservada para a primeira publicação da Alpha 2.0.

## Ambiente de homologação preparado
Alias: https://addin-registros-envios-git-feat-alpha-20-acme-vs-coyote.vercel.app/. Preview anterior 0478910 READY, verificado por HTTP autenticado: landing, taskpane, demo, manifesto, módulo de segurança e PNG disponíveis (200); documento fora de public/ retorna 404. Manifesto separado em [homologacao/manifest-homologacao.xml](../homologacao/manifest-homologacao.xml), ID distinto, SSO desativado, identificação manual. Proteção Vercel preservada. Acesso no iframe Office, visual renderizado, validação funcional real e SSO de teste continuam pendentes. Nenhuma publicação oficial.

## Registro de andamento — 26/11/2026
O usuário informou todos os sete passos iniciais pendentes. Não há evidência de instalação ou homologação manual. Foram adicionados testes de interface simulada (tema persistente, storage indisponível, transição e redução de movimento, isolamento da abertura, ausência de SSO de teste). O resultado de CI não substitui o navegador real. A planilha fictícia foi preparada separadamente para entrega no chat; não contém dados operacionais.

## Ocorrência de homologação isolada — 26/11/2026
O usuário informou que o suplemento recusou abrir no Excel Web ao apontar para Vercel Preview. Causa provável: proteção Vercel/iframe; sem captura de erro HTTP do host Office, causa não confirmada. O usuário preferiu não insistir nesse ambiente. Não marcar como aprovado. Manter proteção Vercel e credenciais SSO existentes. Próxima alternativa: janela controlada no domínio oficial, somente com autorização, backup de dados, plano de reversão e suspensão de uso operacional. Ver [plano de publicação](PLANO-PUBLICACAO-ALPHA-20.md).

## Resultado da homologação oficial — 26/11/2026
Relato do responsável, não execução independente pelo assistente. Produção Vercel READY no SHA `99ae2e9da5b1e1842aa125fd8b4920e2a83868da` na consulta anterior a este registro documental. Checklist: 10 aprovados, 0 reprovados, 1 pendente.

| Item | Resultado informado |
|---|---|
| Instalar manifesto 2.0.0.0 | Aprovado |
| Grupo Excelência e botão Novo registro | Aprovado |
| Identificação Microsoft SSO | Pendente |
| Animação de boas-vindas | Aprovado |
| Registro de E-mail | Aprovado |
| Registro de SMS | Aprovado |
| Histórico e duplicação | Aprovado |
| Resumo e filtros | Aprovado |
| Configurações | Aprovado |
| Segurança das gravações | Aprovado |
| Temas claro e escuro | Aprovado |

A aprovação geral de segurança não fornece evidência individual de cada cenário da matriz (fórmulas, tipos, isolamento entre arquivos, expiração e canal); confirmar cobertura dos casos críticos sem expor dados de clientes. Para SSO, testar no Excel Web autenticado, no domínio oficial, com workbook fictício: abrir E-mail e SMS sem duplicar registro, conferir nome/conta automática e aviso de identificação, e confirmar fallback manual em condição de indisponibilidade quando possível. Não registrar tokens. Saudação personalizada depende do SSO, mas sua ausência isolada não comprova falha do SSO. Não liberar operadores nem executar rollback apenas por teste pendente. Aguardar confirmação de SSO, cobertura crítica e autorização final.


## Autorização do responsável — 27/11/2026

O responsável técnico confirmou sua autorização para concluir as pendências, incluindo aprovação documental do DOC-REG-008 e MAN-REG-002 e distribuição/reabertura operacional (REL-001). Aceite funcional de 31/31 itens e baseline de 132/132 testes preservados. Os documentos Word foram revisados para Alpha 2.1.5; a publicação na pasta corporativa do SharePoint ainda depende de acesso autorizado e confirmação de envio. Não alterar manifesto, SSO, versão, código funcional ou escopo da Alpha 3.0. Em cada atualização futura, revisar README, documentação técnica, manual de instalação/uso, versões Word e necessidade de republicação no SharePoint, preservando histórico e evidências.
