# MAN-REG-001 — Manual de instalação, atualização e utilização

**Produto:** Registra · Alpha 2.1.5 | **Manifesto:** 2.0.0.0 | **Ambiente:** Excel Web / Microsoft 365 | **Situação:** homologação funcional aprovada pelo responsável em 27/11/2026; revisão e publicação documental corporativa pendentes.

## 1. Requisitos
Conta Microsoft 365 autorizada, permissão de edição no workbook e política que permita suplemento personalizado. Navegador atualizado. O suplemento não dispara E-mail ou SMS: registra envios já realizados. Não exige banco próprio.

## 2. Instalação após liberação
1. Acessar https://addin-registros-envios.vercel.app/ e clicar **Baixar manifesto XML**. Confirmar que a versão liberada é 2.0.0.0; salvar como `manifest.xml`, sem extensão .txt. O domínio oficial entregava a Alpha 2.1.5 no baseline verificado.
2. Abrir o workbook autorizado no Excel Web. Para homologação, usar somente a cópia fictícia salva no OneDrive/SharePoint.
3. Em Início > Suplementos > Avançado ou Inserir > Suplementos (nomes variam), escolher Carregar meu suplemento e selecionar o XML. Se bloqueado, consultar administrador; não contornar política corporativa.
4. Na guia Início, localizar **Excelência > Novo registro**. Aguardar tela de boas-vindas ou clicar Continuar. Confirmar nome Registra.
5. Confirmar identificação Microsoft; se indisponível, preencher Realizado por manualmente conforme procedimento interno. Nunca inserir senha/token no formulário.

## 3. Usuários existentes
O mesmo ID e endereço permanecem. O código web muda quando o deploy oficial é publicado, inclusive para instalações antigas. O XML novo atualiza nome, grupo, botão e ícones; pode exigir atualização/reinstalação conforme método de instalação ou ação administrativa. Fechar o painel e atualizar Excel Web após liberação. Não remover o suplemento enquanto houver gravação em curso. Se o cache mantiver ícones/rótulos antigos, seguir orientação do suporte Office.

## 4. Uso básico
- Novo registro: escolher E-mail ou SMS, preencher campos obrigatórios, quantidade e mensagem; Valor é opcional por padrão. Salvar uma vez e aguardar confirmação/ID.
- Histórico: consultar filtros, abrir linha e duplicar no mesmo arquivo; revisar os dados antes de salvar. Duplicação expira em dois minutos e pode ficar indisponível se Office não identificar o workbook.
- Resumo: filtrar período e canal, escolher visualização e perguntas exibidas.
- Configurações: alterar listas, mensagens e perguntas somente conforme autorização; não modificar cabeçalhos originais ou IDs sem revisão.
- Tema claro/escuro é salvo no navegador. Animações respeitam movimento reduzido do sistema.

## 5. Evoluções e suporte
Atualizações de HTML/CSS/JS no domínio permanente não exigem XML novo. Mudança de permissões, comandos ou outro contrato do manifesto requer revisão e atualização. Em caso de erro de gravação, anotar ID/status e conferir a linha antes de repetir. Não compartilhar dados de clientes em chamados. Se o suplemento não abrir, atualizar a página e contactar suporte; não desativar segurança do navegador/Vercel.

## 6. Homologação e aprovação
Preview protegido recusou abrir dentro do Excel Web segundo relato do usuário; não homologado. Teste controlado no domínio oficial somente após autorização, suspensão de uso e backup de dados. [Plano de publicação e rollback](PLANO-PUBLICACAO-ALPHA-20.md) e [matriz de testes](ALPHA-20-HOMOLOGACAO.md). Registro histórico do ambiente isolado. Aceite funcional posterior: 27/11/2026, conforme declaração do responsável; consulte [homologação final](ALPHA-215-HOMOLOGACAO-FINAL.md). Aprovação documental e autorização operacional são distintas.

## 7. Configurações e preferências pessoais

1. Para cadastrar empresas, e-mails de resposta, supervisores ou realizadores, acesse Configurações e selecione o tipo de cadastro. Informe o valor e confirme. Valores já cadastrados não são duplicados.
2. Ao salvar um novo registro de E-mail ou SMS, valores novos dos campos correspondentes também são incluídos na Config. E-mail de resposta aplica-se somente ao formulário E-mail.
3. Ao cadastrar mensagens, selecione E-mail ou SMS, informe nome e texto e salve; respeite o limite configurado para SMS. Os pares nome/texto permanecem associados. As listas são ordenadas alfabeticamente e os intervalos atualizados ficam sem quebra automática de texto. Parágrafos intencionais são preservados. Edição manual na planilha não dispara essa rotina.
4. Em Configurações, crie ou edite perguntas personalizadas, definindo formulário, tipo, obrigatoriedade e formato compatível. A pergunta Valor é opcional por padrão. Não modifique cabeçalhos e IDs manualmente.
5. Em Ordem dos formulários, selecione E-mail ou SMS, mova os campos pelas setas e salve ou restaure a ordem. A ordem visual não altera a coluna de destino.
6. Nas preferências de visibilidade, escolha Não exibir, Opcional ou Sempre exibir para Histórico e Resumo. Essas preferências e a ordem são pessoais por identidade Microsoft e arquivo no navegador; sem SSO são temporárias por sessão, e sem identidade de arquivo ou armazenamento ficam em memória. Não há sincronização entre dispositivos. Ao trocar de conta sem SSO, feche a aba.
7. As perguntas e os cadastros da Config continuam compartilhados entre operadores autorizados. Confira o log de conexão e os avisos de validação antes de gravar.

## 8. Site e demonstração

A página inicial e o artigo possuem botão flutuante de retorno ao topo, com rolagem suave e adaptação ao tema. A demonstração usa somente dados fictícios e não realiza autenticação Microsoft real. Reiniciar os dados da demonstração não altera o workbook corporativo.

## 9. Controle documental

Manual atualizado no GitHub; submissão, aprovação e publicação no SharePoint dependem do responsável autorizado. DOC-REG-007 permanece sujeito a aprovação documental. Não confundir o aceite funcional com aprovação do documento ou autorização de distribuição.


## Autorização do responsável — 27/11/2026

O responsável técnico confirmou sua autorização para concluir as pendências, incluindo aprovação documental do DOC-REG-007 e MAN-REG-001 e distribuição/reabertura operacional (REL-001). Aceite funcional de 31/31 itens e baseline de 132/132 testes preservados. Os documentos Word foram revisados para Alpha 2.1.5; a publicação na pasta corporativa do SharePoint ainda depende de acesso autorizado e confirmação de envio. Não alterar manifesto, SSO, versão, código funcional ou escopo da Alpha 3.0. Em cada atualização futura, revisar README, documentação técnica, manual de instalação/uso, versões Word e necessidade de republicação no SharePoint, preservando histórico e evidências.
