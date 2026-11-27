# MAN-REG-001 — Manual de instalação e atualização

**Produto:** Registra · Alpha 2.0 | **Manifesto:** 2.0.0.0 | **Ambiente:** Excel Web / Microsoft 365 | **Situação:** edição preparada, aguardando homologação e autorização de publicação. Não distribuir como versão oficial antes do aceite.

## 1. Requisitos
Conta Microsoft 365 autorizada, permissão de edição no workbook e política que permita suplemento personalizado. Navegador atualizado. O suplemento não dispara E-mail ou SMS: registra envios já realizados. Não exige banco próprio.

## 2. Instalação após liberação
1. Acessar https://addin-registros-envios.vercel.app/ e clicar **Baixar manifesto XML**. Confirmar que a versão liberada é 2.0.0.0; salvar como `manifest.xml`, sem extensão .txt. Durante o desenvolvimento, esse endereço ainda entrega a versão anterior.
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
Preview protegido recusou abrir dentro do Excel Web segundo relato do usuário; não homologado. Teste controlado no domínio oficial somente após autorização, suspensão de uso e backup de dados. [Plano de publicação e rollback](PLANO-PUBLICACAO-ALPHA-20.md) e [matriz de testes](ALPHA-20-HOMOLOGACAO.md). Responsável, data e aprovação: pendentes.
