# Registra Alpha 3.0 — proposta, não implementada

## Objetivo
Evoluir de um registrador específico de E-mail/SMS para uma plataforma configurável de formulários conectados a pastas de trabalho Excel, inspirada na identificação e seleção de planilhas do projeto Campanhas.

## Escopo proposto
1. Identificador de workbook, abas, tabelas, cabeçalhos e tipos. Detectar possíveis colunas automaticamente, mostrar confiança e alternativas; confirmar mapeamento manualmente. Prévia somente leitura, nunca criar/alterar dados sem confirmação.
2. Construtor visual: criar múltiplos formulários do zero ou por modelo; adicionar, remover, reordenar e nomear campos; texto, número, moeda, porcentagem, data/hora, sim/não, seleção única/múltipla. Campos padrão atuais são apenas modelo legado, não obrigatórios em outros perfis. Data/hora automática ou manual, operador SSO ou seleção manual, ID automático, valores padrão, campo calculado somente com regras seguras e sem injetar fórmulas arbitrárias. Regras de obrigatoriedade, dependência/visibilidade condicional, validação por tipo/faixa, listas oriundas de outra aba, prévia antes de salvar.
3. Perfil compartilhado versionado por workbook: formulário, esquema, colunas, destino, regras e modelos. Preferências pessoais independentes. Histórico de revisões, rascunho/publicação, permissões administrativas a definir com Microsoft 365; não confundir UI com autorização.
4. Histórico e indicadores dinâmicos: filtros e agrupamentos com base nos tipos definidos, soma/contagem/média, campos visíveis por usuário, duplicação segura, busca e exportação autorizada.
5. Segurança e robustez: detecção de mudanças de cabeçalho, mapeamento inválido bloqueia gravação, IDs estáveis, confirmação de destino, prevenção de fórmulas, isolamento de workbook, concorrência, registro de revisão e recuperação de configurações. Compatibilidade com o workbook atual, sem migração destrutiva.
6. Modelos reutilizáveis: E-mail/SMS, atividades, produtividade, controle de qualidade, atendimento, sem obrigar campos não utilizados. Importar/exportar apenas configuração sem dados de clientes, com revisão e versão.

## Fases sugeridas
A: inventário e prévia somente leitura. B: editor visual e schema versionado. C: motor de validação e gravação, compatibilidade legada. D: histórico/resumo dinâmico e gestão de perfis. E: testes em vários workbooks fictícios, segurança, SSO e homologação. Seleção de arquivos externos ao workbook aberto via SharePoint/OneDrive será avaliada separadamente, pois exige integração e permissões; não presumir acesso administrativo. Sem mudança de manifesto ou credenciais nesta proposta.
