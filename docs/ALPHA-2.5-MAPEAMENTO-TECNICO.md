# Alpha 2.5 — Mapeamento técnico dos módulos existentes

## Por que este mapeamento é necessário
O portal web pode ser aberto sem o Excel, mas os formulários existentes não podem ser considerados independentes do suplemento. A inspeção dos scripts mostra dependências diretas de `Excel.run`, `Office.onReady`, `Office.auth.getAccessToken` e APIs do workbook. Esta etapa documenta essas dependências antes de implementar uma camada web real.

## Mapa de operações

| Módulo | Leitura atual | Escrita/ação atual | Dependência que impede uso web independente |
|---|---|---|---|
| E-mail | Aba `Config`: empresas, supervisores, remetentes/respostas, mensagens e campos de configuração | Grava registro na aba `Email` | `Office.onReady`, `Excel.run`, contexto da pasta de trabalho e autenticação Office SSO |
| SMS | Aba `Config`: cadastros, mensagens e limite de caracteres | Grava registro na aba `SMS` | `Office.onReady`, `Excel.run` e contexto da pasta de trabalho |
| Histórico | Abas `Email`, `SMS` e `Config` | Filtra e exibe registros; “Ir para registro” seleciona a linha no Excel; duplicação usa `localStorage` e navega para o formulário | `Excel.run` para ler/selecionar linhas; fluxo de duplicação amarrado às páginas existentes |
| Resumo | Abas `Email` e `SMS` | Consolida contagens e quantidades localmente | `Office.onReady` e `Excel.run` |
| Configurações | Aba `Config` | Adiciona cadastros, mensagens e limite de SMS | `Office.onReady` e `Excel.run` |
| Painel | Não lê planilha | Navega entre páginas | Pode abrir no navegador, mas os módulos de destino continuam dependentes do Office |

## Estratégia recomendada

1. Definir o contrato de planilha suportado: nomes das abas, cabeçalhos, tipos de dados, chave do registro e regras de compatibilidade.
2. Criar uma camada de dados com interface comum e duas implementações:
   - `OfficeWorkbookRepository`: preserva o comportamento atual via Office.js.
   - `GraphWorkbookRepository`: usa Microsoft Graph para os arquivos e operações compatíveis, após autenticação e permissões configuradas.
3. Manter validações e regras de negócio fora da UI, evitando duplicar regras entre suplemento e portal.
4. Começar pela leitura do esquema e do histórico em uma cópia de teste; só depois implementar escrita.
5. Tratar a seleção automática de arquivos via Graph e o link colado como dois caminhos para resolver o mesmo arquivo autorizado.
6. Se a operação não for compatível com Graph ou o arquivo não passar na validação, explicar o motivo e oferecer abrir o arquivo no Excel.

## Configuração externa necessária antes do login real

- Registro de aplicativo no Microsoft Entra ID aprovado pela organização.
- Client ID público e tenant/autoridade corretos.
- Redirect URI de Preview e depois, separadamente, URI de produção.
- Permissões delegadas Graph mínimas e consentimento conforme as políticas do tenant.
- Não usar client secret no frontend; não registrar tokens; não solicitar permissões amplas sem justificar.

## Segurança e dados
- Não adicionar credenciais, IDs inventados ou dados pessoais ao repositório.
- Não testar com arquivos de produção ou dados reais.
- Um link de arquivo não substitui autorização: o usuário deve estar autenticado e possuir acesso.
- A interface atual do portal é navegação inicial; ainda não é um sistema de login nem uma camada de gravação web.

## Critério de conclusão
O portal será considerado funcional sem Excel apenas para as operações implementadas e testadas via Graph. Para cada ação, registrar a capacidade usada (Graph ou Office.js), o resultado do teste, o erro esperado e o fallback.
