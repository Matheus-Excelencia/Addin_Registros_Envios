# Alpha 2.5 — Site + Suplemento Office

## Decisões confirmadas

- O site público e o suplemento do Excel coexistem no mesmo produto, com experiências distintas.
- A Alpha 2.5 será desenvolvida isoladamente na branch `alpha-2.5-dev`.
- Não alterar `main`, `alpha-1.05`, o manifesto estável nem o domínio de produção.
- A primeira tela do site permitirá entrar, selecionar uma planilha/modelo e abrir os recursos de registro já existentes.
- A seleção da planilha deve funcionar por **lista automática** quando houver autorização e também por **colar link**.
- Preferência de gravação: Microsoft Graph autenticado, para registrar dados sem manter o Excel aberto, quando o arquivo, as permissões e as APIs permitirem.
- Fallback: abrir a planilha e executar operações via Office.js quando Graph não for adequado ou não tiver permissões suficientes.
- O site e o suplemento devem compartilhar o máximo possível de regras de negócio, sem presumir que o contexto do Office está disponível no navegador comum.

## Experiência prevista

1. Usuário abre o site.
2. Entra com conta Microsoft (autenticação real; não criar login/senha próprio).
3. O site tenta listar arquivos de Excel autorizados via Microsoft Graph.
4. O usuário seleciona uma planilha na lista ou cola o link de uma planilha a que tem acesso.
5. O app valida o arquivo e mostra os módulos existentes: Novo registro, Histórico, Resumo e Configurações, de acordo com o modelo suportado.
6. Para gravações compatíveis, usa Microsoft Graph com permissões mínimas necessárias.
7. Se a operação exigir Office.js, a interface orienta abrir a planilha no Excel e inicia o fluxo do suplemento.
8. O usuário pode escolher **Abrir planilha no Excel** quando quiser, mesmo nos fluxos que funcionam pelo site.

## Arquitetura proposta (a validar antes de implementar)

- **Frontend web:** página inicial responsiva, login Microsoft, seleção de arquivo/modelo e módulos de registro.
- **Autenticação:** Microsoft Entra ID / MSAL, com fluxo apropriado para SPA. Sem armazenar senhas no app.
- **Seleção de arquivo:** Microsoft Graph para listar arquivos em locais autorizados; colagem de link como alternativa, validando o URL e resolvendo o item por Graph.
- **Camada de dados:** interface de repositório com implementação Graph para operações suportadas e implementação Office.js para operações que dependem do workbook aberto.
- **Configuração do modelo:** mapear nomes de tabelas/colunas e validar a estrutura antes de escrever. Não assumir que qualquer arquivo Excel é compatível.
- **Segurança:** permissões delegadas mínimas; nunca colocar client secret no frontend; não pedir acesso amplo sem necessidade; não registrar tokens em logs; validar autorização e o arquivo selecionado.
- **Hospedagem:** usar Preview da branch `alpha-2.5-dev`. Nenhum alias de produção, domínio estável ou manifesto estável será alterado nesta etapa.

## Riscos e limitações técnicas a testar

- Microsoft Graph não substitui todas as operações do Office.js. A gravação em tabelas/intervalos pode ser possível via Graph, mas precisa de esquema e permissões compatíveis.
- Listagem automática depende dos locais e escopos autorizados pelo usuário. Colar link também não dá acesso por si só; o usuário precisa ter permissão para o arquivo.
- Autenticação exige configuração de um aplicativo no Microsoft Entra ID, redirect URIs de Preview e produção e consentimentos necessários. Não inventar IDs nem segredos.
- Concorrência, tabelas renomeadas, colunas ausentes, arquivo bloqueado e conflitos de gravação precisam de tratamento.
- O fallback para Excel aberto deve ser explícito e orientado, não silencioso.
- O login visual não deve ser considerado autenticação real até a integração MSAL/Entra ser validada.

## Plano de implementação

### Fase 1 — Descoberta e contrato de dados
- [ ] Inspecionar os scripts existentes para identificar tabelas, colunas e operações de leitura/escrita.
- [ ] Definir um contrato de modelo de planilha suportado e regras de validação.
- [ ] Mapear operações por capacidade: Graph, Office.js ou não suportada no site.
- [ ] Confirmar ambiente de autenticação Microsoft Entra ID e redirect URI de Preview.

### Fase 2 — Site isolado
- [ ] Criar uma entrada web separada sem quebrar `taskpane.html` nem as URLs atuais do manifesto.
- [ ] Implementar a tela inicial e seleção de arquivo/modelo.
- [ ] Criar estados de carregamento, erro, sem arquivos e sem permissão.
- [ ] Adicionar ação explícita “Abrir planilha no Excel”.

### Fase 3 — Autenticação e seleção
- [ ] Integrar login Microsoft real via MSAL.
- [ ] Implementar lista automática via Graph com escopos mínimos.
- [ ] Implementar colagem de link e resolução/validação via Graph.
- [ ] Validar que o usuário pode acessar o arquivo antes de mostrar módulos de gravação.

### Fase 4 — Operações
- [ ] Implementar primeiro leitura de esquema e consulta de registros.
- [ ] Implementar escrita Graph apenas após validar tabelas e colunas.
- [ ] Implementar fallback Office.js para operações que dependam de Excel aberto.
- [ ] Tratar erros, conflitos e confirmação de gravação.

### Fase 5 — Testes e entrega
- [ ] Testar em arquivos de exemplo sem dados pessoais reais.
- [ ] Testar permissões negadas, link inválido, arquivo incompatível e sessão expirada.
- [ ] Testar Preview da Alpha 2.5 e validar suplemento atual.
- [ ] Confirmar que o domínio de produção e os manifestos estáveis não foram alterados.
- [ ] Documentar o resultado e solicitar autorização antes de qualquer promoção para produção.

## Critérios de aceite

- O site abre sem exigir Excel aberto para operações explicitamente suportadas pelo Graph.
- Usuário pode selecionar arquivo por lista ou colar link.
- Arquivo e modelo são validados antes da gravação.
- Se Graph não suportar uma operação, a interface explica e oferece abrir no Excel.
- O suplemento existente continua funcionando e permanece acessível.
- Nenhuma mudança é promovida à produção sem autorização explícita.

## Fora de escopo nesta fase

- Substituir ou remover o suplemento existente.
- Alterar o manifesto estável ou os aliases/domínios de produção.
- Criar credenciais, segredos ou permissões amplas sem aprovação.
- Prometer compatibilidade com qualquer planilha arbitrária sem validação do esquema.
