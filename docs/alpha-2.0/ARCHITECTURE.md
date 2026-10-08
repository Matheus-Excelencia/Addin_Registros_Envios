# Alpha 2.0 — Arquitetura inicial

## Objetivo
Evoluir o suplemento Excel de formulários fixos para uma plataforma configurável por modelos, preservando a Alpha 1.05 como versão estável.

## Restrições de projeto
- Não depender de acesso administrativo ao SharePoint.
- Não alterar a branch `alpha-1.05`.
- Não alterar nem migrar planilhas operacionais sem confirmação e teste em cópia.
- O suplemento só pode operar nos arquivos e recursos aos quais o usuário já tem permissão.
- Recursos que exijam consentimento de administrador do Microsoft 365 não podem ser pré-requisitos. SSO corporativo fica condicionado à disponibilidade real; identidade não verificada nunca deve ser simulada.
- Não armazenar segredos no código do navegador nem em planilhas.
- Nunca escrever em coluna apenas porque ela ocupa determinada posição: resolver por mapeamento de cabeçalho e solicitar revisão quando houver ambiguidade ou alteração.

## Arquitetura proposta
1. **Suplemento Excel (Office.js)**: interface, formulários dinâmicos, validação inicial, leitura dos cabeçalhos, gravação e navegação para células.
2. **Hospedagem/API**: Vercel é a candidata para hospedar a interface e endpoints server-side, após confirmar o projeto e as permissões existentes.
3. **Armazenamento de configuração**: banco externo com controle de acesso e políticas por usuário, preferencialmente Supabase/Postgres se um projeto autorizado e economicamente adequado estiver disponível.
4. **Planilhas Excel/SharePoint**: continuam sendo o destino dos registros operacionais quando o usuário possui acesso de edição. Configurações técnicas, versões e auditoria não devem ser espalhadas pela planilha de negócio.
5. **Sem dependência de privilégio administrativo**: iniciar desenvolvimento com branch isolada e arquivos de teste locais. Integrações organizacionais ficam atrás de interfaces substituíveis.

## Componentes de configuração
- Modelos: nome, descrição, status, destino e versão publicada.
- Módulos por modelo: Cadastro, Registro de E-mail, Registro de SMS, Histórico, Pesquisa e Resumo.
- Campos: identificador estável, rótulo, tipo, ordem, obrigatoriedade, padrão, validações, opções e vínculo de coluna.
- Regras condicionais: condição, operador, valor e campos exibidos/obrigatórios.
- Mapeamentos: identificador da planilha, nome da aba, cabeçalho, coluna atual e última validação.
- Versões: snapshots imutáveis das configurações publicadas e motivo da alteração.
- Permissões: capacidade de usar modelos versus configurar/publicar, aplicadas no servidor quando houver backend.
- Auditoria: ator verificado quando possível, data/hora, ação, modelo, versão e resumo da mudança. Falha de identificação deve ser registrada como não verificada, nunca inventada.

## Modelo relacional inicial (proposta, não implantado)
- `models`: identidade e estado do modelo.
- `model_versions`: snapshots imutáveis do modelo.
- `model_modules`: módulos habilitados por versão.
- `model_fields`: campos, validações, opções e mapeamentos por versão.
- `model_rules`: regras condicionais por versão.
- `model_permissions`: permissões por usuário/grupo, caso a identidade e o backend estejam configurados.
- `audit_events`: eventos de auditoria com acesso restrito.
- `workbook_bindings`: referências e estado de validação dos vínculos com planilhas.

As tabelas devem ter IDs estáveis, chaves estrangeiras, timestamps e políticas de acesso. As regras de autorização devem ser aplicadas no backend/banco; esconder botões na interface não é segurança.

## Gravação segura no Excel
1. Identificar arquivo e aba selecionados.
2. Detectar e exibir os cabeçalhos encontrados.
3. Resolver campos por mapeamento explícito.
4. Parar e pedir revisão se houver cabeçalhos ausentes, duplicados ou alterados.
5. Validar os dados e a permissão de edição.
6. Gravar somente nas colunas mapeadas, sem apagar dados alheios.
7. Recarregar/verificar o resultado e mostrar sucesso ou erro real.
8. Usar ID estável quando possível; para dados antigos sem ID, guardar referência à aba/linha e confirmar a identidade do registro antes de navegar.

## CPF e dados pessoais
- Validar CPF por dígitos verificadores, além de formato.
- Solicitar CPF da testemunha quando a resposta à pergunta condicional for “Sim”.
- Coletar apenas os dados necessários ao processo, restringir acesso e evitar CPF completo em logs técnicos.
- Não usar CPF como identificador primário do registro.

## Estratégia de entrega
1. Criar a branch isolada e documentar arquitetura (iniciado).
2. Auditar a estrutura atual e criar ambiente de teste sem tocar na Alpha 1.05.
3. Confirmar opção real de persistência e autenticação antes de criar banco ou gerar custos.
4. Implementar primeiro um modelo mínimo e formulário dinâmico com dados de teste.
5. Implementar detecção de cabeçalhos, mapeamentos e gravação segura em cópia.
6. Adicionar regras condicionais, listas e campos tipados.
7. Adicionar Histórico, Pesquisa e navegação.
8. Adicionar versões, permissões, auditoria e Resumo.
9. Testar concorrência, alterações de cabeçalhos, permissões, compatibilidade e recuperação.

## Estado
Este documento registra a proposta inicial. Nenhum banco Supabase foi criado, nenhuma integração foi autorizada e nenhuma planilha foi alterada.