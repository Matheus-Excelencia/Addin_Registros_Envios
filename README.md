# Registro de Envios — Índice do repositório

Este repositório contém a aplicação Office Add-in de Registro de Envios e o ambiente isolado de desenvolvimento da Alpha 2.0.

## 1. Aplicação-base (raiz)

As páginas HTML e o manifesto permanecem na raiz para preservar os endpoints atuais. Os arquivos CSS e JavaScript foram movidos para `app/styles/` e `app/scripts/`, e os HTMLs foram atualizados para carregar os novos caminhos.

| Grupo | Arquivos | Responsabilidade |
|---|---|---|
| Entrada | `manifest.xml`, `taskpane.html` | Manifesto e painel principal |
| Novo registro | `novo.html` | Escolha do módulo E-mail ou SMS |
| E-mail | `email.html` | Formulário e gravação de registros de E-mail |
| SMS | `sms.html` | Formulário e gravação de registros SMS |
| Histórico | `historico.html` | Consulta, filtros e ações sobre registros |
| Resumo | `resumo.html` | Indicadores consolidados |
| Configurações | `config.html` | Listas, mensagens e parâmetros |
| Comandos Office | `commands.html` | Página de comandos referenciada pelo manifesto |
| Estilos | `app/styles/` | CSS de todos os módulos |
| Scripts | `app/scripts/` | JavaScript de todos os módulos |
| Recursos | `assets/` | Ícones da aplicação-base |

> Não mova as páginas HTML nem altere os endpoints do manifesto sem atualizar e testar as rotas e a configuração de hospedagem.

## 2. Alpha 2.0 — desenvolvimento isolado

A Alpha 2.0 fica em [`alpha2-prototype/`](alpha2-prototype/).

| Caminho | Conteúdo |
|---|---|
| [`alpha2-prototype/index.html`](alpha2-prototype/index.html) | Protótipo visual com dados fictícios |
| [`alpha2-prototype/assets/`](alpha2-prototype/assets/) | Ícone e recursos visuais próprios |
| [`alpha2-prototype/docs/`](alpha2-prototype/docs/) | Arquitetura, escopo e plano de trabalho |
| [`alpha2-prototype/tests/`](alpha2-prototype/tests/) | Roteiros de validação |
| [`alpha2-prototype/manifest/`](alpha2-prototype/manifest/) | Notas para o manifesto independente de teste |

### Documentação da Alpha 2.0

- [Guia da pasta Alpha 2.0](alpha2-prototype/README.md)
- [Arquitetura](alpha2-prototype/docs/ARCHITECTURE.md)
- [Descrição do protótipo](alpha2-prototype/docs/PROTOTIPO.md)
- [Plano de desenvolvimento](alpha2-prototype/docs/README.md)
- [Roteiros de teste](alpha2-prototype/tests/README.md)
- [Notas dos recursos visuais](alpha2-prototype/assets/README.md)
- [Notas do manifesto de teste](alpha2-prototype/manifest/README.md)

## 3. Regras de isolamento e segurança

1. A Alpha 1.05 e a produção devem permanecer estáveis.
2. Desenvolver a Alpha 2.0 na branch `alpha-2.0-dev`.
3. Não apontar o manifesto da Alpha 2.0 para a URL de produção.
4. Não usar dados pessoais reais nos testes do protótipo.
5. Não criar serviços pagos, banco de dados ou credenciais sem avaliar necessidade e custo.
6. Não depender de privilégios administrativos do SharePoint.
7. Não mesclar a Alpha 2.0 na linha estável sem autorização explícita.

## 4. Estado conhecido

- A aplicação-base e o protótipo Alpha 2.0 coexistem nesta branch, mas têm funções distintas.
- O protótipo ainda usa dados fictícios e não grava no Excel nem em serviços externos.
- O manifesto separado da Alpha 2.0 depende de uma URL de Preview própria confirmada.
- A implantação independente na Vercel foi recusada com erro 403; nenhuma proteção da produção foi alterada.

## 5. Repositório e branch

- Repositório: [Matheus-Excelencia/Addin_Registros_Envios](https://github.com/Matheus-Excelencia/Addin_Registros_Envios)
- Branch de organização/desenvolvimento: `alpha-2.0-dev`
- Aplicação de produção (referência, não alterar neste trabalho): [addin-registros-envios.vercel.app](https://addin-registros-envios.vercel.app/)
