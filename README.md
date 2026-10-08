# Registro de Envios — Repositório

Repositório do Office Add-in para Excel Web. Esta branch `alpha-2.0-dev` mantém a aplicação-base herdada da Alpha 1.05 no diretório raiz e isola o trabalho novo da Alpha 2.0 em `alpha2-prototype/`.

> **Proteção da versão estável:** este README não altera a aplicação. A Alpha 1.05 e o deploy de produção devem continuar isolados; não mesclar a Alpha 2.0 sem autorização explícita.

## Onde encontrar cada coisa

| Caminho | Finalidade |
|---|---|
| [`alpha2-prototype/`](alpha2-prototype/) | Ambiente isolado de desenvolvimento da Alpha 2.0 |
| [`alpha2-prototype/index.html`](alpha2-prototype/index.html) | Protótipo navegável com dados fictícios |
| [`alpha2-prototype/assets/`](alpha2-prototype/assets/) | Ícone e recursos visuais exclusivos da Alpha 2.0 |
| [`alpha2-prototype/docs/`](alpha2-prototype/docs/) | Arquitetura, escopo e documentação do protótipo |
| [`alpha2-prototype/tests/`](alpha2-prototype/tests/) | Roteiros de teste; usar apenas dados fictícios |
| [`alpha2-prototype/manifest/`](alpha2-prototype/manifest/) | Local reservado ao manifesto separado de teste |
| [`manifest.xml`](manifest.xml) | Manifesto da aplicação-base presente na raiz |
| `taskpane.*`, `novo.*`, `email.*`, `sms.*`, `historico.*`, `resumo.*`, `config.*`, `commands.*` | Arquivos da aplicação-base herdada da Alpha 1.05 |
| [`assets/`](assets/) | Ícones/recursos da aplicação-base; não misturar com o ícone da Alpha 2.0 |

## Estrutura resumida

```text
Addin_Registros_Envios/
├── README.md                    # Índice geral deste repositório
├── manifest.xml                 # Manifesto da aplicação-base
├── taskpane.*, novo.*
├── email.*, sms.*
├── historico.*, resumo.*
├── config.*, commands.*
├── assets/                      # Recursos da aplicação-base
└── alpha2-prototype/
    ├── README.md                # Guia do ambiente Alpha 2.0
    ├── index.html               # Protótipo isolado
    ├── assets/                  # Identidade visual Alpha 2.0
    ├── docs/                    # Arquitetura e planejamento
    ├── tests/                   # Roteiros de teste
    └── manifest/                # Preparação do manifesto de teste
```

## Por que os arquivos da aplicação-base continuam na raiz?

Eles são mantidos no lugar para evitar quebrar o caminho de entrada e a configuração de deploy já existentes. A organização da Alpha 2.0 acontece dentro de sua própria pasta, sem mover arquivos executáveis da aplicação-base nem mudar o manifesto usado por ela.

## Alpha 2.0 — estado atual

- O protótipo é visual e usa dados fictícios.
- Ainda não grava no Excel nem em um serviço externo.
- O manifesto separado só deve ser criado depois que houver uma URL de Preview exclusiva e verificada.
- A publicação independente na Vercel está bloqueada por falta de permissão de criação/implantação (erro 403).
- Não foram desativadas proteções nem alterado o deploy de produção.

## Regras de desenvolvimento

1. Trabalhar na branch `alpha-2.0-dev`.
2. Não alterar `main` ou `alpha-1.05` para desenvolver a Alpha 2.0.
3. Não apontar o manifesto de teste para a produção da Alpha 1.05.
4. Não usar dados pessoais reais no protótipo.
5. Não criar banco, serviço pago ou credenciais sem verificar necessidade e custo.
6. Não depender de privilégios administrativos do SharePoint.
7. Não mesclar nem publicar a Alpha 2.0 em produção sem autorização explícita.

## Links úteis

- [Pasta Alpha 2.0](alpha2-prototype/)
- [Arquitetura](alpha2-prototype/docs/ARCHITECTURE.md)
- [Documentação do protótipo](alpha2-prototype/docs/PROTOTIPO.md)
- [Plano de desenvolvimento](alpha2-prototype/docs/README.md)
- [Roteiros de teste](alpha2-prototype/tests/README.md)
- [Manifesto de teste — notas](alpha2-prototype/manifest/README.md)
- [Recursos visuais — notas](alpha2-prototype/assets/README.md)
