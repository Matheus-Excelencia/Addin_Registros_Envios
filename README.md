<div align="center">
<img src="public/assets/brand-registros.svg" width="96" alt="Registra">

# Registra

**Suplemento corporativo para Microsoft Excel Web**

Anteriormente denominado Registros de Envios; o workbook mantém seu nome original.

*Menos retrabalho. Mais controle.*

![Versão](https://img.shields.io/badge/versão-Alpha%202.1.5-087f72)
[![Validação main](https://github.com/Matheus-Excelencia/Addin_Registros_Envios/actions/workflows/alpha20.yml/badge.svg?branch=main)](https://github.com/Matheus-Excelencia/Addin_Registros_Envios/actions/workflows/alpha20.yml)
![Excel](https://img.shields.io/badge/Excel-Web-217346)
[![License: uso corporativo](https://img.shields.io/badge/license-uso%20corporativo%20restrito-555)](LICENSE)
![Office.js](https://img.shields.io/badge/Office.js-API-blue)
[![Vercel: Alpha 2.1.5](.github/assets/vercel-status.svg)](https://addin-registros-envios.vercel.app/)

[Landing page](https://addin-registros-envios.vercel.app/) · [Demonstração](https://addin-registros-envios.vercel.app/demo/index.html) · [Instalação](#instalação) · [Funcionalidades](#funcionalidades) · [Arquitetura](#arquitetura) · [Documentação](#documentação) · [Licença](#licença)

</div>

> **Alpha 2.1.5 — homologação funcional aprovada pelo responsável em 27/11/2026.** 31/31 validações manuais declaradas e 132/132 testes automatizados aprovados. Aprovação documental e distribuição autorizadas pelo responsável; envio ao SharePoint pendente. Backup Alpha 1.06 preservado.

## Visão geral

O Registra padroniza o registro operacional de campanhas de E-mail e SMS/URA realizadas no Konfiansa. **O suplemento não dispara mensagens.** Ele registra os envios já realizados no workbook corporativo `Registros de Envios.xlsx`, armazenado no SharePoint.

## Funcionalidades

| Módulo | Responsabilidade |
|---|---|
| Novo registro | Escolha entre E-mail e SMS |
| E-mail | Registro, validação de endereço, assunto e mensagem |
| SMS | Registro, contador e limite configurável |
| Histórico | Filtros, ordenação, duplicação, perguntas personalizadas e navegação até a linha |
| Resumo | Consolidação por período, cards, lista e gráficos |
| Configurações | Listas, mensagens, perguntas dinâmicas e formatos Excel |
| Identidade | Microsoft SSO, com contingência manual nos formulários |

### Alpha 2.0 — homologação controlada

- Identidade visual compartilhada, temas persistentes e abertura de sete segundos com botão Continuar.
- Resumo em cards, lista e gráficos; regras de visibilidade de perguntas no Histórico e Resumo.
- Duplicação pelo fluxo existente do Histórico, incluindo respostas personalizadas.
- Demonstração sem SSO, com reinício dos dados fictícios e da abertura; autenticação real preservada.
- Marco histórico: homologação funcional e SSO posteriormente aprovados na Alpha 2.1.5.

### Novidades da Alpha 1.06

- Campo **Valor** opcional por padrão, numérico e contábil em reais.
- Perguntas configuráveis por formulário, tipo, obrigatoriedade e formato.
- Edição, desativação e reativação de perguntas; coluna, tipo e destino protegidos na edição.
- Avisos de duplicidade de coluna e contador de SMS na configuração.
- Saudação de boas-vindas uma vez por sessão. Digitação e esvanecimento posteriormente aprovados na homologação Alpha 2.1.5 (UX-005/UX-006).

## Instalação

**Ambiente oficial:** Excel para a Web, com conta Microsoft 365 e permissão de edição no workbook.

1. [Baixe o manifesto XML](https://addin-registros-envios.vercel.app/manifest.xml) (salve como `manifest.xml`, não como TXT).
2. Durante a homologação, abra somente o workbook fictício no Excel Web. Após aceite, use o workbook corporativo autorizado.
3. Acesse **Início → Suplementos → Avançado** ou **Inserir → Suplementos → Mais Suplementos**, conforme a interface disponível.
4. Escolha **Carregar meu suplemento**, selecione o XML e confirme o carregamento.
5. Na guia Início, procure o grupo **Excelência** e o botão **Novo registro**.
6. Confirme a abertura do painel e a identificação do usuário nos formulários.

> Os nomes dos menus podem variar. Se o carregamento personalizado estiver indisponível, consulte o administrador Microsoft 365. Não contorne políticas corporativas. Excel Desktop não é o ambiente oficial.

## Arquitetura

```mermaid
flowchart LR
    K[Konfiansa] -->|Envio realizado| U[Operador]
    U --> X[Excel Web]
    X --> M[Manifesto XML]
    M --> P[Task Pane / Vercel]
    P --> E[E-mail e SMS]
    P --> C[Configurações]
    P --> H[Histórico e Resumo]
    E --> A[Office.js / Excel API]
    C --> A
    H --> A
    A --> W[Workbook / SharePoint]
```

HTML, CSS e JavaScript compõem a interface; Office.js integra o workbook. O código é versionado no GitHub e hospedado na Vercel. Não há banco de dados independente nessa arquitetura.

## Modelo de dados

| Coluna | Email | SMS |
|---|---|---|
| A | Data | Data |
| B | Realizado por | Realizado por |
| C | Empresa | Empresa |
| D | Qtde | Qtde |
| E | Email Resposta | Supervisor |
| F | Supervisor | Obs |
| G | Obs | Historico Externo |
| H | Historico Externo | Mensagem |
| I | Mensagem | ID REGISTRO |
| J | Assunto | Valor |
| K | ID REGISTRO | — |
| L | Valor | — |

A aba `Config` contém empresas, e-mails de resposta, supervisores, realizadores, nomes e textos de mensagens de E-mail/SMS e limite SMS (A–I). Metadados adicionais incluem PERGUNTA PERSONALIZADA, FORMULARIO, TIPO RESPOSTA, OBRIGATORIA, FORMATO EXCEL, COLUNA EXCEL e PERGUNTA ATIVA, EXIBIR HISTORICO e EXIBIR RESUMO. Perguntas adicionais usam cabeçalhos próprios. Preserve a estrutura oficial; não presuma migração automática de qualquer planilha antiga.

## Regras e segurança

- Quantidade inteira positiva; campos obrigatórios preenchidos; e-mail com formato válido.
- SMS: limite padrão de 160 caracteres, configurável; bloqueio sem truncamento. O contador não estima segmentos da operadora.
- Data automática, ID de registro e confirmação de gravação; duplicação gera novo registro.
- A proteção de concorrência não equivale a transações de banco de dados.
- SSO via Office.auth.getAccessToken; fallback manual nos formulários. Não armazene tokens ou credenciais.
- Acesso ao workbook sujeito às permissões do Microsoft 365 e SharePoint.

## Qualidade e histórico

| Versão / marco | Evidência |
|---|---|
| Alpha 1.01–1.05 | Consolidação dos módulos e operação Excel Web (outubro/2026) |
| [PR #11](../../pull/11) | 22 testes; homologação 33/33 |
| [PR #12](../../pull/12) | 26 testes; homologação 33/33 |
| [PR #13](../../pull/13) | 31 testes; homologação 28/30 |
| [PR #14](../../pull/14) | [33 testes aprovados](../../actions/runs/45456326476); 20 itens manuais não validados |

As suítes são cumulativas. [132/132 testes aprovados](https://github.com/Matheus-Excelencia/Addin_Registros_Envios/actions/runs/46027246753) no baseline funcional `ba1b10c`; [31/31 validações manuais aprovadas pelo responsável](docs/ALPHA-215-HOMOLOGACAO-FINAL.md). A validação automatizada não substitui o teste no host Office.

## Organização do repositório e desenvolvimento

A raiz contém apenas configuração, documentação e ferramentas de desenvolvimento. A Vercel publica exclusivamente `public/`; os endereços existentes do Excel não mudam.

```text
.
├── .github/workflows/    # Validação automatizada
├── docs/                 # Segurança, homologação e manuais
├── public/               # Arquivos web publicados
│   ├── assets/           # Ícones e identidade visual
│   ├── demo/             # Simulação com dados fictícios
│   ├── manifest.xml      # Contrato Office
│   ├── index.html        # Landing page
│   └── ...               # Páginas e recursos do suplemento
├── scripts/              # Sincronização da demonstração
├── tests/                # Testes automatizados
├── CONTRIBUTING.md
├── LICENSE
├── README.md
└── vercel.json
```

| Módulo (em public/) | Finalidade |
|---|---|
| taskpane.* / novo.* | Menu e abertura |
| email.* / sms.* / perguntas.js | Formulários e gravação |
| config.* | Configurações |
| historico.* / resumo.* | Consulta e indicadores |
| registro-security.js | Proteção de gravações e duplicação |
| brand.* / question-visibility.js | Tema e regras de exibição |
| commands.* / manifest.xml | Integração Office |

Com Node.js 22, execute na raiz:

```bash
node scripts/sync-demo.cjs --check
node --test tests/*.test.cjs
```

Consulte [CONTRIBUTING.md](CONTRIBUTING.md) para o fluxo de desenvolvimento e publicação. O manifesto para usuários é sempre o da produção, não o arquivo bruto da branch.

## Documentação

- [Manual Alpha 2.1.5 (autorizado; envio ao SharePoint pendente)](docs/MAN-REG-001-ALPHA-20.md)
- [Revisão de segurança](docs/ALPHA-20-SEGURANCA.md)
- [Homologação histórica Alpha 2.0](docs/ALPHA-20-HOMOLOGACAO.md)
- [Homologação final Alpha 2.1.5](docs/ALPHA-215-HOMOLOGACAO-FINAL.md)
- [Documentação técnica Alpha 2.1.5](docs/ALPHA-215.md)
- [Política de atualização](docs/POLITICA-ATUALIZACAO.md)

- [Pasta oficial de documentação — SharePoint](https://excelenciacobrancaempres629.sharepoint.com/:f:/s/backofficearquivos/IgB1hL4JmiUnRJ6PWKmwB5aQAVWwebeL5AhbYGjeiiRju2M?e=48kyBS) (acesso corporativo).
- DOC-REG-007 — revisão Alpha 2.1.5 autorizada pelo responsável; envio ao SharePoint pendente.
- [Aplicação hospedada](https://addin-registros-envios.vercel.app/).
- [Histórico de commits](../../commits/main).

---

<div align="center">

**Grupo Excelência · Registra**  
Responsabilidade técnica: Matheus Passos  
Uso corporativo · Alpha 2.1.5 — aceite funcional registrado

</div>

## Homologação isolada
[Instruções e manifesto exclusivo de teste](homologacao/README.md). Use apenas dados fictícios. O preview é protegido; SSO de produção não é reutilizado.

## Preparação para lançamento
[Plano de publicação e rollback](docs/PLANO-PUBLICACAO-ALPHA-20.md) · [Backup Alpha 1.06](https://github.com/Matheus-Excelencia/Addin_Registros_Envios/tree/backup/alpha-1.06-pre-registra). Preview Office bloqueado; publicação controlada autorizada. Aceite funcional e autorização de distribuição/reabertura registrados.

## Licença

[Uso corporativo restrito — LICENSE](LICENSE). Não é um projeto de código aberto; não há concessão de licença pública de reutilização. Direitos de terceiros são preservados. A titularidade e eventuais permissões adicionais dependem dos respectivos titulares.

## Alpha 2.0.6
- Avisos temporários e validação próxima ao campo.
- Configurações → Ordem dos formulários: campos padrão e personalizados, por canal, originalmente salvos na aba Config (`ORDEM FORMULARIOS`); substituídos pelas preferências pessoais na Alpha 2.1.5. O manifesto 2.0.0.0 permanece estável.
- [Proposta Alpha 3.0](docs/ROADMAP-ALPHA-30.md).

## Alpha 2.1.5
- Ordem dos formulários e exibição no Histórico/Resumo são pessoais, separadas por identidade Microsoft e workbook no navegador. Sem SSO, preferências temporárias de sessão; sem identidade de arquivo, somente em memória.
- Criação e edição das perguntas continuam compartilhadas. Colunas legadas de visualização são preservadas, mas não utilizadas para definir a preferência pessoal.
- Log de conexão com Config permanece visível e seção de ordenação usa a identidade visual Registra.
- Manifesto 2.0.0.0 permanece inalterado. [Roadmap Alpha 3.0](docs/ROADMAP-ALPHA-30.md).

### Organização da Config
- Cadastros realizados em Configurações ou nos formulários E-mail/SMS são ordenados por lista, em português brasileiro, a cada inclusão de valor novo. Campos existentes não são duplicados.
- Mensagens são ordenadas pelo nome, mantendo o texto correspondente; `wrapText` é desativado nos intervalos atualizados, sem remover quebras de linha inseridas no texto. Edições manuais diretamente na planilha não disparam a ordenação.
- A página inicial e o artigo possuem botão de retorno ao topo; layout e temas restaurados e aprovados.


## Autorização do responsável — 27/11/2026

O responsável técnico confirmou sua autorização para concluir as pendências, incluindo aprovação documental do DOC-REG-007 e MAN-REG-001 e distribuição/reabertura operacional (REL-001). Aceite funcional de 31/31 itens e baseline de 132/132 testes preservados. Os documentos Word foram revisados para Alpha 2.1.5; a publicação na pasta corporativa do SharePoint ainda depende de acesso autorizado e confirmação de envio. Não alterar manifesto, SSO, versão, código funcional ou escopo da Alpha 3.0. Em cada atualização futura, revisar README, documentação técnica, manual de instalação/uso, versões Word e necessidade de republicação no SharePoint, preservando histórico e evidências.
