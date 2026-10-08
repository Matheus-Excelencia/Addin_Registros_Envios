# Registro de Envios — Office Add-in para Excel Web

Office Add-in desenvolvido para **Excel Web / Microsoft 365** com o objetivo de registrar, padronizar e centralizar os envios de **E-mail** e **SMS/URA** realizados no sistema **Konfiansa**.

O projeto utiliza uma arquitetura baseada em **Office Add-ins**, com **Task Pane**, JavaScript e integração direta com a pasta de trabalho do Excel por meio da **Excel JavaScript API**.

> **Documentação técnica do projeto:** [Documentação Projeto Registro de Envios](https://excelenciacobrancaempres629.sharepoint.com/:w:/s/backofficearquivos/IQB79cWT2VE7RYT1MOfd8GDtAf6Lf5umxpsdou1uH6neI0s?e=5BDbt1)

---

## 1. Visão geral

O Add-in funciona como uma camada de registro operacional após a execução dos disparos no Konfiansa.

O disparo das mensagens **não é realizado pelo Add-in**. O processo de envio permanece no Konfiansa, enquanto o Add-in é responsável pelo registro estruturado das informações no arquivo `Registros de Envios.xlsx`.

### Fluxo de negócio

```
Konfiansa
   │
   ▼
CAMPANHA
   ├── E-mail em Lote
   └── SMS/URA em Lote
          │
          ▼
      Envio realizado
          │
          ▼
   Registros de Envios
          │
     ┌────┴────┐
     ▼         ▼
   Email      SMS
```

---

## 2. Objetivos técnicos

O projeto foi estruturado para:

- Centralizar os registros de disparos realizados;
- Padronizar os dados armazenados no Excel;
- Separar a operação de E-mail da operação de SMS;
- Utilizar uma configuração centralizada na aba `Config`;
- Reduzir digitação repetitiva por meio de listas configuráveis;
- Registrar automaticamente data e hora;
- Identificar o usuário do Microsoft 365 quando a plataforma disponibilizar essa informação;
- Validar os dados antes da gravação;
- Preservar registros existentes durante novas inserções;
- Permitir manutenção e evolução independente dos módulos.

---

# 3. Plataforma e arquitetura

### Plataforma oficial

O projeto é destinado ao:

**Excel para a Web / Microsoft 365**

A aplicação utiliza um **Office Add-in** carregado como painel lateral dentro do Excel.

### Componentes

```
Excel Web
   │
   ├── Office Add-in
   │      │
   │      ├── Manifest
   │      ├── Task Pane
   │      ├── Módulo E-mail
   │      └── Módulo SMS
   │
   └── Excel JavaScript API
          │
          ▼
   Registros de Envios.xlsx
          │
          ├── Email
          ├── SMS
          └── Config
```

### Hospedagem

A aplicação web do Add-in é hospedada na **Vercel**:

```
https://addin-registros-envios.vercel.app/
```

O `manifest.xml` referencia a aplicação hospedada e define as configurações utilizadas pelo Office para carregar o Add-in.

---

# 4. Instalação no Excel Web

## Pré-requisitos

- Microsoft 365;
- Excel para a Web;
- Acesso ao arquivo `Registros de Envios.xlsx`;
- Arquivo `manifest.xml` do projeto;
- Permissão para carregar Add-ins personalizados na organização.

## Instalação

1. Abrir o **Excel Web**.
2. Abrir o arquivo `Registros de Envios.xlsx`.
3. Acessar **Início → Suplementos → Avançado**.
4. Selecionar **Carregar meu suplemento**.
5. Escolher **Carregar de um arquivo**.
6. Selecionar o arquivo `manifest.xml`.
7. Abrir o Add-in **Registros de Envios**.
8. O painel lateral deverá apresentar as opções:
   - **E-mail**
   - **SMS**

Dependendo da configuração do Microsoft 365, os menus podem aparecer em **Inserir → Suplementos → Mais Suplementos**.

> Se a opção de carregamento de Add-ins personalizados não estiver disponível, a organização pode possuir uma política administrativa restringindo esse recurso.

---

# 5. Estrutura do repositório

```
Addin_Registros_Envios/
│
├── manifest.xml
│
├── taskpane.html
├── taskpane.css
├── taskpane.js
│
├── commands.html
├── commands.js
│
├── email.html
├── email.css
├── email.js
│
├── sms.html
├── sms.css
├── sms.js
│
├── assets/
│   ├── icon-16.png
│   ├── icon-32.png
│   └── icon-80.png
│
└── README.md
```

---

# 6. Responsabilidade dos arquivos

## `manifest.xml`

Arquivo principal de configuração do Office Add-in.

Define, entre outros itens:

- Identidade do Add-in;
- Nome e descrição;
- Ícones;
- URLs da aplicação;
- Permissões;
- Requisitos de API;
- Configuração de autenticação;
- Comandos e elementos da interface do Excel;
- Localização do Task Pane.

O manifesto utiliza a ação `ShowTaskpane` para abrir a interface principal do Add-in.

## `taskpane.html`

Tela inicial do Add-in.

Responsável pela seleção do tipo de registro:

- E-mail;
- SMS.

## `taskpane.js`

Responsável pela inicialização do Office e pela navegação entre os módulos.

## `email.html / email.css / email.js`

Implementam o fluxo completo de registro de E-mail.

Responsabilidades principais:

- Carregar configurações;
- Carregar mensagens;
- Identificar usuário;
- Validar campos;
- Registrar data e hora;
- Inserir dados na aba `Email`;
- Atualizar a aba `Config` quando necessário.

## `sms.html / sms.css / sms.js`

Implementam o fluxo completo de registro de SMS.

Responsabilidades principais:

- Carregar configurações;
- Carregar mensagens específicas de SMS;
- Controlar limite de caracteres;
- Validar campos;
- Registrar data e hora;
- Inserir dados na aba `SMS`;
- Atualizar a aba `Config` quando necessário.

## `commands.html / commands.js`

Arquivos relacionados à configuração de comandos do Office Add-in.

O manifesto mantém a referência ao `FunctionFile`, embora o botão principal atualmente utilize `ShowTaskpane`.

## `assets/`

Contém os ícones utilizados pelo manifesto e pela interface do Add-in.

---

# 7. Modelo de dados do workbook

O arquivo utilizado pela aplicação é:

```
Registros de Envios.xlsx
```

As principais planilhas são:

- `Email`
- `SMS`
- `Config`

---

## 7.1 Aba Email

Estrutura atual:

| Coluna | Campo |
|---|---|
| A | Data |
| B | Realizado por |
| C | Empresa |
| D | Qtde |
| E | Email Resposta |
| F | Supervisor |
| G | Obs |
| H | Historico Externo |
| I | Mensagem |
| J | Assunto |

### Regras

**Data**

Gerada automaticamente no momento do salvamento.

**Realizado por**

O Add-in tenta identificar automaticamente o usuário conectado ao Microsoft 365.

**Empresa**

Obrigatória. Pode ser selecionada a partir da configuração existente ou informada manualmente.

**Qtde**

Obrigatória e deve representar uma quantidade inteira positiva.

**Email Resposta**

Obrigatório.

**Supervisor**

Obrigatório.

**Obs**

Campo opcional.

**Historico Externo**

Obrigatório.

**Mensagem**

Obrigatória. O usuário seleciona o nome da mensagem e o Add-in grava o texto completo configurado.

**Assunto**

Obrigatório.

---

# 8. Módulo SMS

## 8.1 Aba SMS

Estrutura atual:

| Coluna | Campo |
|---|---|
| A | Data |
| B | Realizado por |
| C | Empresa |
| D | Qtde |
| E | Supervisor |
| F | Obs |
| G | Historico Externo |
| H | Mensagem |

### Regras

- Data gerada automaticamente;
- Identificação automática do usuário quando disponível;
- Empresa obrigatória;
- Quantidade obrigatória;
- Supervisor obrigatório;
- Observação opcional;
- Histórico Externo obrigatório;
- Mensagem obrigatória.

---

## 8.2 Limite de caracteres

O módulo SMS utiliza um limite padrão de:

```
160 caracteres
```

O valor pode ser configurado pela coluna `LIMITE SMS` da aba `Config`.

O formulário:

- Exibe contador de caracteres;
- Bloqueia o salvamento quando o limite é excedido;
- Não corta automaticamente o conteúdo;
- Permite que o usuário corrija o texto ou selecione outra mensagem.

Mensagens contendo variáveis, como `[NomCliente]`, devem ser avaliadas considerando o tamanho final após a substituição.

---

# 9. Aba Config

A aba `Config` funciona como fonte central de dados auxiliares e mensagens.

Estrutura atual:

| Coluna | Campo |
|---|---|
| A | EMPRESAS |
| B | EMAIL RESPOSTA |
| C | SUPERVISORES |
| D | REALIZADO POR |
| E | NOME MENSAGEM EMAIL |
| F | TEXTO MENSAGEM EMAIL |
| G | NOME MENSAGEM SMS |
| H | TEXTO MENSAGEM SMS |
| I | LIMITE SMS |

## Atualização automática

Durante o uso dos formulários, novos valores podem ser adicionados à configuração quando ainda não estiverem cadastrados.

A lógica contempla principalmente:

- Empresas;
- Supervisores;
- E-mails de resposta;
- Usuários de realização.

A comparação utiliza normalização para evitar duplicações causadas por diferenças de maiúsculas/minúsculas ou espaços.

---

# 10. Arquitetura de mensagens

As mensagens são separadas por módulo.

## E-mail

Utiliza:

```
NOME MENSAGEM EMAIL
TEXTO MENSAGEM EMAIL
```

O usuário seleciona o nome da mensagem e o sistema recupera o texto correspondente.

## SMS

Utiliza:

```
NOME MENSAGEM SMS
TEXTO MENSAGEM SMS
```

A seleção de SMS não utiliza as mensagens configuradas para E-mail.

Essa separação evita que alterações em uma categoria afetem a outra.

---

# 11. Integração com Excel

A comunicação com o workbook utiliza:

```
Office.js
Excel JavaScript API
Excel.run()
```

O código acessa as planilhas por nome e identifica dinamicamente a estrutura dos cabeçalhos quando necessário.

A inserção dos registros ocorre na próxima linha disponível da tabela de dados utilizada, preservando os registros existentes.

Quando uma planilha ainda não possui cabeçalho reconhecível, o módulo pode inicializar a estrutura esperada antes da inserção.

---

# 12. Identificação do usuário

O módulo utiliza a API de autenticação do Office quando disponível:

```
Office.auth.getAccessToken()
```

O token é utilizado para tentar obter informações do usuário conectado, priorizando campos como:

- `name`;
- `preferred_username`;
- `email`;
- `upn`.

Caso a identificação automática não esteja disponível no ambiente, o formulário mantém comportamento de contingência para permitir o preenchimento manual quando aplicável.

---

# 13. Validação e integridade dos dados

As validações ocorrem antes da gravação no Excel.

Entre as principais regras:

- Campos obrigatórios não podem permanecer vazios;
- Quantidades devem ser inteiras e positivas;
- Mensagens devem existir na configuração;
- SMS não pode ultrapassar o limite configurado;
- Registros são adicionados na próxima linha disponível;
- A estrutura das colunas deve permanecer compatível com o código.

A validação no front-end reduz registros inconsistentes antes da chamada à Excel JavaScript API.

---

# 14. Fluxo operacional completo

```
1. Realizar o disparo no Konfiansa
                 ↓
2. Acessar CAMPANHA
                 ↓
3. Selecionar E-mail em Lote
   ou SMS/URA em Lote
                 ↓
4. Concluir o envio
                 ↓
5. Abrir Registros de Envios no Excel Web
                 ↓
6. Selecionar E-mail ou SMS
                 ↓
7. Preencher os dados
                 ↓
8. Executar a validação
                 ↓
9. Salvar
                 ↓
10. Registrar na aba correspondente
                 ↓
11. Atualizar Config quando necessário
```

---

# 15. Segurança

O código-fonte não deve conter:

- Senhas;
- Tokens permanentes;
- Chaves privadas;
- Credenciais;
- Segredos de APIs;
- Dados pessoais desnecessários.

Informações de autenticação devem ser tratadas pelos mecanismos oficiais da plataforma e mantidas fora do código-fonte sempre que aplicável.

O repositório também não deve receber credenciais reais em commits, arquivos de configuração ou documentação.

---

# 16. Desenvolvimento e manutenção

Antes de alterar o projeto, deve-se identificar o módulo afetado e preservar a separação entre:

```
Task Pane
   │
   ├── E-mail
   │
   └── SMS
```

Alterações estruturais no workbook devem ser tratadas com atenção, principalmente nas abas:

```
Email
SMS
Config
```

Mudanças de nomes de colunas, ordem dos campos ou nomes das planilhas podem exigir alterações no JavaScript.

### Recomendações de manutenção

- Alterar somente o módulo necessário;
- Evitar lógica duplicada;
- Manter mensagens de E-mail e SMS separadas;
- Preservar as validações existentes;
- Testar o salvamento após alterações;
- Verificar se registros antigos permanecem intactos;
- Validar a integração com a aba `Config`;
- Atualizar a documentação quando houver alteração funcional ou arquitetural.

---

# 17. Versionamento

O projeto é mantido em Git/GitHub.

Repositório:

```
Matheus-Excelencia/Addin_Registros_Envios
```

Os commits devem utilizar mensagens objetivas e em inglês, seguindo uma convenção semelhante a:

```
feat: add new feature
fix: correct validation issue
docs: update project documentation
refactor: reorganize module
chore: maintenance change
```

Exemplos:

```
feat: add SMS character limit configuration
fix: correct SMS form save handler
docs: update project documentation
```

---

# 18. Deploy

O projeto utiliza a Vercel para disponibilizar a aplicação web.

Fluxo simplificado:

```
Alteração no código
       ↓
Git / GitHub
       ↓
Deploy
       ↓
Vercel
       ↓
Aplicação Web
       ↓
Excel Web
```

O `manifest.xml` deve permanecer apontando para os endereços válidos da aplicação publicada.

Após alterações relevantes, deve-se validar o carregamento do Task Pane e os módulos E-mail e SMS no Excel Web.

---

# 19. Compatibilidade

A plataforma considerada oficialmente para este projeto é:

**Excel Web / Microsoft 365**

A solução foi estruturada considerando as limitações e o modelo de execução de Office Add-ins no Excel para a Web.

O projeto não depende de VBA ou UserForms tradicionais do Excel Desktop.

---

# 20. Estado atual da implementação

A implementação atual contempla:

- Office Add-in para Excel Web;
- Task Pane;
- Seleção entre E-mail e SMS;
- Registro de E-mail;
- Registro de SMS;
- Integração com `Email`, `SMS` e `Config`;
- Registro automático de data e hora;
- Tentativa de identificação do usuário Microsoft 365;
- Cadastro dinâmico de empresas;
- Cadastro dinâmico de supervisores;
- Cadastro dinâmico de e-mails de resposta;
- Mensagens independentes para E-mail e SMS;
- Limite configurável de caracteres para SMS;
- Contador de caracteres;
- Validação antes do salvamento;
- Registro de Histórico Externo;
- Persistência do texto completo das mensagens;
- Hospedagem da aplicação na Vercel.

---

# 21. Referência técnica

| Item | Valor |
|---|---|
| Projeto | Registro de Envios |
| Plataforma | Excel Web / Microsoft 365 |
| Tecnologia | Office Add-in |
| Interface | Task Pane |
| API | Office.js / Excel JavaScript API |
| Sistema de origem | Konfiansa |
| Processos | E-mail em Lote / SMS/URA em Lote |
| Módulos | E-mail / SMS |
| Workbook | Registros de Envios.xlsx |
| Hospedagem | Vercel |
| Repositório | Matheus-Excelencia/Addin_Registros_Envios |

---

## Documentação completa

A documentação técnica detalhada do projeto está disponível no SharePoint:

**[Documentação Projeto Registro de Envios](https://excelenciacobrancaempres629.sharepoint.com/:w:/s/backofficearquivos/IQCWHNmJGBi_R5tA79npAMo9AdartnTpMtqaM0UFbAeXiAk?e=tDYYkV)**

A documentação contém detalhes de instalação, arquitetura, estrutura dos arquivos, modelo de dados, regras dos módulos, autenticação, manutenção e operação do projeto.


---

# Atualização de implementação — 05/10/2026

## Versão de referência

- Linha de desenvolvimento: `main`
- Versão estável anterior: `alpha-1.04`
- Próxima versão estável de referência: `alpha-1.05`

## Módulos incorporados à documentação

A implementação atual passa a considerar oficialmente os seguintes módulos:

- **Novo Registro** — tela separada para escolha entre E-mail e SMS;
- **E-mail** — registro completo de disparos;
- **SMS** — registro completo com limite configurável;
- **Histórico** — consulta, filtros, ordenação, duplicação e navegação para a linha do registro;
- **Resumo** — consolidação dos registros por período;
- **Configurações** — fonte central das listas e parâmetros utilizados pelos formulários.

## Histórico

O Histórico consulta as abas `Email` e `SMS` e permite filtrar por:

- Tipo;
- Data inicial;
- Data final;
- Empresa;
- Realizado por;
- Supervisor;
- Pesquisa.

Também permite ordenar por:

- Mais recente;
- Mais antiga;
- Empresa A-Z;
- Realizado por A-Z;
- Supervisor A-Z.

Cada registro possui as ações **Duplicar** e **Ir para registro**.

A navegação para a planilha utiliza a linha real do registro, portanto também funciona para registros antigos que ainda não possuam `ID REGISTRO`.

A mensagem exibida no Histórico utiliza um campo redimensionável verticalmente para facilitar a leitura de textos maiores.

## Resumo

O módulo Resumo apresenta a consolidação dos registros de E-mail e SMS conforme o período selecionado.

O carregamento do módulo possui indicação de status para diferenciar o carregamento da atualização concluída.

## Configurações

A aba `Config` permanece como fonte central para:

- Empresas;
- E-mails de resposta;
- Supervisores;
- Realizado por;
- Mensagens de E-mail;
- Mensagens de SMS;
- Limite SMS.

Os formulários utilizam essas listas como fonte de seleção e podem atualizar valores configuráveis quando necessário.

## Identificação do usuário

Os formulários tentam identificar automaticamente o usuário Microsoft 365.

A interface informa visualmente:

- identificação automática realizada; ou
- necessidade de preenchimento manual quando a plataforma não disponibilizar a identificação.

## Confirmação de gravação

Após o salvamento, E-mail e SMS apresentam confirmação visual de sucesso e o `ID REGISTRO` gerado.

## ID REGISTRO

Os registros novos utilizam um identificador no padrão:

`ENV-...`

A estrutura atual é:

### Email

| Coluna | Campo |
|---|---|
| A | Data |
| B | Realizado por |
| C | Empresa |
| D | Qtde |
| E | Email Resposta |
| F | Supervisor |
| G | Obs |
| H | Historico Externo |
| I | Mensagem |
| J | Assunto |
| K | ID REGISTRO |

### SMS

| Coluna | Campo |
|---|---|
| A | Data |
| B | Realizado por |
| C | Empresa |
| D | Qtde |
| E | Supervisor |
| F | Obs |
| G | Historico Externo |
| H | Mensagem |
| I | ID REGISTRO |

O código possui compatibilidade com planilhas existentes que ainda não tenham essa coluna: quando `ID REGISTRO` estiver ausente, o módulo cria o cabeçalho na próxima coluna disponível antes de gravar o novo registro.

## Integridade de gravação

A gravação de E-mail e SMS foi reforçada para cenários de uso simultâneo.

O módulo recarrega o intervalo utilizado antes de inserir o registro e verifica se o `ID REGISTRO` gerado apareceu na célula esperada.

Se a confirmação falhar, existem novas tentativas limitadas. Caso a gravação continue sem confirmação, o sistema informa o erro ao usuário em vez de apresentar um salvamento como concluído.

Essa proteção reduz o risco de inconsistências em uso concorrente, mas não transforma o Excel em um banco de dados transacional.

## Indicadores de carregamento

Os módulos Configurações, Histórico e Resumo possuem indicação de carregamento e conclusão da atualização.

Os formulários E-mail e SMS também mantêm seus estados de processamento durante operações de leitura e gravação.

## Cache

Os arquivos JavaScript principais utilizam versionamento de consulta para evitar que o Excel Web mantenha versões antigas dos scripts após um deploy.

## Estado de validação

Na validação funcional atual, os módulos:

- E-mail;
- SMS;
- Histórico;
- Resumo

foram testados no Excel Web e estão funcionando.

A etapa seguinte é concluir a validação funcional da tela de **Configurações** antes do fechamento definitivo da próxima versão estável.

## Histórico de alterações

### 05/10/2026

- Documentação atualizada para refletir a arquitetura e os módulos atuais;
- Histórico documentado com filtros, ordenação, duplicação e navegação;
- Resumo documentado;
- Configurações documentadas;
- Identificação automática do usuário documentada;
- Confirmação de gravação e `ID REGISTRO` documentados;
- Compatibilidade com planilhas antigas documentada;
- Proteção de gravação concorrente documentada;
- Estado de validação registrado para a próxima versão alpha.


# Atualização de implementação — 06/10/2026

## Fechamento da validação recente

A validação funcional mais recente confirmou o funcionamento dos módulos principais no Excel Web:

- E-mail;
- SMS;
- Histórico;
- Resumo;
- Configurações.

O problema recente de identificação automática do usuário no módulo E-mail foi diagnosticado como **cache do navegador**. Após limpar o cache e testar em outro navegador, o SSO voltou a funcionar normalmente.

**Conclusão:** não houve necessidade de alterar a lógica de autenticação, o manifest.xml ou a configuração do Microsoft 365/Entra ID para corrigir esse incidente.

## Validação de e-mail

O campo **E-mail Resposta** possui validação básica de formato antes do salvamento e também ao cadastrar novos valores na configuração.

A regra aceita formatos normais como:

    nome@empresa.com

incluindo domínios como .com, .net, .org e outros formatos válidos segundo a verificação estrutural utilizada.

A validação confirma o **formato**, não a existência real da caixa postal ou do domínio.

Mensagem utilizada quando o formato não é aceito:

    ⚠ Informe um e-mail válido. Exemplo: nome@empresa.com

## Identificação automática e contingência

O comportamento estabelecido para **Realizado por** permanece:

1. Tentar identificar automaticamente o usuário Microsoft 365 por Office.auth.getAccessToken();
2. Se a plataforma não disponibilizar a identificação, permitir preenchimento manual;
3. Registrar o estado da identificação na interface;
4. Não expor tokens, credenciais ou informações sensíveis.

O incidente de cache de 06/10/2026 foi tratado como problema de ambiente/local e não como falha estrutural do SSO.

## Compatibilidade com registros existentes

A aplicação mantém compatibilidade com planilhas que ainda não possuam ID REGISTRO.

Quando a coluna não existe, o módulo cria o cabeçalho na próxima coluna disponível antes de gravar o novo registro.

Estruturas atuais:

### Email

| Coluna | Campo |
|---|---|
| A | Data |
| B | Realizado por |
| C | Empresa |
| D | Qtde |
| E | Email Resposta |
| F | Supervisor |
| G | Obs |
| H | Historico Externo |
| I | Mensagem |
| J | Assunto |
| K | ID REGISTRO |

### SMS

| Coluna | Campo |
|---|---|
| A | Data |
| B | Realizado por |
| C | Empresa |
| D | Qtde |
| E | Supervisor |
| F | Obs |
| G | Historico Externo |
| H | Mensagem |
| I | ID REGISTRO |

## Controle de concorrência

A gravação continua protegida por recarga do intervalo utilizado, geração do identificador e confirmação da célula esperada.

O sistema realiza tentativas limitadas e somente apresenta o salvamento como concluído quando a gravação é confirmada.

Essa proteção reduz conflitos em uso simultâneo, sem tratar o Excel como banco transacional.

## Cache dos scripts

Os arquivos JavaScript utilizam parâmetros de versionamento na URL para reduzir o risco de execução de versões antigas após deploy.

Caso o navegador continue utilizando uma versão antiga, a limpeza de cache e/ou teste em outro navegador é uma medida operacional válida para confirmar o comportamento atual.

## Versionamento Alpha

A linha alpha-1.05 foi atualizada para incorporar o estado atual validado do projeto.

O fechamento desta atualização inclui:

- validação dos módulos principais;
- validação do SSO após limpeza de cache;
- validação de formato de e-mail;
- compatibilidade com ID REGISTRO;
- proteção de gravação;
- Histórico;
- Resumo;
- Configurações;
- separação de mensagens E-mail/SMS;
- limite configurável de SMS;
- operação exclusivamente no Excel Web.

## Histórico de alterações

### 06/10/2026

- Atualização do estado documentado do projeto;
- Confirmação do funcionamento do SSO após limpeza de cache;
- Registro do incidente como problema de cache do navegador;
- Documentação da validação de formato de e-mail;
- Consolidação das regras de compatibilidade e gravação;
- Atualização do estado da linha Alpha.
