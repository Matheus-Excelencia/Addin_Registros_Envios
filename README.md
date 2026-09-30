# 📋 Add-in — Registros de Envios

Office Add-in para **Excel Web** desenvolvido para registrar e organizar os envios de E-mail e SMS/URA realizados no sistema **Konfiansa**.

O Add-in disponibiliza um painel lateral dentro do Excel para registrar os envios realizados e manter um histórico centralizado no arquivo `Registros de Envios.xlsx`.

---

# 🚀 Como instalar no Excel Web

## 1. Abra o Excel Web

Abra o **Excel para a Web / Microsoft 365** e o arquivo:

```
Registros de Envios.xlsx
```

## 2. Abra os Suplementos

No Excel Web, acesse:

**Início → Suplementos → Avançado**

Dependendo da versão do Excel, a opção também pode aparecer em:

**Inserir → Suplementos → Mais Suplementos**

## 3. Carregue o Add-in

Selecione:

**Carregar meu suplemento**

Depois escolha:

**Carregar de um arquivo**

## 4. Selecione o `manifest.xml`

Selecione o arquivo `manifest.xml` localizado na pasta principal do projeto.

O manifesto contém as configurações do Add-in e o endereço da aplicação web hospedada.

## 5. Abra o Add-in

Depois de carregá-lo, abra **Registros de Envios**.

O painel lateral apresenta:

- 📧 **E-mail**
- 📱 **SMS**

> Caso a opção **Carregar meu suplemento** não esteja disponível, o carregamento de Add-ins personalizados pode estar restrito pelas configurações do Microsoft 365 da organização.

---

# 🎯 Objetivo do projeto

O objetivo do projeto é manter um controle centralizado dos envios realizados no **Konfiansa**, especificamente através da área **CAMPANHA**:

- 📧 **E-mail em Lote**
- 📱 **SMS/URA em Lote**

O Add-in **não realiza o disparo das mensagens no Konfiansa**. O envio continua sendo realizado diretamente no sistema Konfiansa.

O Add-in é utilizado para:

- Registrar os envios realizados;
- Padronizar os registros;
- Manter histórico;
- Organizar as informações;
- Facilitar consultas futuras;
- Centralizar os registros de E-mail e SMS no Excel.

---

# 🔄 Fluxo do processo

```
                         KONFIANSA
                             │
                             ▼
                          CAMPANHA
                             │
                  ┌──────────┴──────────┐
                  ▼                     ▼
             E-mail em Lote       SMS/URA em Lote
                  │                     │
                  ▼                     ▼
             Envio realizado       Envio realizado
                  │                     │
                  └──────────┬──────────┘
                             ▼
                    REGISTROS DE ENVIOS
                             │
                  ┌──────────┴──────────┐
                  ▼                     ▼
               Aba Email             Aba SMS
```

---

# 📧 E-mail em Lote

No Konfiansa:

```
CAMPANHA → E-mail em Lote
```

Após realizar o envio, o usuário abre:

```
Registros de Envios
        ↓
      E-mail
```

O envio realizado é então registrado na aba **Email**.

## Estrutura da aba Email

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

## Regras do formulário de E-mail

- **Data:** registra automaticamente a data e a hora no momento do salvamento.
- **Realizado por:** tenta identificar automaticamente o usuário conectado ao Microsoft 365.
- **Empresa:** obrigatório; permite selecionar uma empresa existente ou informar uma nova.
- **Qtde:** quantidade obrigatória.
- **Email Resposta:** obrigatório; permite selecionar um e-mail existente ou informar um novo.
- **Supervisor:** obrigatório; permite selecionar um supervisor existente ou informar um novo.
- **Obs:** opcional.
- **Historico Externo:** obrigatório.
- **Mensagem:** obrigatória; o usuário seleciona o nome da mensagem e o texto completo configurado é armazenado.
- **Assunto:** obrigatório.

---

# 📱 SMS/URA em Lote

No Konfiansa:

```
CAMPANHA → SMS/URA em Lote
```

Após realizar o envio, o usuário abre:

```
Registros de Envios
        ↓
       SMS
```

O envio realizado é então registrado na aba **SMS**.

## Estrutura da aba SMS

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

## Regras do formulário de SMS

- **Data:** registra automaticamente a data e a hora no momento do salvamento.
- **Realizado por:** tenta identificar automaticamente o usuário conectado ao Microsoft 365.
- **Empresa:** obrigatório; permite selecionar uma empresa existente ou informar uma nova.
- **Qtde:** quantidade obrigatória.
- **Supervisor:** obrigatório; permite selecionar um supervisor existente ou informar um novo.
- **Obs:** opcional.
- **Historico Externo:** obrigatório e armazenado na coluna G.
- **Mensagem:** obrigatória e armazenada na coluna H.

## Limite de caracteres do SMS

O módulo de SMS possui limite máximo de:

```
160 caracteres
```

O formulário apresenta um contador de caracteres.

Quando a mensagem ultrapassa 160 caracteres:

- O salvamento é bloqueado;
- A mensagem não é cortada automaticamente;
- O usuário precisa corrigir o texto ou selecionar outra mensagem.

> Mensagens que possuem variáveis, como `[NomCliente]`, podem ficar maiores após a substituição pelo valor real.

---

# ⚙️ Aba Config

A aba `Config` centraliza as informações utilizadas pelos formulários.

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

## Atualização da Config

Quando suportado pelo formulário, novos:

- Empresas;
- Supervisores;
- E-mails de resposta;

podem ser adicionados à aba `Config` para utilização futura.

---

# 💬 Mensagens

As mensagens de E-mail e SMS são independentes.

## Mensagens de E-mail

Utilizam:

```
NOME MENSAGEM EMAIL
TEXTO MENSAGEM EMAIL
```

O formulário apresenta o nome da mensagem e armazena o texto completo configurado.

## Mensagens de SMS

Utilizam:

```
NOME MENSAGEM SMS
TEXTO MENSAGEM SMS
```

O formulário apresenta o nome da mensagem e armazena o texto completo configurado.

As mensagens de SMS estão sujeitas à validação de 160 caracteres.

---

# 📁 Estrutura do projeto

```
Addin-Registros-Envios/
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

# 🧩 Principais arquivos

### `manifest.xml`

Arquivo principal de configuração do Office Add-in. Define identidade, permissões, ícones, URLs e configurações de integração com o Excel.

### `taskpane.html`

Tela inicial do painel lateral, onde o usuário escolhe entre E-mail e SMS.

### `taskpane.css`

Contém os estilos da tela inicial.

### `taskpane.js`

Controla a navegação inicial e a abertura dos módulos de E-mail e SMS.

### `email.html / email.css / email.js`

Módulo completo de registro de E-mail.

Responsável por carregar as configurações, validar os campos, identificar o usuário, registrar data/hora, salvar o registro e atualizar as configurações quando necessário.

### `sms.html / sms.css / sms.js`

Módulo completo de registro de SMS.

Responsável por carregar as configurações, validar os campos, controlar o limite de 160 caracteres, registrar data/hora, salvar o registro e atualizar as configurações quando necessário.

### `commands.html / commands.js`

Arquivos relacionados aos comandos configurados no Office Add-in.

### `assets/`

Contém os ícones utilizados pelo manifesto do Add-in.

---

# 📝 Fluxo completo

```
1. Realizar o envio no Konfiansa
                 ↓
2. Acessar CAMPANHA
                 ↓
3. Selecionar:
   - E-mail em Lote
   - SMS/URA em Lote
                 ↓
4. Envio realizado
                 ↓
5. Abrir Registros de Envios no Excel
                 ↓
6. Selecionar E-mail ou SMS
                 ↓
7. Preencher o formulário
                 ↓
8. Salvar
                 ↓
9. Registro armazenado na aba correspondente
```

---

# ☁️ Hospedagem

A aplicação web do Add-in é hospedada na **Vercel**.

O Excel utiliza o `manifest.xml` para identificar e carregar a aplicação hospedada.

Arquitetura:

```
Excel Web
    ↓
manifest.xml
    ↓
Aplicação Web hospedada
    ↓
Task Pane
    ↓
E-mail / SMS
    ↓
Pasta de trabalho do Excel
```

---

# 🛠️ Manutenção

Antes de realizar alterações:

1. Fazer uma cópia ou utilizar o controle de versão do Git.
2. Alterar somente o módulo necessário.
3. Testar a alteração no Excel Web.
4. Testar o E-mail.
5. Testar o SMS.
6. Testar o salvamento.
7. Conferir as colunas do Excel.
8. Conferir a aba `Config`.
9. Confirmar que os registros anteriores não foram sobrescritos.
10. Fazer o commit após os testes.

---

# ⚠️ Observações importantes

## Estrutura das planilhas

Não alterar a ordem das colunas das abas `Email` ou `SMS` sem revisar o código JavaScript.

## E-mail e SMS são independentes

Não misturar:

```
NOME MENSAGEM EMAIL
TEXTO MENSAGEM EMAIL
```

com:

```
NOME MENSAGEM SMS
TEXTO MENSAGEM SMS
```

Cada módulo possui sua própria configuração de mensagens.

## Historico Externo

O campo `Historico Externo` faz parte dos dois módulos:

- E-mail → coluna H
- SMS → coluna G

## Validação do SMS

A validação de 160 caracteres deve permanecer ativa, salvo se a regra de negócio for alterada intencionalmente.

---

# 🔐 Segurança

Nunca armazenar diretamente no código-fonte:

- Senhas;
- Tokens de acesso;
- Chaves privadas;
- Credenciais de API;
- Informações sensíveis de autenticação.

As credenciais devem ser armazenadas de forma segura e fora do código-fonte.

---

# 📦 Versionamento

O projeto é mantido em Git/GitHub.

Padrões recomendados para commits:

```
feat: nova funcionalidade
fix: correção de erro
docs: alterações na documentação
refactor: reestruturação do código
chore: manutenção
```

Exemplo:

```
docs: update README with installation and Konfiansa workflow
```

---

# 📝 Histórico de alterações

## Versão atual

- Estrutura inicial do Add-in para Excel Web;
- Painel lateral;
- Cadastro de E-mail;
- Cadastro de SMS;
- Integração com a aba `Config`;
- Registro automático de data e hora;
- Identificação do usuário do Microsoft 365;
- Cadastro de empresas;
- Cadastro de supervisores;
- Cadastro de e-mails de resposta;
- Mensagens independentes para E-mail e SMS;
- Validação de 160 caracteres para SMS;
- Campo Historico Externo no E-mail;
- Campo Historico Externo no SMS;
- Histórico Externo do SMS salvo na coluna G;
- Mensagem do SMS salva na coluna H;
- Documentação do fluxo Konfiansa → CAMPANHA → Registros de Envios.

---

# 📌 Informações do projeto

**Projeto:** Registros de Envios  
**Plataforma principal:** Excel Web / Microsoft 365  
**Interface:** Office Add-in / Task Pane  
**Sistema de origem:** Konfiansa  
**Processos:** E-mail em Lote e SMS/URA em Lote  
**Módulos:** E-mail e SMS  
**Hospedagem:** Vercel  
**Repositório:** Matheus-Excelencia/Addin_Registros_Envios

---

# 👨‍💻 Estrutura de desenvolvimento

O projeto deve manter a separação entre os módulos:

```
taskpane
    ↓
Navegação inicial

email
    ↓
Registro de E-mail

sms
    ↓
Registro de SMS

Config
    ↓
Dados auxiliares e mensagens
```

Alterações futuras devem respeitar essa separação para reduzir o risco de impactos entre os módulos.

---
