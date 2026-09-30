# 📋 Add-in — Sending Records

Office Add-in for **Excel Web** designed to record and organize email and SMS/IVR batch campaigns performed in the **Konfiansa** system.

The Add-in provides a side panel inside Excel where users can register completed campaigns and keep a centralized history in the `Registros de Envios.xlsx` workbook.

---

# 🚀 How to Install in Excel Web

## 1. Open Excel Web

Open **Excel for the Web / Microsoft 365** and open:

```
Registros de Envios.xlsx
```

## 2. Open Add-ins

In Excel Web, go to:

**Home → Add-ins → Advanced**

Depending on the Excel version, the option may also appear under:

**Insert → Add-ins → More Add-ins**

## 3. Load the Add-in

Select:

**Upload My Add-in**

Then choose:

**Upload from File**

## 4. Select `manifest.xml`

Select the `manifest.xml` file located in the root folder of this project.

The manifest contains the Add-in configuration and the address of the hosted web application.

## 5. Open the Add-in

After loading it, open **Registros de Envios**.

The side panel provides:

- 📧 **Email**
- 📱 **SMS**

> If **Upload My Add-in** is not available, custom Add-in loading may be restricted by the organization's Microsoft 365 settings.

---

# 🎯 Project Purpose

The purpose of this project is to maintain a centralized record of campaigns performed in **Konfiansa**, specifically through the **CAMPANHA** area:

- 📧 **Email em Lote**
- 📱 **SMS/URA em Lote**

The Add-in **does not send messages through Konfiansa**. The actual campaign is performed in Konfiansa.

The Add-in is used to:

- Record completed campaigns;
- Standardize campaign records;
- Maintain history;
- Organize information;
- Facilitate future consultation;
- Keep email and SMS records centralized in Excel.

---

# 🔄 Process Flow

```
                         KONFIANSA
                             │
                             ▼
                          CAMPANHA
                             │
                  ┌──────────┴──────────┐
                  ▼                     ▼
            Email em Lote        SMS/URA em Lote
                  │                     │
                  ▼                     ▼
            Campaign sent        Campaign sent
                  │                     │
                  └──────────┬──────────┘
                             ▼
                    REGISTROS DE ENVIOS
                             │
                  ┌──────────┴──────────┐
                  ▼                     ▼
               Email tab             SMS tab
```

---

# 📧 Email — Batch Campaigns

In Konfiansa:

```
CAMPANHA → Email em Lote
```

After the campaign is performed, the user opens:

```
Registros de Envios
        ↓
      Email
```

The completed campaign is then recorded in the **Email** worksheet.

### Email worksheet structure

| Column | Field |
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

### Email form rules

- **Data:** automatically records date and time when saved.
- **Realizado por:** attempts to identify the Microsoft 365 user automatically.
- **Empresa:** required; existing values can be selected and new values can be entered.
- **Qtde:** required quantity.
- **Email Resposta:** required; existing values can be selected or a new value can be entered.
- **Supervisor:** required; existing values can be selected or a new value can be entered.
- **Obs:** optional.
- **Historico Externo:** required.
- **Mensagem:** required; the user selects the message name, while the complete configured text is stored.
- **Assunto:** required.

---

# 📱 SMS — Batch Campaigns

In Konfiansa:

```
CAMPANHA → SMS/URA em Lote
```

After the campaign is performed, the user opens:

```
Registros de Envios
        ↓
       SMS
```

The completed campaign is then recorded in the **SMS** worksheet.

### SMS worksheet structure

| Column | Field |
|---|---|
| A | Data |
| B | Realizado por |
| C | Empresa |
| D | Qtde |
| E | Supervisor |
| F | Obs |
| G | Historico Externo |
| H | Mensagem |

### SMS form rules

- **Data:** automatically records date and time when saved.
- **Realizado por:** attempts to identify the Microsoft 365 user automatically.
- **Empresa:** required; existing values can be selected and new values can be entered.
- **Qtde:** required quantity.
- **Supervisor:** required; existing values can be selected or a new value can be entered.
- **Obs:** optional.
- **Historico Externo:** required and stored in column G.
- **Mensagem:** required and stored in column H.

## SMS character limit

The SMS module enforces a maximum of:

```
160 characters
```

The form displays a character counter.

If the message exceeds 160 characters:

- Saving is blocked;
- The message is not automatically truncated;
- The user must correct the message or select another one.

> Messages containing variables such as `[NomCliente]` may become longer after the real value is inserted.

---

# ⚙️ Config Worksheet

The `Config` worksheet centralizes the information used by the forms.

| Column | Field |
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

## Automatic configuration updates

When supported by the form, newly entered:

- Companies;
- Supervisors;
- Reply emails;

can be added to the `Config` worksheet for future use.

---

# 💬 Messages

Email and SMS messages are independent.

## Email messages

Use:

```
NOME MENSAGEM EMAIL
TEXTO MENSAGEM EMAIL
```

The form displays the message name and stores the complete configured text.

## SMS messages

Use:

```
NOME MENSAGEM SMS
TEXTO MENSAGEM SMS
```

The form displays the message name and stores the complete configured text.

SMS messages are subject to the 160-character validation.

---

# 📁 Project Structure

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

# 🧩 Main Files

### `manifest.xml`

Main Office Add-in configuration file. Defines the Add-in identity, permissions, icons, URLs, and Excel integration settings.

### `taskpane.html`

Initial side-panel interface where the user chooses Email or SMS.

### `taskpane.css`

Styles for the initial side panel.

### `taskpane.js`

Controls the initial navigation and opens the Email and SMS modules.

### `email.html / email.css / email.js`

Complete Email registration module.

Responsible for loading configuration data, validating fields, identifying the user, recording date/time, saving the record, and updating configuration data when required.

### `sms.html / sms.css / sms.js`

Complete SMS registration module.

Responsible for loading configuration data, validating fields, enforcing the 160-character limit, recording date/time, saving the record, and updating configuration data when required.

### `commands.html / commands.js`

Command-related files used by the Office Add-in configuration.

### `assets/`

Contains the Add-in icons used by the Office manifest.

---

# 📝 Complete Workflow

```
1. Perform the campaign in Konfiansa
                 ↓
2. Open CAMPANHA
                 ↓
3. Select:
   - Email em Lote
   - SMS/URA em Lote
                 ↓
4. Campaign is sent
                 ↓
5. Open Registros de Envios in Excel
                 ↓
6. Select Email or SMS
                 ↓
7. Fill in the registration form
                 ↓
8. Save
                 ↓
9. Record is stored in the appropriate worksheet
```

---

# ☁️ Hosting

The Add-in web application is hosted on **Vercel**.

The Excel Add-in uses `manifest.xml` to identify and load the hosted application.

Architecture:

```
Excel Web
    ↓
manifest.xml
    ↓
Hosted Web Application
    ↓
Task Pane
    ↓
Email / SMS
    ↓
Excel Workbook
```

---

# 🛠️ Maintenance

Before making changes:

1. Create a backup or use Git version control.
2. Change only the required module.
3. Test the change in Excel Web.
4. Test Email.
5. Test SMS.
6. Test saving.
7. Verify the Excel columns.
8. Verify the `Config` worksheet.
9. Confirm existing records are not overwritten.
10. Commit the change after testing.

---

# ⚠️ Important Notes

## Worksheet structure

Do not change the order of columns in the `Email` or `SMS` worksheets without reviewing the JavaScript code.

## Email and SMS are independent

Do not mix:

```
NOME MENSAGEM EMAIL
TEXTO MENSAGEM EMAIL
```

with:

```
NOME MENSAGEM SMS
TEXTO MENSAGEM SMS
```

Each module has its own message configuration.

## Historico Externo

The `Historico Externo` field is currently part of both modules:

- Email → column H
- SMS → column G

## SMS validation

The 160-character validation must remain enabled unless the business rule is intentionally changed.

---

# 🔐 Security

Never store the following directly in the source code:

- Passwords;
- Access tokens;
- Private keys;
- API credentials;
- Sensitive authentication information.

Credentials should be stored securely and outside the source code.

---

# 📦 Version Control

The project is maintained in Git/GitHub.

Recommended commit prefixes:

```
feat: new functionality
fix: bug fix
docs: documentation changes
refactor: code restructuring
chore: maintenance
```

Example:

```
docs: update README with installation and Konfiansa workflow
```

---

# 📝 Change History

## Current version

- Initial Excel Web Add-in structure;
- Side panel interface;
- Email registration;
- SMS registration;
- `Config` worksheet integration;
- Automatic date and time;
- Microsoft 365 user identification;
- Company registration;
- Supervisor registration;
- Reply email registration;
- Independent Email and SMS messages;
- SMS 160-character validation;
- External History field for Email;
- External History field for SMS;
- SMS External History saved in column G;
- SMS Message saved in column H;
- Documentation of the Konfiansa → CAMPANHA → Registros de Envios workflow.

---

# 📌 Project Information

**Project:** Registros de Envios  
**Primary platform:** Excel Web / Microsoft 365  
**Interface:** Office Add-in / Task Pane  
**Source system:** Konfiansa  
**Campaign processes:** Email em Lote and SMS/URA em Lote  
**Modules:** Email and SMS  
**Hosting:** Vercel  
**Repository:** Matheus-Excelencia/Addin_Registros_Envios

---

# 👨‍💻 Development Structure

Keep the project separated by module:

```
taskpane
    ↓
Initial navigation

email
    ↓
Email registration

sms
    ↓
SMS registration

Config
    ↓
Auxiliary data and messages
```

Future changes should preserve this separation to reduce the risk of impacting other modules.

---

# ✅ Release Checklist

Before considering a new version ready:

- [ ] Add-in opens in Excel Web
- [ ] Initial screen works
- [ ] Email module opens
- [ ] SMS module opens
- [ ] Date and time are recorded
- [ ] User identification works
- [ ] Company field works
- [ ] Supervisor field works
- [ ] Reply email field works
- [ ] External History works
- [ ] Message field works
- [ ] Subject field works
- [ ] Quantity field works
- [ ] SMS 160-character limit works
- [ ] Record is saved to the correct worksheet
- [ ] Config updates correctly when required
- [ ] Email and SMS messages remain independent
- [ ] Existing records are not overwritten
- [ ] README is updated
- [ ] Changes are committed to Git
