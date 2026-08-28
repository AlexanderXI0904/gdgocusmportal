# GDGoC Microsoft Graph Email Sender

A simple Flask-based bulk email sender using:

- Microsoft Graph API
- Microsoft Entra ID
- MSAL
- Excel `.xlsx` / `.xlsm`
- HTML email
- Dynamic Excel column insertion
- CC management
- WYSIWYG email editor
- HTML source editor
- Email preview
- Microsoft Graph importance settings

Designed for local/team use.

---

# Features

## Microsoft Authentication

The application authenticates through Microsoft Entra ID using MSAL.

Required:

- Client ID
- Tenant ID
- Microsoft Graph `Mail.Send` delegated permission

The application does not require an app password or SMTP.

---

# Excel

The application accepts:

- `.xlsx`
- `.xlsm`

The workbook is stored locally inside:

```text
uploads/