# Azure AD App Registration Guide (for IT Admin)

SHL Medical (`shl-group.com`) blocks all public Microsoft application IDs.
This tool requires a **company-registered Azure AD application**.

Please send this document to your IT / Azure AD administrator.

---

## App Registration Steps

### 1. Create App Registration

1. Open [Azure Portal](https://portal.azure.com) → **Microsoft Entra ID** → **App registrations**
2. Click **New registration**
3. Settings:
   - **Name**: `URL Batch Downloader` (or any name)
   - **Supported account types**: **Accounts in this organizational directory only** (Single tenant)
   - **Redirect URI**: Platform = **Mobile and desktop applications**
     - URI: `http://localhost`
4. Click **Register**
5. Copy the **Application (client) ID** → give this to the user

### 2. Enable Public Client Flow

1. Go to **Authentication**
2. Under **Advanced settings** → **Allow public client flows** → set to **Yes**
3. Save

### 3. API Permissions (Delegated)

1. Go to **API permissions** → **Add a permission**
2. Select **Microsoft Graph** → **Delegated permissions**
3. Add:
   - `Files.Read.All`
   - `Sites.Read.All`
4. Click **Grant admin consent for [tenant]**

### 4. Give Client ID to User

The user enters the **Application (client) ID** in the tool's **Client ID** field, then clicks **Sign in with Microsoft**.

---

## Troubleshooting

| Error | Solution |
|-------|----------|
| AADSTS700016 | App not registered in tenant – complete steps above |
| AADSTS53003 | Conditional Access – ensure app is allowed; user completes MFA in browser |
| AADSTS65001 | Admin consent not granted – IT must click "Grant admin consent" |

---

## Minimum Permissions Rationale

- `Files.Read.All` – read and download SharePoint / OneDrive files via Microsoft Graph
- `Sites.Read.All` – access SharePoint site content

No write permissions are required.
