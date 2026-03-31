# gopas-jsdays-2026

**SharePoint Document Access Granting Solution** -- Conference demo for GOPAS JS Days 2026.

A button in a SharePoint document library that grants a predefined user read access to the selected document and logs the action to an audit list -- powered by an Azure Function and Microsoft Graph API.

---

## Architecture

```
SharePoint Document Library
        |
        |  User selects a document, clicks "Grant Access"
        v
SPFx ListView Command Set (GrantAccessCommandSet)
        |
        |  POST { siteId, listId, itemId }
        v
Azure Function (grantAccess)
        |
        |--- Graph API: GET  /sites/{siteId}/drive
        |--- Graph API: POST /drives/{driveId}/items/{itemId}/invite  (read access)
        |--- Graph API: POST /sites/{siteId}/lists/AuditLog/items     (audit entry)
        |
        v
Response --> SPFx shows success/error dialog
```

---

## Project Structure

| Folder | What it does | Technology |
|--------|-------------|------------|
| `server/` | Backend API -- grants permissions and writes audit log | Azure Functions v4, TypeScript, Microsoft Graph |
| `spfx/` | SharePoint UI -- "Grant Access" button in document libraries | SPFx 1.22, ListView Command Set |
| `.github/workflows/` | CI/CD pipeline for the Azure Function | GitHub Actions |

---

## Key Files

### Server

| File | Purpose |
|------|---------|
| `src/config.ts` | Reads and validates environment variables |
| `src/graphClient.ts` | Creates an authenticated Microsoft Graph client |
| `src/grantAccess.ts` | HTTP trigger -- the main business logic |
| `src/index.ts` | Entry point, imports the function |
| `host.json` | Azure Functions runtime configuration |

### SPFx

| File | Purpose |
|------|---------|
| `src/extensions/grantAccess/GrantAccessCommandSet.ts` | ListView Command Set -- shows button, calls Azure Function |
| `src/extensions/grantAccess/GrantAccessCommandSet.manifest.json` | Component manifest with `GRANT_ACCESS` command definition |
| `config/config.json` | Bundle configuration pointing to the extension |
| `config/package-solution.json` | Solution packaging settings |
| `sharepoint/assets/elements.xml` | CustomAction registration (document libraries) |
| `sharepoint/assets/ClientSideInstance.xml` | Client-side component instance registration |

---

## Prerequisites

- **Node.js** >= 22.14.0
- **npm** (comes with Node.js)
- **Azure subscription** with permissions to create Function Apps
- **Microsoft 365 tenant** with SharePoint Online
- **Entra ID App Registration** with `Sites.ReadWrite.All` (application permission, admin consented)
- **PnP PowerShell** (for site collection app catalog and custom action management)

---

## Quick Start

### Server (Azure Function)

```bash
cd server
npm install
npm run build
```

### SPFx (SharePoint Extension)

```bash
cd spfx
npm install
npx heft test --clean --production && npx heft package-solution --production
```

The `.sppkg` file will be in `spfx/sharepoint/solution/spfx.sppkg`.

---

## Environment Variables (Azure Function)

Configure these in the Function App's **Application Settings**:

| Variable | Description | Example |
|----------|-------------|---------|
| `TENANT_ID` | Entra ID tenant ID | `ce5e571a-c44e-...` |
| `CLIENT_ID` | App registration client ID | `227cbdd6-42a6-...` |
| `CLIENT_SECRET` | App registration client secret | `3yj8Q~nQ83~...` |
| `SHAREPOINT_SITE_URL` | Target SharePoint site URL | `https://contoso.sharepoint.com/sites/Demo` |
| `TARGET_USER_EMAIL` | Email of the user to grant access to | `user@contoso.com` |
| `AUDIT_LIST_NAME` | Name of the SharePoint audit list | `AuditLog` |

---

## Deployment

### Azure Function

Deployment is automated via GitHub Actions. On every push to `main`, the workflow:

1. Installs dependencies
2. Compiles TypeScript
3. Prunes dev dependencies
4. Deploys to the `accessgrant` Function App via zip deploy

### SPFx Extension

1. Build the `.sppkg` package
2. Upload to a **site collection app catalog** (recommended) or tenant app catalog
3. Install the app on the target site
4. Register the custom action on the document library (if not auto-activated)

```powershell
Connect-PnPOnline -Url https://contoso.sharepoint.com/sites/Demo -Interactive -ClientId <your-client-id>

# Create site collection app catalog (one-time)
Add-PnPSiteCollectionAppCatalog -Site https://contoso.sharepoint.com/sites/Demo

# If custom action needs manual registration
Add-PnPCustomAction -Title "GrantAccess" -Name "GrantAccess" `
  -Location "ClientSideExtension.ListViewCommandSet.CommandBar" `
  -ClientSideComponentId "a1b2c3d4-e5f6-4a7b-8c9d-0e1f2a3b4c5d" `
  -RegistrationType List -RegistrationId 101
```

---

## CORS

Add your SharePoint domain to the Function App's CORS settings:

**Function App > API > CORS** --> add `https://contoso.sharepoint.com`

---

## Graph API Permissions

The Entra ID app registration requires:

| Permission | Type | Purpose |
|-----------|------|---------|
| `Sites.ReadWrite.All` | Application | Read site/drive info, create sharing links, write to audit list |

Admin consent is required.

---

## Tech Stack

| Layer | Technology | Version |
|-------|-----------|---------|
| Backend runtime | Azure Functions | v4 |
| Backend language | TypeScript | 6.x |
| Backend auth | @azure/identity | ClientSecretCredential |
| API client | @microsoft/microsoft-graph-client | 3.x |
| Frontend framework | SharePoint Framework (SPFx) | 1.22.2 |
| Frontend extension | ListView Command Set | - |
| CI/CD | GitHub Actions | - |
| Hosting | Azure (Consumption plan) | West Europe |

---

## License

MIT
