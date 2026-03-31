# Tutorial: Build a SharePoint Document Access Granting Solution

A complete, step-by-step guide for beginners. By the end, you will have a working solution where a SharePoint user clicks a button in a document library, an Azure Function grants read access to a specific user via Microsoft Graph, and an audit log entry is created.

---

## Table of Contents

- [What You Will Learn](#what-you-will-learn)
- [What You Will Build](#what-you-will-build)
- [Prerequisites](#prerequisites)
- [Part 1: Azure Setup](#part-1-azure-setup)
  - [Step 1: Create a Resource Group](#step-1-create-a-resource-group)
  - [Step 2: Create an Azure Function App](#step-2-create-an-azure-function-app)
  - [Step 3: Create an Entra ID App Registration](#step-3-create-an-entra-id-app-registration)
  - [Step 4: Grant API Permissions](#step-4-grant-api-permissions)
  - [Step 5: Create a Client Secret](#step-5-create-a-client-secret)
- [Part 2: SharePoint Setup](#part-2-sharepoint-setup)
  - [Step 6: Create a SharePoint Site](#step-6-create-a-sharepoint-site)
  - [Step 7: Create the AuditLog List](#step-7-create-the-auditlog-list)
  - [Step 8: Upload a Test Document](#step-8-upload-a-test-document)
- [Part 3: Build the Azure Function (Server)](#part-3-build-the-azure-function-server)
  - [Step 9: Initialize the Project](#step-9-initialize-the-project)
  - [Step 10: Configure TypeScript](#step-10-configure-typescript)
  - [Step 11: Create host.json](#step-11-create-hostjson)
  - [Step 12: Create the Configuration Module](#step-12-create-the-configuration-module)
  - [Step 13: Create the Graph Client](#step-13-create-the-graph-client)
  - [Step 14: Create the Grant Access Function](#step-14-create-the-grant-access-function)
  - [Step 15: Create the Entry Point](#step-15-create-the-entry-point)
  - [Step 16: Build and Verify](#step-16-build-and-verify)
- [Part 4: Build the SPFx Extension](#part-4-build-the-spfx-extension)
  - [Step 17: Scaffold the SPFx Project](#step-17-scaffold-the-spfx-project)
  - [Step 18: Write the ListView Command Set](#step-18-write-the-listview-command-set)
  - [Step 19: Configure the Manifest](#step-19-configure-the-manifest)
  - [Step 20: Configure the Bundle](#step-20-configure-the-bundle)
  - [Step 21: Configure the SharePoint Assets](#step-21-configure-the-sharepoint-assets)
  - [Step 22: Build the SPFx Package](#step-22-build-the-spfx-package)
- [Part 5: Deploy Everything](#part-5-deploy-everything)
  - [Step 23: Configure Environment Variables](#step-23-configure-environment-variables)
  - [Step 24: Set Up GitHub Actions CI/CD](#step-24-set-up-github-actions-cicd)
  - [Step 25: Configure CORS](#step-25-configure-cors)
  - [Step 26: Deploy SPFx to SharePoint](#step-26-deploy-spfx-to-sharepoint)
  - [Step 27: Register the Custom Action](#step-27-register-the-custom-action)
- [Part 6: Test the Solution](#part-6-test-the-solution)
  - [Step 28: End-to-End Test](#step-28-end-to-end-test)
- [Troubleshooting](#troubleshooting)
- [Concepts Reference](#concepts-reference)

---

## What You Will Learn

| Concept | What it teaches |
|---------|----------------|
| Azure Functions v4 | How to create a serverless HTTP API with TypeScript |
| Microsoft Graph API | How to grant permissions on SharePoint items and write to lists |
| Entra ID (Azure AD) | How to register an app and authenticate with client credentials |
| SharePoint Framework | How to build a ListView Command Set extension |
| GitHub Actions | How to automate deployment of an Azure Function |
| PnP PowerShell | How to manage SharePoint site collection app catalogs and custom actions |

---

## What You Will Build

```
+---------------------------+         +-----------------------------+
|  SharePoint Document      |         |  Azure Function App         |
|  Library                  |         |  (Consumption Plan)         |
|                           |  POST   |                             |
|  [Grant Access] button ---|-------->|  /api/grant-access          |
|  (SPFx Extension)         |         |                             |
|                           |  200 OK |  1. Get drive for site      |
|  "Access granted!" <------|---------|  2. Invite user (read)      |
|                           |         |  3. Write to AuditLog list  |
+---------------------------+         +-------------|---------------+
                                                    |
                                                    | Microsoft Graph API
                                                    v
                                      +-----------------------------+
                                      |  Microsoft 365              |
                                      |  - SharePoint permissions   |
                                      |  - AuditLog list entries    |
                                      +-----------------------------+
```

---

## Prerequisites

Before you start, make sure you have the following installed and ready:

### Software

| Tool | Version | How to install |
|------|---------|---------------|
| Node.js | >= 22.14.0 | Download from [nodejs.org](https://nodejs.org) |
| npm | Comes with Node.js | Installed automatically |
| Git | Any recent version | Download from [git-scm.com](https://git-scm.com) |
| Visual Studio Code | Any recent version | Download from [code.visualstudio.com](https://code.visualstudio.com) |
| PnP PowerShell | Latest | `Install-Module PnP.PowerShell -Scope CurrentUser` |
| Yeoman + SPFx generator | Latest | `npm install -g yo @microsoft/generator-sharepoint` |

### Accounts & Access

| Requirement | Why you need it |
|------------|----------------|
| Azure subscription | To host the Azure Function |
| Microsoft 365 tenant | To use SharePoint Online |
| Global Admin or SharePoint Admin | To create app registrations and grant API permissions |
| GitHub account | To set up CI/CD with GitHub Actions |

### Verify your setup

Open a terminal and run:

```bash
node --version    # Should show v22.x.x or higher
npm --version     # Should show 10.x.x or higher
git --version     # Should show any version
yo --version      # Should show 5.x.x
```

---

## Part 1: Azure Setup

### Step 1: Create a Resource Group

A Resource Group is a container for related Azure resources. All resources for this project will live in one group.

1. Go to [portal.azure.com](https://portal.azure.com)
2. Click **Create a resource** (top left, or search bar)
3. Search for **Resource group** and click **Create**
4. Fill in:
   - **Subscription**: Select your subscription
   - **Resource group name**: `jsdays2026`
   - **Region**: `West Europe` (or your nearest region)
5. Click **Review + create** --> **Create**

**Why a Resource Group?** It lets you manage, monitor, and delete all related resources together. When the demo is over, you delete one resource group and everything is cleaned up.

---

### Step 2: Create an Azure Function App

The Function App hosts your serverless API.

1. In the Azure Portal, click **Create a resource**
2. Search for **Function App** and click **Create**
3. Fill in the **Basics** tab:
   - **Subscription**: Your subscription
   - **Resource Group**: `jsdays2026`
   - **Function App name**: `accessgrant` (must be globally unique -- Azure will append a random suffix to the URL)
   - **Runtime stack**: `Node.js`
   - **Version**: `22`
   - **Region**: `West Europe`
   - **Operating System**: `Windows`
   - **Hosting plan**: `Consumption (Serverless)` -- this is the cheapest option, ideal for demos
4. Click **Review + create** --> **Create**
5. Wait for the deployment to complete (about 2 minutes)

**Why Consumption Plan?** You only pay when the function runs. For a demo with low traffic, this is essentially free (1 million free executions per month).

**Note your Function App URL.** It will look like:
```
https://accessgrant-XXXXXXX.westeurope-01.azurewebsites.net
```
You will need this later.

---

### Step 3: Create an Entra ID App Registration

The Azure Function needs an identity to call Microsoft Graph. An App Registration provides this.

1. In the Azure Portal, go to **Microsoft Entra ID** (formerly Azure Active Directory)
2. Click **App registrations** in the left menu
3. Click **+ New registration**
4. Fill in:
   - **Name**: `JSDays2026`
   - **Supported account types**: `Accounts in this organizational directory only` (single tenant)
   - **Redirect URI**: Leave empty (we use client credentials, not interactive login)
5. Click **Register**

After registration, you'll see the **Overview** page. Note down two values:

| Value | Where to find it | Example |
|-------|-----------------|---------|
| **Application (client) ID** | Overview page | `227cbdd6-42a6-441b-bba0-b72b39ff656b` |
| **Directory (tenant) ID** | Overview page | `ce5e571a-c44e-4b16-aa88-c8dc2fa5a367` |

**Why an App Registration?** Microsoft Graph requires authentication. The app registration gives your function a "username and password" (client ID and secret) to authenticate as an application (not a user).

---

### Step 4: Grant API Permissions

Your app needs permission to read/write SharePoint data via Graph.

1. In your App Registration, click **API permissions** in the left menu
2. Click **+ Add a permission**
3. Select **Microsoft Graph**
4. Select **Application permissions** (not Delegated)
5. Search for `Sites.ReadWrite.All` and check it
6. Click **Add permissions**
7. Click **Grant admin consent for [your tenant]** (you need admin rights for this)
8. Confirm when prompted

You should see a green checkmark next to the permission.

**Why `Sites.ReadWrite.All`?** This permission allows the function to:
- Read site and drive information
- Create sharing invitations (grant access)
- Write items to the AuditLog list

**Why Application permissions?** The function runs without a signed-in user (it's a backend service), so it needs application-level permissions, not delegated (user) permissions.

---

### Step 5: Create a Client Secret

The client secret is the "password" for your app registration.

1. In your App Registration, click **Certificates & secrets** in the left menu
2. Click **+ New client secret**
3. Fill in:
   - **Description**: `JSDays2026`
   - **Expires**: `6 months` (sufficient for a demo)
4. Click **Add**
5. **Immediately copy the secret Value** (you won't be able to see it again!)

| Value | Where to find it |
|-------|-----------------|
| **Client Secret** | Certificates & secrets > Value column |

**Store this securely.** Never commit it to source code. It goes into Azure Function App Settings only.

---

## Part 2: SharePoint Setup

### Step 6: Create a SharePoint Site

1. Go to your SharePoint Admin Center (`https://[tenant]-admin.sharepoint.com`)
2. Click **Sites** > **Active sites** > **+ Create**
3. Choose **Team site** (or Communication site)
4. Fill in:
   - **Site name**: `JSDays2026`
   - **Site address**: `JSDays2026`
5. Click **Create**

Your site URL will be: `https://[tenant].sharepoint.com/sites/JSDays2026`

---

### Step 7: Create the AuditLog List

This list stores a record every time access is granted.

1. Go to your new site
2. Click **Site contents** (gear icon > Site contents)
3. Click **+ New** > **List**
4. Choose **Blank list**
5. Name it `AuditLog` and click **Create**
6. Add two columns:

| Column name | Type | How to add |
|-------------|------|-----------|
| `GrantedTo` | Single line of text | Click **+ Add column** > **Single line of text** |
| `Timestamp` | Single line of text | Click **+ Add column** > **Single line of text** |

**Why Single line of text for Timestamp?** The Azure Function sends the timestamp as an ISO 8601 string (e.g., `2026-03-31T07:30:00.000Z`). Using a text column keeps things simple for a demo.

The `Title` column already exists by default -- the function uses it to store the item ID.

---

### Step 8: Upload a Test Document

1. Go to the **Documents** library on your site
2. Click **+ New** > **Upload** > **Files**
3. Upload any file (e.g., `Document.docx`)

You now have a document to test with.

---

## Part 3: Build the Azure Function (Server)

### Step 9: Initialize the Project

Create the server folder and initialize it:

```bash
mkdir server
cd server
npm init -y
```

Install the dependencies:

```bash
npm install @azure/functions @microsoft/microsoft-graph-client @azure/identity
npm install -D typescript @types/node azure-functions-core-tools
```

| Package | Purpose |
|---------|---------|
| `@azure/functions` | Azure Functions v4 programming model for Node.js |
| `@microsoft/microsoft-graph-client` | Official Microsoft Graph SDK |
| `@azure/identity` | Authentication library (provides `ClientSecretCredential`) |
| `typescript` | TypeScript compiler |
| `@types/node` | TypeScript type definitions for Node.js |
| `azure-functions-core-tools` | Local development tools for Azure Functions |

Add a build script to `package.json`:

```json
{
  "main": "dist/index.js",
  "scripts": {
    "build": "tsc"
  }
}
```

**Why `main: "dist/index.js"`?** Azure Functions v4 uses the `main` field to find your entry point. After TypeScript compiles, the output goes to `dist/`.

---

### Step 10: Configure TypeScript

Create `tsconfig.json`:

```json
{
  "compilerOptions": {
    "target": "ES2020",
    "module": "commonjs",
    "rootDir": "src",
    "outDir": "dist",
    "strict": true,
    "esModuleInterop": true
  }
}
```

| Setting | What it does |
|---------|-------------|
| `target: "ES2020"` | Compiles to modern JavaScript (Azure Functions Node 22 supports this) |
| `module: "commonjs"` | Azure Functions requires CommonJS modules |
| `rootDir: "src"` | Source files live in `src/` |
| `outDir: "dist"` | Compiled files go to `dist/` |
| `strict: true` | Enables all strict type-checking options |
| `esModuleInterop: true` | Allows `import x from "y"` syntax with CommonJS modules |

---

### Step 11: Create host.json

Create `host.json` in the `server/` root (not inside `src/`):

```json
{
  "version": "2.0",
  "extensionBundle": {
    "id": "Microsoft.Azure.Functions.ExtensionBundle",
    "version": "[4.*, 5.0.0)"
  }
}
```

**Why host.json?** This file tells the Azure Functions runtime which version and extensions to use. The extension bundle includes bindings for HTTP triggers, timers, queues, etc.

---

### Step 12: Create the Configuration Module

Create `src/config.ts`:

```typescript
function requireEnv(name: string): string {
  const value = process.env[name];
  if (!value) {
    throw new Error(`Missing required environment variable: ${name}`);
  }
  return value;
}

export const TENANT_ID = requireEnv("TENANT_ID");
export const CLIENT_ID = requireEnv("CLIENT_ID");
export const CLIENT_SECRET = requireEnv("CLIENT_SECRET");
export const SHAREPOINT_SITE_URL = requireEnv("SHAREPOINT_SITE_URL");
export const TARGET_USER_EMAIL = requireEnv("TARGET_USER_EMAIL");
export const AUDIT_LIST_NAME = requireEnv("AUDIT_LIST_NAME");
```

**Why this pattern?** If any environment variable is missing, the function crashes immediately on startup with a clear error message. This is much better than getting a cryptic error later when the code tries to use an `undefined` value.

**Why environment variables?** Secrets like `CLIENT_SECRET` should never be in source code. Environment variables are configured in the Azure Portal and are encrypted at rest.

---

### Step 13: Create the Graph Client

Create `src/graphClient.ts`:

```typescript
import { ClientSecretCredential } from "@azure/identity";
import { Client } from "@microsoft/microsoft-graph-client";
import { TokenCredentialAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials";
import { TENANT_ID, CLIENT_ID, CLIENT_SECRET } from "./config";

export function getGraphClient(): Client {
  const credential = new ClientSecretCredential(TENANT_ID, CLIENT_ID, CLIENT_SECRET);

  const authProvider = new TokenCredentialAuthenticationProvider(credential, {
    scopes: ["https://graph.microsoft.com/.default"],
  });

  return Client.initWithMiddleware({ authProvider });
}
```

**What happens here, line by line:**

1. `ClientSecretCredential` -- creates a credential object using your tenant ID, client ID, and client secret. This is the "identity" of your app.
2. `TokenCredentialAuthenticationProvider` -- wraps the credential so the Graph SDK can use it to get access tokens automatically.
3. `Client.initWithMiddleware` -- creates a Graph client that automatically attaches the access token to every request.

**Why `https://graph.microsoft.com/.default`?** The `.default` scope means "give me all the permissions that were granted to this app in Entra ID" (i.e., `Sites.ReadWrite.All`).

---

### Step 14: Create the Grant Access Function

Create `src/grantAccess.ts`:

```typescript
import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { getGraphClient } from "./graphClient";
import { TARGET_USER_EMAIL, AUDIT_LIST_NAME } from "./config";

interface GrantAccessBody {
  siteId: string;
  listId: string;
  itemId: string;
}

async function grantAccess(req: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
  try {
    const { siteId, listId, itemId } = (await req.json()) as GrantAccessBody;
    const client = getGraphClient();

    // Step 1a: Get the default drive ID for the site
    console.log(`Getting default drive for site ${siteId}...`);
    const drive = await client.api(`/sites/${siteId}/drive`).get();
    const driveId = drive.id;

    // Step 1b: Grant read permission on the item via sharing invite
    console.log(`Granting read access to ${TARGET_USER_EMAIL} on item ${itemId}...`);
    await client.api(`/drives/${driveId}/items/${itemId}/invite`).post({
      recipients: [{ email: TARGET_USER_EMAIL }],
      roles: ["read"],
      requireSignIn: true,
      sendInvitation: false,
    });

    // Step 2: Create audit log entry in the SharePoint list
    console.log(`Writing audit log to list ${AUDIT_LIST_NAME}...`);
    await client.api(`/sites/${siteId}/lists/${AUDIT_LIST_NAME}/items`).post({
      fields: {
        Title: itemId,
        GrantedTo: TARGET_USER_EMAIL,
        Timestamp: new Date().toISOString(),
      },
    });

    return {
      status: 200,
      jsonBody: { success: true, itemId, grantedTo: TARGET_USER_EMAIL },
    };
  } catch (err: any) {
    return {
      status: 500,
      jsonBody: { error: err.message },
    };
  }
}

app.http("grantAccess", {
  methods: ["POST"],
  authLevel: "anonymous",
  route: "grant-access",
  handler: grantAccess,
});
```

**Let's break down each Graph API call:**

#### Call 1: Get the default drive

```
GET /sites/{siteId}/drive
```

Every SharePoint site has a default "drive" (the Documents library). We need the drive ID to work with files inside it.

#### Call 2: Grant read access via invite

```
POST /drives/{driveId}/items/{itemId}/invite
```

This creates a sharing invitation on the document:

| Parameter | Value | Why |
|-----------|-------|-----|
| `recipients` | `[{ email: "user@..." }]` | Who gets access |
| `roles` | `["read"]` | Read-only permission |
| `requireSignIn` | `true` | Recipient must sign in (security) |
| `sendInvitation` | `false` | Don't send an email notification |

#### Call 3: Write audit log

```
POST /sites/{siteId}/lists/AuditLog/items
```

Creates a new item in the AuditLog list with the item ID, who got access, and when.

#### Function registration

```typescript
app.http("grantAccess", {
  methods: ["POST"],
  authLevel: "anonymous",
  route: "grant-access",
  handler: grantAccess,
});
```

| Setting | What it means |
|---------|-------------|
| `methods: ["POST"]` | Only accepts POST requests |
| `authLevel: "anonymous"` | No function key required (the SPFx extension calls it without a key) |
| `route: "grant-access"` | URL will be `/api/grant-access` |

**Why anonymous auth?** For a demo, this keeps things simple. In production, you would use function keys or Azure AD authentication.

---

### Step 15: Create the Entry Point

Create `src/index.ts`:

```typescript
import "./grantAccess";
```

**Why just an import?** Azure Functions v4 uses a "code-first" approach. When `index.ts` imports `grantAccess.ts`, the `app.http(...)` call at the bottom of that file registers the function with the runtime. No `function.json` files needed.

---

### Step 16: Build and Verify

```bash
npm run build
```

You should see no errors. Check that `dist/` contains:

```
dist/
  config.js
  graphClient.js
  grantAccess.js
  index.js
```

Your Azure Function is ready.

---

## Part 4: Build the SPFx Extension

### Step 17: Scaffold the SPFx Project

From the repository root:

```bash
mkdir spfx
cd spfx
yo @microsoft/sharepoint
```

When prompted, choose:

| Prompt | Answer |
|--------|--------|
| Solution name | `gopas-jsdays-2026-spfx` (or accept default) |
| Which type of client-side component? | **Extension** |
| Which type of extension? | **ListView Command Set** |
| What is your Command Set name? | `GrantAccess` |

Wait for the scaffolding to complete and dependencies to install.

**What is a ListView Command Set?** It's a type of SPFx extension that adds custom buttons (commands) to the toolbar of SharePoint lists and document libraries. When a user selects items in the list, your code controls which buttons are visible and what happens when they're clicked.

---

### Step 18: Write the ListView Command Set

Replace the generated file `src/extensions/grantAccess/GrantAccessCommandSet.ts` with:

```typescript
import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  type Command,
  type IListViewCommandSetExecuteEventParameters,
  type ListViewStateChangedEventArgs
} from '@microsoft/sp-listview-extensibility';
import { Dialog } from '@microsoft/sp-dialog';
import { HttpClient, type HttpClientResponse } from '@microsoft/sp-http';

const LOG_SOURCE = 'GrantAccessCommandSet';
const FUNCTION_URL = 'https://YOUR-FUNCTION-APP.azurewebsites.net/api/grant-access';

export default class GrantAccessCommandSet extends BaseListViewCommandSet<{}> {

  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized');
    this.context.listView.listViewStateChangedEvent.add(this, this._onListViewStateChanged);
    return Promise.resolve();
  }

  private _onListViewStateChanged = (_args: ListViewStateChangedEventArgs): void => {
    const command: Command = this.tryGetCommand('GRANT_ACCESS');
    if (command) {
      // Show button only when exactly one item is selected
      command.visible = this.context.listView.selectedRows?.length === 1;
    }
    this.raiseOnChange();
  };

  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
      case 'GRANT_ACCESS':
        Log.info(LOG_SOURCE, 'Grant Access command executed');
        void this._grantAccess();
        break;
      default:
        throw new Error('Unknown command');
    }
  }

  private async _grantAccess(): Promise<void> {
    const row = this.context.listView.selectedRows![0];
    const siteId = this.context.pageContext.site.id.toString();
    const listId = this.context.pageContext.list!.id.toString();
    const itemId = row.getValueByName('UniqueId');

    try {
      const response: HttpClientResponse = await this.context.httpClient.post(
        FUNCTION_URL,
        HttpClient.configurations.v1,
        {
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ siteId, listId, itemId }),
        }
      );

      const result = await response.json();

      if (response.ok) {
        await Dialog.alert('Access granted successfully');
      } else {
        await Dialog.alert(`Error: ${result.error}`);
      }
    } catch (err: unknown) {
      const message = err instanceof Error ? err.message : String(err);
      Log.warn(LOG_SOURCE, `Grant access failed: ${message}`);
      await Dialog.alert(`Error: ${message}`);
    }
  }
}
```

**Important:** Replace `YOUR-FUNCTION-APP` in the `FUNCTION_URL` with your actual Azure Function URL from [Step 2](#step-2-create-an-azure-function-app).

**Key methods explained:**

| Method | When it runs | What it does |
|--------|-------------|-------------|
| `onInit()` | Extension loads | Subscribes to list view state changes |
| `_onListViewStateChanged()` | User selects/deselects items | Shows the button only when exactly 1 item is selected |
| `onExecute()` | User clicks a command button | Routes to the right handler based on command ID |
| `_grantAccess()` | "Grant Access" button clicked | Sends POST request to Azure Function, shows result dialog |

**Why `this.context.httpClient`?** SPFx provides a built-in HTTP client that handles authentication and CORS correctly. Never use `fetch()` directly in SPFx -- always use `HttpClient` or `AadHttpClient`.

**Why `UniqueId`?** Every item in SharePoint has a `UniqueId` (a GUID). This is the stable identifier used by the Graph API to reference files, unlike the integer `ID` which is list-scoped.

---

### Step 19: Configure the Manifest

Edit `src/extensions/grantAccess/GrantAccessCommandSet.manifest.json`:

```json
{
  "$schema": "https://developer.microsoft.com/json-schemas/spfx/command-set-extension-manifest.schema.json",
  "id": "a1b2c3d4-e5f6-4a7b-8c9d-0e1f2a3b4c5d",
  "alias": "GrantAccessCommandSet",
  "componentType": "Extension",
  "extensionType": "ListViewCommandSet",
  "version": "*",
  "manifestVersion": 2,
  "requiresCustomScript": false,
  "items": {
    "GRANT_ACCESS": {
      "title": { "default": "Grant Access" },
      "type": "command"
    }
  }
}
```

| Field | Purpose |
|-------|---------|
| `id` | Unique GUID identifying this component. You can generate your own. |
| `items.GRANT_ACCESS` | Defines a command named `GRANT_ACCESS` with the button label "Grant Access". This matches the `event.itemId` in `onExecute()`. |

---

### Step 20: Configure the Bundle

Edit `config/config.json` to point to your extension:

```json
{
  "$schema": "https://developer.microsoft.com/json-schemas/spfx-build/config.2.0.schema.json",
  "version": "2.0",
  "bundles": {
    "grant-access-command-set": {
      "components": [
        {
          "entrypoint": "./lib/extensions/grantAccess/GrantAccessCommandSet.js",
          "manifest": "./src/extensions/grantAccess/GrantAccessCommandSet.manifest.json"
        }
      ]
    }
  },
  "externals": {},
  "localizedResources": {}
}
```

**Why this file matters:** The SPFx build system uses `config.json` to know which components to bundle. The `entrypoint` points to the compiled JavaScript (in `lib/`), and the `manifest` points to the source manifest.

---

### Step 21: Configure the SharePoint Assets

These XML files tell SharePoint to register your extension as a custom action on document libraries.

#### `sharepoint/assets/elements.xml`

```xml
<?xml version="1.0" encoding="utf-8"?>
<Elements xmlns="http://schemas.microsoft.com/sharepoint/">
    <CustomAction
        Title="GrantAccess"
        RegistrationId="101"
        RegistrationType="List"
        Location="ClientSideExtension.ListViewCommandSet.CommandBar"
        ClientSideComponentId="a1b2c3d4-e5f6-4a7b-8c9d-0e1f2a3b4c5d">
    </CustomAction>
</Elements>
```

#### `sharepoint/assets/ClientSideInstance.xml`

```xml
<?xml version="1.0" encoding="utf-8"?>
<Elements xmlns="http://schemas.microsoft.com/sharepoint/">
    <ClientSideComponentInstance
        Title="GrantAccess"
        Location="ClientSideExtension.ListViewCommandSet.CommandBar"
        ListTemplateId="101"
        Properties="{}"
        ComponentId="a1b2c3d4-e5f6-4a7b-8c9d-0e1f2a3b4c5d" />
</Elements>
```

| Setting | Value | Meaning |
|---------|-------|---------|
| `RegistrationId` / `ListTemplateId` | `101` | Document Library template. Use `100` for generic lists. |
| `Location` | `ClientSideExtension.ListViewCommandSet.CommandBar` | Places the button in the top command bar |
| `ClientSideComponentId` / `ComponentId` | `a1b2c3d4-...` | Must match the `id` in your manifest |

#### Reference in `config/package-solution.json`

Make sure the `features` section references both XML files:

```json
"features": [
  {
    "title": "Application Extension - Deployment of custom action",
    "description": "Deploys a custom action with ClientSideComponentId association",
    "id": "bc933309-1340-43e9-a167-928815f554d8",
    "version": "1.0.0.0",
    "assets": {
      "elementManifests": [
        "elements.xml",
        "ClientSideInstance.xml"
      ]
    }
  }
]
```

---

### Step 22: Build the SPFx Package

```bash
cd spfx
npm install
npx heft test --clean --production && npx heft package-solution --production
```

The output file is `sharepoint/solution/spfx.sppkg`.

**What is an `.sppkg`?** It's a SharePoint Package -- a zip file containing your bundled JavaScript, manifest, and XML configuration. You upload this to a SharePoint app catalog to deploy your extension.

---

## Part 5: Deploy Everything

### Step 23: Configure Environment Variables

In the Azure Portal, go to your Function App > **Settings** > **Environment variables**.

Add these application settings:

| Name | Value |
|------|-------|
| `TENANT_ID` | Your tenant ID from [Step 3](#step-3-create-an-entra-id-app-registration) |
| `CLIENT_ID` | Your client ID from [Step 3](#step-3-create-an-entra-id-app-registration) |
| `CLIENT_SECRET` | Your client secret from [Step 5](#step-5-create-a-client-secret) |
| `SHAREPOINT_SITE_URL` | `https://[tenant].sharepoint.com/sites/JSDays2026` |
| `TARGET_USER_EMAIL` | The email of the user who should receive access |
| `AUDIT_LIST_NAME` | `AuditLog` |

Click **Save** and then **Restart** the Function App.

**Why restart?** Environment variable changes require a restart for the function runtime to pick them up.

---

### Step 24: Set Up GitHub Actions CI/CD

Create `.github/workflows/main_accessgrant.yml` in your repository:

```yaml
name: Build and deploy Node.js project to Azure Function App - accessgrant

on:
  push:
    branches:
      - main
  workflow_dispatch:

env:
  NODE_VERSION: '22.x'

jobs:
  build-and-deploy:
    runs-on: windows-latest
    permissions:
      id-token: write
      contents: read

    steps:
      - name: 'Checkout GitHub Action'
        uses: actions/checkout@v4

      - name: Setup Node ${{ env.NODE_VERSION }} Environment
        uses: actions/setup-node@v3
        with:
          node-version: ${{ env.NODE_VERSION }}

      - name: 'Build project'
        shell: pwsh
        run: |
          pushd './server'
          npm install
          npm run build --if-present
          npm prune --production
          popd

      - name: Login to Azure
        uses: azure/login@v2
        with:
          client-id: ${{ secrets.AZURE_CLIENT_ID }}
          tenant-id: ${{ secrets.AZURE_TENANT_ID }}
          subscription-id: ${{ secrets.AZURE_SUBSCRIPTION_ID }}

      - name: 'Run Azure Functions Action'
        uses: Azure/functions-action@v1
        id: fa
        with:
          app-name: 'accessgrant'
          slot-name: 'Production'
          package: server
```

**Important:** Replace the secret names with the ones generated by Azure when you configured GitHub deployment, or add them manually in GitHub > Settings > Secrets > Actions.

**What each step does:**

| Step | Purpose |
|------|---------|
| Checkout | Gets your code from GitHub |
| Setup Node | Installs Node.js 22 on the build runner |
| Build | Installs deps, compiles TypeScript, removes dev dependencies (saves ~300MB) |
| Login | Authenticates with Azure using OIDC (secure, no passwords stored) |
| Deploy | Zips the `server/` folder and deploys it to the Function App |

**Why `npm prune --production`?** Dev dependencies like `azure-functions-core-tools` are ~300MB. Removing them before deployment keeps the zip small and the deployment fast.

---

### Step 25: Configure CORS

Your SPFx extension (running on `sharepoint.com`) makes HTTP requests to your Azure Function (running on `azurewebsites.net`). Browsers block cross-origin requests by default.

1. In the Azure Portal, go to your Function App
2. Click **API** > **CORS** in the left menu
3. Add your SharePoint domain: `https://[tenant].sharepoint.com`
4. Click **Save**

**Why CORS?** Cross-Origin Resource Sharing is a browser security feature. Without it, the browser will block the POST request from SharePoint to your Azure Function. Adding your SharePoint domain tells the Function App "requests from this origin are allowed."

---

### Step 26: Deploy SPFx to SharePoint

#### Option A: Site Collection App Catalog (Recommended for demos)

A site collection app catalog limits the extension to a single site.

```powershell
# Connect to your site
Connect-PnPOnline -Url https://[tenant].sharepoint.com/sites/JSDays2026 -Interactive -ClientId <your-client-id>

# Create the site collection app catalog (one-time)
Add-PnPSiteCollectionAppCatalog -Site https://[tenant].sharepoint.com/sites/JSDays2026

# Upload and deploy the package
Add-PnPApp -Path "./spfx/sharepoint/solution/spfx.sppkg" -Scope Site -Publish -Overwrite

# Install on the site
Install-PnPApp -Identity "spfx-client-side-solution" -Scope Site
```

#### Option B: Tenant App Catalog

A tenant app catalog makes the extension available across all sites. Use this only if you want it available everywhere.

1. Go to the **SharePoint Admin Center** > **More features** > **Apps** > **App Catalog**
2. Upload `spfx.sppkg`
3. Do **NOT** check "Make this solution available to all sites" (unless you want that)
4. Go to your site > **Site contents** > **New** > **App** > install it

---

### Step 27: Register the Custom Action

If the "Grant Access" button doesn't appear after deploying the app, you need to register the custom action manually:

```powershell
Connect-PnPOnline -Url https://[tenant].sharepoint.com/sites/JSDays2026 -Interactive -ClientId <your-client-id>

# Check current custom actions
Get-PnPCustomAction -Scope Web

# If the GrantAccess action is missing, add it
Add-PnPCustomAction -Title "GrantAccess" -Name "GrantAccess" `
  -Location "ClientSideExtension.ListViewCommandSet.CommandBar" `
  -ClientSideComponentId "a1b2c3d4-e5f6-4a7b-8c9d-0e1f2a3b4c5d" `
  -RegistrationType List -RegistrationId 101
```

**If there's an old/wrong custom action**, remove it first:

```powershell
# List all custom actions to find the wrong one
Get-PnPCustomAction -Scope Web | Format-List Id, Title, ClientSideComponentId

# Remove by ID
Remove-PnPCustomAction -Identity "THE-GUID-HERE" -Scope Web -Force
```

---

## Part 6: Test the Solution

### Step 28: End-to-End Test

1. Go to `https://[tenant].sharepoint.com/sites/JSDays2026/Shared%20Documents/Forms/AllItems.aspx`
2. Select a document by clicking its checkbox
3. Look for the **"Grant Access"** button in the command bar (top toolbar)
4. Click it
5. Wait for the dialog

**Expected results:**

| What to check | Expected outcome |
|--------------|-----------------|
| Dialog message | "Access granted successfully" |
| AuditLog list | New item with Title (item GUID), GrantedTo (email), Timestamp (ISO date) |
| Document permissions | Target user now has read access |

**To verify permissions:**

1. Right-click the document > **Manage access**
2. You should see the target user listed with "Can view" permissions

**To verify the audit log:**

1. Click **AuditLog** in the site navigation
2. You should see a new entry with the item ID, granted email, and timestamp

---

## Troubleshooting

### Common Issues

| Symptom | Cause | Fix |
|---------|-------|-----|
| "Grant Access" button doesn't appear | Custom action not registered or wrong ComponentId | Run `Get-PnPCustomAction -Scope Web` and verify. Re-register if needed. |
| "Grant Access" button not visible | No item selected, or more than one selected | Select exactly one document |
| `ERR_NAME_NOT_RESOLVED` | Wrong Function URL in SPFx code | Check the Function App's default domain in Azure Portal and update `FUNCTION_URL` |
| `404 Not Found` | Function not registered (missing env vars causes startup crash) | Add all 6 environment variables and restart the Function App |
| `500 Internal Server Error` | Graph API call failed | Check the Function App's **Log stream** for detailed error messages |
| CORS error in browser console | SharePoint domain not added to CORS | Add `https://[tenant].sharepoint.com` in Function App > API > CORS |
| `"Failed to execute 'json' on 'Response'"` | Function returned non-JSON response (HTML error page) | Usually means the function URL is wrong or CORS is blocking |
| `"The field or property 'X' does not exist"` | Missing column in SharePoint list | Add the missing column to the AuditLog list |
| GitHub Actions build fails on `npm run test` | The default test script exits with code 1 | Remove the `npm run test` line from the workflow, or change the test script in `package.json` |
| GitHub Actions deploy fails with 500 | Package too large for zip deploy | Add `npm prune --production` before deploy to remove dev dependencies |
| No functions listed in Azure Portal | Environment variables missing, causing startup crash | Add all env vars, restart, check again |

### How to Debug

**Azure Function logs:**
1. Azure Portal > Function App > **Log stream**
2. Trigger the function and watch the real-time logs
3. The `console.log` statements in the code will appear here

**Browser console:**
1. Open browser DevTools (F12) > Console tab
2. Look for red errors when clicking "Grant Access"
3. Check the Network tab for the actual HTTP request/response

---

## Concepts Reference

### Azure Functions v4 Programming Model

Azure Functions v4 for Node.js uses a "code-first" approach:

```
Traditional (v3):            Code-first (v4):
function.json  ------>       app.http("name", { ... })
index.js                     handler function in same file
```

No `function.json` files needed. You register functions directly in code using `app.http()`, `app.timer()`, etc.

### Microsoft Graph API

Microsoft Graph is the unified API for all Microsoft 365 services. Every Graph request follows this pattern:

```
https://graph.microsoft.com/v1.0/{resource}
```

Examples used in this project:

| Operation | Method | Endpoint |
|-----------|--------|----------|
| Get site's default drive | GET | `/sites/{siteId}/drive` |
| Create sharing invitation | POST | `/drives/{driveId}/items/{itemId}/invite` |
| Create list item | POST | `/sites/{siteId}/lists/{listName}/items` |

### SPFx Extension Types

| Type | Where it appears | Use case |
|------|-----------------|----------|
| Application Customizer | Header/footer of every page | Notifications, banners, navigation |
| Field Customizer | Inside list columns | Custom column rendering |
| **ListView Command Set** | List/library toolbar | **Custom action buttons (this project)** |

### Deployment Options Comparison

| Approach | Scope | Best for |
|----------|-------|----------|
| Site Collection App Catalog | Single site | Demos, targeted solutions |
| Tenant App Catalog (manual install) | Specific sites | Team solutions |
| Tenant App Catalog (skip feature deployment) | All sites | Organization-wide tools |

---

> **Congratulations!** You've built a complete SharePoint + Azure solution that grants document access via Microsoft Graph. This pattern -- SPFx frontend calling an Azure Function backend that uses Graph API -- is a foundational architecture for many Microsoft 365 solutions.
