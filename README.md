# SharepointEmbedded-Demo

Case documents for Dynamics 365. Each case gets its own SharePoint Embedded container, and its id is stored in `new_sharepointcontainer`. On the case form you can upload, preview, download and search the case's files with Copilot. You can also give external contacts access.

## How it works

- **The control** (`CaseDocuments/`) signs in as the current user with MSAL. It calls Microsoft Graph from the browser for:
  - listing, previewing (`nb=true`), downloading and uploading files
  - Copilot search (`/copilot/retrieval`), filtered to the case's own files
- **The plugin** (`CaseDocumentsPlugin/`) runs with the app's own permissions. It only creates containers and manages who has access.
  - When a case is created, it creates the container. If a case has no container yet, the container is created when the tab opens.
  - Colleagues get the same access in the container that they have to the case in Dynamics: Edit if they can write the case, Read if they can only read it. This is checked when someone opens the tab and whenever the case owner changes.
  - External contacts are invited as Microsoft Entra guests, using the email address on their contact record. They get Read or Edit, chosen from the dropdown.
- **The solution** (`CaseDocumentsActions/src`, unpacked) contains:
  - the signed plugin
  - two asynchronous steps: case Create, and case Update of `ownerid`
  - the sign-in redirect page `new_msalredirect.html`
  - these custom APIs, each returning JSON in `Result`:

| Action | Input |
|---|---|
| `new_EnsureCaseContainer` | `CaseId` |
| `new_ListCaseAccess` | `CaseId` |
| `new_GrantCaseContactAccess` | `CaseId`, `ContactId`, `Role` (`reader` or `writer`) |
| `new_ShareCaseByEmail` | `CaseId`, `Email`, `Role` |
| `new_UpdateCaseAccess` | `CaseId`, `PermissionId`, `Role` |
| `new_RemoveCaseAccess` | `CaseId`, `PermissionId` (external contacts only) |

The plugin reads these environment variables: `new_tenantid`, `new_owningappid`, `new_graphclientsecret` and `new_containertypeid`. The control reads all of them except the secret.

## Setup

In app registration `20201f3e-941f-4590-b09a-0d7623cb3844`:

1. Under Authentication, add the single-page application redirect URI `https://mpcustomdemo2.crm19.dynamics.com/WebResources/new_msalredirect.html`.
2. Add these Microsoft Graph permissions and grant admin consent:
   - delegated `FileStorageContainer.Selected`
   - application `FileStorageContainer.Selected`, `User.Invite.All` and `User.Read.All`
3. Give the container type both delegated and application permissions for the app.
4. To invite external contacts, turn on external sharing for the container type and the tenant.

## Deploy

1. Copy `CaseDocuments/` into the PCF project. In `ControlManifest.Input.xml`, use `css/SharePointEmbedded.css` and add:

```xml
<feature-usage>
  <uses-feature name="Utility" required="true" />
  <uses-feature name="WebAPI" required="true" />
</feature-usage>
```

2. In the PCF project, run:

```powershell
npm install @azure/msal-browser@4.30.0
npm run build
pac pcf push --publisher-prefix new
```

3. From this repo, run `./Register-CaseDocuments.ps1`. It needs the .NET SDK and a `pac` profile signed in to `https://mpcustomdemo2.crm19.dynamics.com`. It builds the plugin, copies the publisher from `SharePointEmbededDemo`, packs the solution with `pac solution pack` and imports it.
