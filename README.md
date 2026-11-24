# SharepointEmbedded-Demo

Case documents control for SharePoint Embedded. One container per Dynamics case. The case column `new_sharepointcontainer` stores the container id.

## Replace the scaffold

Copy these files over the PCF project that `pac pcf init` created:

- `CaseDocuments/index.ts`
- `CaseDocuments/SharePointEmbedded.tsx`
- `CaseDocuments/css/SharePointEmbedded.css`

In `CaseDocuments/ControlManifest.Input.xml`, change the stylesheet path to `css/SharePointEmbedded.css` and, inside the `<control>` element, add:

```xml
<feature-usage>
  <uses-feature name="Utility" required="true" />
  <uses-feature name="WebAPI" required="true" />
</feature-usage>
```

Keep the existing `sampleProperty` binding. That property is the container id.

Then build and push to mpcustomdemo:

```powershell
npm run build
pac org who
pac pcf push --publisher-prefix new
```

`pac org who` must show `https://mpcustomdemo2.crm19.dynamics.com` before the push.

## What the control calls

The control does not call Microsoft Graph from the browser. It calls four unbound Dataverse custom APIs. Each one returns a string output parameter named `Result` containing JSON.

| Action | Input | `Result` JSON |
|---|---|---|
| `new_EnsureCaseContainer` | `CaseId` | `{ "containerId" }` |
| `new_ListCaseDocuments` | `ContainerId` | `{ "files": [{ "id", "name", "size", "modifiedOn" }] }` |
| `new_PreviewCaseDocument` | `ContainerId`, `ItemId` | `{ "previewUrl" }` |
| `new_DownloadCaseDocument` | `ContainerId`, `ItemId` | `{ "downloadUrl" }` |
| `new_SearchCaseDocuments` | `ContainerId`, `Query` | `{ "hits": [{ "itemId", "name", "webUrl", "text" }] }` |

List, preview, and download use Microsoft Graph on the case container. Copilot search calls `POST /copilot/retrieval` with `dataSource` `sharePointEmbedded`, container type `332fda6e-b080-4793-9b64-a7c91166f03d`, and the case container id. Discoverability is currently off on `kunde-demo`. Copilot search returns hits only after discoverability is enabled for that container type.

The owning app is `20201f3e-941f-4590-b09a-0d7623cb3844`. The plugin reads `new_containertypeid`, `new_owningappid`, `new_tenantid`, and `new_graphclientsecret`.

Saving a case creates its container. Opening the documents tab on an older case creates the container then. The container id is written to `new_sharepointcontainer`. `Register-CaseDocuments.ps1` builds the plugin, registers the five actions, and publishes the control to mpcustomdemo.
