import { IInputs, IOutputs } from "./generated/ManifestTypes";
import { createCaseFiles, errorMessage, GraphSettings } from "./CaseDocumentsGraph";
import { AccessEntry, CaseAccess, ContactChoice } from "./CaseAccessPanel";
import { CaseContainer, SharePointEmbedded, SharePointEmbeddedProps } from "./SharePointEmbedded";
import { caseTextFor } from "./case-documents-text";
import * as React from "react";

interface ActionResult {
  Result?: string;
}

interface DataverseActionClient {
  execute(request: object): Promise<{ json(): Promise<ActionResult> }>;
}

interface EnvironmentVariableRow {
  schemaname: string;
  defaultvalue?: string;
  environmentvariabledefinition_environmentvariablevalue?: { value?: string }[];
}

const settingNames = ["new_owningappid", "new_tenantid", "new_containertypeid"];

async function callAction(webApi: ComponentFramework.WebApi, operationName: string, parameters: Record<string, string>): Promise<unknown> {
  const parameterTypes: Record<string, { typeName: string; structuralProperty: 1 }> = {};
  Object.keys(parameters).forEach((key) => {
    parameterTypes[key] = { typeName: "Edm.String", structuralProperty: 1 };
  });

  const request = {
    ...parameters,
    getMetadata: () => ({ boundParameter: null, parameterTypes, operationName, operationType: 0 }),
  };

  try {
    const response = await (webApi as unknown as DataverseActionClient).execute(request);
    const body = await response.json();
    return body.Result ? (JSON.parse(body.Result) as unknown) : {};
  } catch (caught) {
    const message = errorMessage(caught, "The request failed.");
    if (/not find|Resource not found|does not exist/i.test(message)) {
      throw new Error(`${operationName} is not registered in this environment. Run Register-CaseDocuments.ps1.`);
    }
    throw new Error(message);
  }
}

async function loadGraphSettings(context: ComponentFramework.Context<IInputs>): Promise<GraphSettings> {
  const filter = settingNames.map((name) => `schemaname eq '${name}'`).join(" or ");
  const [variables, user] = await Promise.all([
    context.webAPI.retrieveMultipleRecords(
      "environmentvariabledefinition",
      `?$select=schemaname,defaultvalue&$filter=${filter}&$expand=environmentvariabledefinition_environmentvariablevalue($select=value)`
    ),
    context.webAPI.retrieveRecord("systemuser", context.userSettings.userId.replace(/[{}]/g, ""), "?$select=domainname"),
  ]);

  const value = (name: string): string => {
    const row = (variables.entities as EnvironmentVariableRow[]).find((entity) => entity.schemaname === name);
    const current = row?.environmentvariabledefinition_environmentvariablevalue?.[0]?.value ?? row?.defaultvalue;
    if (!current) {
      throw new Error(`Set a value for environment variable ${name}.`);
    }
    return current;
  };

  return {
    clientId: value("new_owningappid"),
    tenantId: value("new_tenantid"),
    containerTypeId: value("new_containertypeid"),
    loginHint: (user.domainname as string | undefined) ?? "",
  };
}

function readCaseId(context: ComponentFramework.Context<IInputs>): string {
  const mode = context.mode as ComponentFramework.Mode & { contextInfo?: { entityId?: string } };
  return (mode.contextInfo?.entityId ?? "").replace(/[{}]/g, "");
}

export class CaseDocuments implements ComponentFramework.ReactControl<IInputs, IOutputs> {
  private props: Omit<SharePointEmbeddedProps, "caseId">;

  public init(context: ComponentFramework.Context<IInputs>): void {
    const webApi = context.webAPI;
    const files = createCaseFiles(() => loadGraphSettings(context));

    const access: CaseAccess = {
      async list(caseId) {
        const result = (await callAction(webApi, "new_ListCaseAccess", { CaseId: caseId })) as { entries?: AccessEntry[] };
        return result.entries ?? [];
      },
      async pickContact(): Promise<ContactChoice | null> {
        const [picked] = await context.utils.lookupObjects({ entityTypes: ["contact"], defaultEntityType: "contact", allowMultiSelect: false });
        if (!picked) {
          return null;
        }
        const id = picked.id.replace(/[{}]/g, "");
        const record = await webApi.retrieveRecord("contact", id, "?$select=emailaddress1");
        const email = typeof record.emailaddress1 === "string" ? record.emailaddress1 : "";
        return { id, name: picked.name ?? "Contact", email };
      },
      async grant(caseId, contactId, role) {
        await callAction(webApi, "new_GrantCaseContactAccess", { CaseId: caseId, ContactId: contactId, Role: role });
      },
      async ensureGuest(caseId, email) {
        const result = (await callAction(webApi, "new_EnsureCaseGuest", { CaseId: caseId, Email: email })) as { created?: boolean };
        return result.created === true;
      },
      async shareByEmail(caseId, email, role) {
        await callAction(webApi, "new_ShareCaseByEmail", { CaseId: caseId, Email: email, Role: role });
      },
      async update(caseId, permissionId, role) {
        await callAction(webApi, "new_UpdateCaseAccess", { CaseId: caseId, PermissionId: permissionId, Role: role });
      },
      async remove(caseId, permissionId) {
        await callAction(webApi, "new_RemoveCaseAccess", { CaseId: caseId, PermissionId: permissionId });
      },
    };

    this.props = {
      text: caseTextFor(context.userSettings.languageId),
      files,
      access,
      ensureContainer: async (caseId) => {
        const result = (await callAction(webApi, "new_EnsureCaseContainer", { CaseId: caseId })) as Partial<CaseContainer>;
        if (!result.containerId) {
          throw new Error("Container creation did not return an id.");
        }
        return { containerId: result.containerId, role: result.role ?? "reader" };
      },
    };
  }

  public updateView(context: ComponentFramework.Context<IInputs>): React.ReactElement {
    return React.createElement(SharePointEmbedded, { ...this.props, caseId: readCaseId(context) });
  }

  public getOutputs(): IOutputs {
    return {};
  }

  public destroy(): void {
    // Nothing to release.
  }
}
