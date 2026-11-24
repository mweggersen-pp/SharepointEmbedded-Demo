import { IInputs, IOutputs } from "./generated/ManifestTypes";
import { CaseFile, DocumentClient, SearchHit, SharePointEmbedded, SharePointEmbeddedProps } from "./SharePointEmbedded";
import * as React from "react";

interface ActionResult {
  Result?: string;
}

interface DataverseActionResponse {
  json(): Promise<ActionResult>;
}

interface DataverseActionClient {
  execute(request: object): Promise<DataverseActionResponse>;
}

function dataverseActions(webApi: ComponentFramework.WebApi): DataverseActionClient {
  return webApi as unknown as DataverseActionClient;
}

function serviceMessage(caught: unknown): Error {
  const message = caught instanceof Error ? caught.message : "The document request failed.";
  if (/not find|not a valid|Resource not found|does not exist/i.test(message)) {
    return new Error(
      "The document actions are not registered in this environment yet. Container creation, list, preview, download, and Copilot search call new_EnsureCaseContainer, new_ListCaseDocuments, new_PreviewCaseDocument, new_DownloadCaseDocument, and new_SearchCaseDocuments."
    );
  }
  return new Error(message);
}

async function callAction(
  context: ComponentFramework.Context<IInputs>,
  operationName: string,
  parameters: Record<string, string>
): Promise<unknown> {
  const parameterTypes: Record<string, { typeName: string; structuralProperty: 1 }> = {};
  Object.keys(parameters).forEach((key) => {
    parameterTypes[key] = { typeName: "Edm.String", structuralProperty: 1 };
  });

  const request = {
    ...parameters,
    getMetadata: () => ({
      boundParameter: null,
      parameterTypes,
      operationName,
      operationType: 0,
    }),
  };

  try {
    const response = await dataverseActions(context.webAPI).execute(request);
    const body = await response.json();
    if (!body.Result) {
      return {};
    }
    return JSON.parse(body.Result) as unknown;
  } catch (caught) {
    throw serviceMessage(caught);
  }
}

function readCaseId(context: ComponentFramework.Context<IInputs>): string {
  const mode = context.mode as ComponentFramework.Mode & { contextInfo?: { entityId?: string } };
  return (mode.contextInfo?.entityId ?? "").replace(/[{}]/g, "");
}

function createClient(context: ComponentFramework.Context<IInputs>): DocumentClient {
  return {
    async ensure(caseId: string): Promise<string> {
      const payload = (await callAction(context, "new_EnsureCaseContainer", { CaseId: caseId })) as {
        containerId?: string;
      };
      if (!payload.containerId) {
        throw new Error("Container creation did not return an id.");
      }
      return payload.containerId;
    },
    async list(containerId: string): Promise<CaseFile[]> {
      const payload = (await callAction(context, "new_ListCaseDocuments", { ContainerId: containerId })) as {
        files?: CaseFile[];
      };
      return payload.files ?? [];
    },
    async preview(containerId: string, itemId: string): Promise<string> {
      const payload = (await callAction(context, "new_PreviewCaseDocument", {
        ContainerId: containerId,
        ItemId: itemId,
      })) as { previewUrl?: string };
      if (!payload.previewUrl) {
        throw new Error("Preview did not return a URL.");
      }
      return payload.previewUrl;
    },
    async download(containerId: string, itemId: string): Promise<string> {
      const payload = (await callAction(context, "new_DownloadCaseDocument", {
        ContainerId: containerId,
        ItemId: itemId,
      })) as { downloadUrl?: string };
      if (!payload.downloadUrl) {
        throw new Error("Download did not return a URL.");
      }
      return payload.downloadUrl;
    },
    async search(containerId: string, query: string): Promise<SearchHit[]> {
      const payload = (await callAction(context, "new_SearchCaseDocuments", {
        ContainerId: containerId,
        Query: query,
      })) as { hits?: SearchHit[] };
      return payload.hits ?? [];
    },
  };
}

export class CaseDocuments implements ComponentFramework.ReactControl<IInputs, IOutputs> {
  private notifyOutputChanged: () => void;
  private ensuredContainerId = "";
  private boundCaseId = "";

  constructor() {
    // React controls do not keep instance state here.
  }

  public init(
    context: ComponentFramework.Context<IInputs>,
    notifyOutputChanged: () => void,
    state: ComponentFramework.Dictionary
  ): void {
    void context;
    void state;
    this.notifyOutputChanged = notifyOutputChanged;
  }

  public updateView(context: ComponentFramework.Context<IInputs>): React.ReactElement {
    const recordId = readCaseId(context);
    if (this.boundCaseId !== recordId) {
      this.boundCaseId = recordId;
      this.ensuredContainerId = "";
    }
    const storedId = context.parameters.sampleProperty.raw ?? "";
    const containerId = this.ensuredContainerId !== "" ? this.ensuredContainerId : storedId;
    const props: SharePointEmbeddedProps = {
      containerId,
      caseId: readCaseId(context),
      client: createClient(context),
      onContainerId: (id: string) => {
        if (!id || this.ensuredContainerId === id) {
          return;
        }
        this.ensuredContainerId = id;
        if (recordId) {
          void context.webAPI
            .updateRecord("incident", recordId, { new_sharepointcontainer: id })
            .then(
              () => undefined,
              () => undefined
            );
        }
        this.notifyOutputChanged();
      },
    };
    return React.createElement(SharePointEmbedded, props);
  }

  public getOutputs(): IOutputs {
    if (!this.ensuredContainerId) {
      return {};
    }
    return { sampleProperty: this.ensuredContainerId };
  }

  public destroy(): void {
    this.notifyOutputChanged = () => undefined;
  }
}
