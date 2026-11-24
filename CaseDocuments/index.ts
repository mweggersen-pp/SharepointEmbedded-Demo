import { IInputs, IOutputs } from "./generated/ManifestTypes";
import { CaseFile, DocumentClient, SearchHit, SharePointEmbedded, SharePointEmbeddedProps } from "./SharePointEmbedded";
import * as React from "react";

interface ActionResult {
  Result?: string;
}

function serviceMessage(caught: unknown): Error {
  const message = caught instanceof Error ? caught.message : "The document request failed.";
  if (/not find|not a valid|Resource not found|does not exist/i.test(message)) {
    return new Error(
      "The document actions are not registered in this environment yet. List, preview, download, and Copilot search call new_ListCaseDocuments, new_PreviewCaseDocument, new_DownloadCaseDocument, and new_SearchCaseDocuments."
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
    const response = await context.webAPI.execute(request);
    const body = (await response.json()) as ActionResult;
    if (!body.Result) {
      return {};
    }
    return JSON.parse(body.Result) as unknown;
  } catch (caught) {
    throw serviceMessage(caught);
  }
}

function createClient(context: ComponentFramework.Context<IInputs>): DocumentClient {
  return {
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
    const props: SharePointEmbeddedProps = {
      containerId: context.parameters.sampleProperty.raw ?? "",
      client: createClient(context),
    };
    return React.createElement(SharePointEmbedded, props);
  }

  public getOutputs(): IOutputs {
    return {};
  }

  public destroy(): void {
    this.notifyOutputChanged = () => undefined;
  }
}
