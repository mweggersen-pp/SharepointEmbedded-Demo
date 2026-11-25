import { PublicClientApplication } from "@azure/msal-browser";

const graphRoot = "https://graph.microsoft.com/v1.0";
const scopes = ["https://graph.microsoft.com/FileStorageContainer.Selected"];
const simpleUploadLimit = 4 * 1024 * 1024;
// Upload session chunks must be multiples of 320 KiB.
const uploadChunkSize = 320 * 1024 * 16;

export interface GraphSettings {
  clientId: string;
  tenantId: string;
  containerTypeId: string;
  loginHint: string;
}

export interface CaseFile {
  id: string;
  name: string;
  size?: number;
  modifiedOn?: string;
  webUrl?: string;
}

export interface SearchHit {
  itemId: string;
  name: string;
  text: string;
}

export type FileShareRole = "read" | "write";

export interface FileShare {
  key: string;
  permissionId: string;
  name: string;
  email: string;
  role: FileShareRole;
}

export interface CaseFiles {
  list(containerId: string): Promise<CaseFile[]>;
  preview(containerId: string, itemId: string): Promise<string>;
  download(containerId: string, itemId: string): Promise<string>;
  upload(containerId: string, file: File): Promise<void>;
  remove(containerId: string, itemId: string): Promise<void>;
  listShares(containerId: string, itemId: string): Promise<FileShare[]>;
  share(containerId: string, itemId: string, email: string, role: FileShareRole): Promise<void>;
  unshare(containerId: string, itemId: string, permissionId: string): Promise<void>;
  search(containerId: string, query: string): Promise<SearchHit[]>;
  signIn(): Promise<void>;
}

interface DriveItem {
  id: string;
  name: string;
  size?: number;
  lastModifiedDateTime?: string;
  webUrl?: string;
  file?: object;
  "@microsoft.graph.downloadUrl"?: string;
}

interface RetrievalHit {
  webUrl?: string;
  extracts?: { text?: string }[];
}

interface PermissionIdentity {
  user?: { displayName?: string; email?: string };
  siteUser?: { displayName?: string; email?: string; loginName?: string };
}

interface ItemPermission {
  id: string;
  roles?: string[];
  inheritedFrom?: object;
  grantedToV2?: PermissionIdentity;
  grantedToIdentitiesV2?: PermissionIdentity[];
  invitation?: { email?: string };
}

function toFileShares(permission: ItemPermission): FileShare[] {
  const roles = permission.roles ?? [];
  if (permission.inheritedFrom || roles.includes("owner")) {
    return [];
  }
  const role: FileShareRole = roles.includes("write") ? "write" : "read";
  const identities = permission.grantedToIdentitiesV2 ?? (permission.grantedToV2 ? [permission.grantedToV2] : []);
  const people = identities
    .map((identity) => identity.user ?? identity.siteUser)
    .filter((person): person is { displayName?: string; email?: string } => person !== undefined)
    .map((person) => ({ name: person.displayName ?? "", email: person.email ?? "" }));
  if (people.length === 0 && permission.invitation?.email) {
    people.push({ name: "", email: permission.invitation.email });
  }
  return people.map((person, index) => ({ key: `${permission.id}-${index}`, permissionId: permission.id, role, ...person }));
}

interface MsalSession {
  app: PublicClientApplication;
  settings: GraphSettings;
}

// Dataverse webAPI rejects with a plain { message } object rather than an Error.
export function errorMessage(caught: unknown, fallback: string): string {
  if (caught instanceof Error) {
    return caught.message;
  }
  if (typeof caught === "string" && caught) {
    return caught;
  }
  if (typeof caught === "object" && caught !== null) {
    if ("message" in caught && typeof caught.message === "string" && caught.message) {
      return caught.message;
    }
    if ("error" in caught) {
      return errorMessage(caught.error, fallback);
    }
    const details = JSON.stringify(caught);
    if (details && details !== "{}") {
      return `${fallback} ${details.slice(0, 400)}`;
    }
  }
  return fallback;
}

export function isSignInRequired(caught: unknown): boolean {
  return caught instanceof Error && caught.name === "SignInRequired";
}

function signInRequired(): Error {
  const error = new Error("Sign in to SharePoint to work with the documents on this case.");
  error.name = "SignInRequired";
  return error;
}

function formatNorwegianDate(iso: string): string {
  const date = new Date(iso);
  const day = date.toLocaleDateString("nb-NO", { day: "2-digit", month: "2-digit", year: "numeric" });
  const time = date.toLocaleTimeString("nb-NO", { hour: "2-digit", minute: "2-digit" });
  return `${day} kl. ${time}`;
}

function sameUrl(left: string, right: string): boolean {
  return decodeURIComponent(left).toLowerCase() === decodeURIComponent(right).toLowerCase();
}

async function graphError(response: Response): Promise<string> {
  try {
    const body = (await response.json()) as { error?: { message?: string } };
    return body.error?.message ?? `Microsoft Graph returned ${response.status}.`;
  } catch {
    return `Microsoft Graph returned ${response.status}.`;
  }
}

export function createCaseFiles(loadSettings: () => Promise<GraphSettings>): CaseFiles {
  let pending: Promise<MsalSession> | null = null;
  let session: MsalSession | null = null;

  const start = (): Promise<MsalSession> => {
    pending ??= (async () => {
      const settings = await loadSettings();
      const app = new PublicClientApplication({
        auth: {
          clientId: settings.clientId,
          authority: `https://login.microsoftonline.com/${settings.tenantId}`,
          redirectUri: `${window.location.origin}/WebResources/new_msalredirect.html`,
        },
        cache: { cacheLocation: "sessionStorage" },
      });
      await app.initialize();
      session = { app, settings };
      return session;
    })().catch((caught: unknown) => {
      pending = null;
      throw caught;
    });
    return pending;
  };

  const token = async (): Promise<string> => {
    const { app, settings } = await start();
    const hint = settings.loginHint.toLowerCase();
    const account = app.getAllAccounts().find((candidate) => candidate.username.toLowerCase() === hint);
    try {
      const result = account
        ? await app.acquireTokenSilent({ scopes, account })
        : await app.ssoSilent({ scopes, loginHint: settings.loginHint });
      return result.accessToken;
    } catch {
      throw signInRequired();
    }
  };

  const graph = async <T>(path: string, init: RequestInit = {}): Promise<T> => {
    const headers: Record<string, string> = { Authorization: `Bearer ${await token()}` };
    if (typeof init.body === "string") {
      headers["Content-Type"] = "application/json";
    }
    const response = await fetch(`${graphRoot}${path}`, { ...init, headers });
    if (!response.ok) {
      throw new Error(await graphError(response));
    }
    return response.status === 204 ? (undefined as T) : ((await response.json()) as T);
  };

  const drive = (containerId: string) => `/drives/${encodeURIComponent(containerId)}`;

  const list = async (containerId: string): Promise<CaseFile[]> => {
    const page = await graph<{ value: DriveItem[] }>(
      `${drive(containerId)}/root/children?$select=id,name,size,lastModifiedDateTime,webUrl,file&$top=200`
    );
    return page.value
      .filter((item) => item.file)
      .map((item) => ({
        id: item.id,
        name: item.name,
        size: item.size,
        modifiedOn: item.lastModifiedDateTime ? formatNorwegianDate(item.lastModifiedDateTime) : undefined,
        webUrl: item.webUrl,
      }));
  };

  return {
    list,

    async preview(containerId, itemId) {
      const result = await graph<{ getUrl?: string }>(`${drive(containerId)}/items/${encodeURIComponent(itemId)}/preview`, {
        method: "POST",
        body: "{}",
      });
      if (!result.getUrl) {
        throw new Error("SharePoint did not return a preview for this file.");
      }
      return `${result.getUrl}${result.getUrl.includes("?") ? "&" : "?"}nb=true`;
    },

    async download(containerId, itemId) {
      const item = await graph<DriveItem>(`${drive(containerId)}/items/${encodeURIComponent(itemId)}`);
      const url = item["@microsoft.graph.downloadUrl"];
      if (!url) {
        throw new Error("SharePoint did not return a download link for this file.");
      }
      return url;
    },

    async upload(containerId, file) {
      const path = `${drive(containerId)}/root:/${encodeURIComponent(file.name)}:`;
      if (file.size <= simpleUploadLimit) {
        await graph<DriveItem>(`${path}/content?@microsoft.graph.conflictBehavior=rename`, { method: "PUT", body: file });
        return;
      }

      const uploadSession = await graph<{ uploadUrl: string }>(`${path}/createUploadSession`, {
        method: "POST",
        body: JSON.stringify({ item: { "@microsoft.graph.conflictBehavior": "rename" } }),
      });
      for (let start = 0; start < file.size; start += uploadChunkSize) {
        const end = Math.min(start + uploadChunkSize, file.size);
        // The upload URL is pre-authenticated and rejects an Authorization header.
        const response = await fetch(uploadSession.uploadUrl, {
          method: "PUT",
          headers: { "Content-Range": `bytes ${start}-${end - 1}/${file.size}` },
          body: file.slice(start, end),
        });
        if (!response.ok) {
          throw new Error(await graphError(response));
        }
      }
    },

    async remove(containerId, itemId) {
      await graph<undefined>(`${drive(containerId)}/items/${encodeURIComponent(itemId)}`, { method: "DELETE" });
    },

    async listShares(containerId, itemId) {
      const page = await graph<{ value: ItemPermission[] }>(`${drive(containerId)}/items/${encodeURIComponent(itemId)}/permissions`);
      return page.value.reduce<FileShare[]>((all, permission) => all.concat(toFileShares(permission)), []);
    },

    async share(containerId, itemId, email, role) {
      await graph<unknown>(`${drive(containerId)}/items/${encodeURIComponent(itemId)}/invite`, {
        method: "POST",
        // SharePoint Embedded rejects invites that ask Microsoft to send the sharing email.
        body: JSON.stringify({ recipients: [{ email }], roles: [role], requireSignIn: true, sendInvitation: false }),
      });
    },

    async unshare(containerId, itemId, permissionId) {
      await graph<undefined>(
        `${drive(containerId)}/items/${encodeURIComponent(itemId)}/permissions/${encodeURIComponent(permissionId)}`,
        { method: "DELETE" }
      );
    },

    async search(containerId, query) {
      const { settings } = await start();
      const [files, result] = await Promise.all([
        list(containerId),
        graph<{ retrievalHits?: RetrievalHit[] }>("/copilot/retrieval", {
          method: "POST",
          body: JSON.stringify({
            queryString: query,
            dataSource: "sharePointEmbedded",
            dataSourceConfiguration: { sharePointEmbedded: { containerTypeId: settings.containerTypeId } },
            maximumNumberOfResults: 25,
          }),
        }),
      ]);
      // Retrieval searches every container of the type, so keep only hits on this case's files.
      return (result.retrievalHits ?? []).reduce<SearchHit[]>((hits, hit) => {
        const file = files.find((candidate) => candidate.webUrl && hit.webUrl && sameUrl(candidate.webUrl, hit.webUrl));
        if (file) {
          const text = (hit.extracts ?? []).map((extract) => extract.text ?? "").filter(Boolean).join("\n\n");
          hits.push({ itemId: file.id, name: file.name, text });
        }
        return hits;
      }, []);
    },

    async signIn() {
      // The popup must open in the click handler, so reuse the started session when there is one.
      const { app, settings } = session ?? (await start());
      await app.acquireTokenPopup({ scopes, loginHint: settings.loginHint });
    },
  };
}
