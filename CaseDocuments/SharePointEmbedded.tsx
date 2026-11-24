import * as React from "react";

export interface CaseFile {
  id: string;
  name: string;
  size?: number;
  modifiedOn?: string;
}

export interface SearchHit {
  itemId?: string;
  name?: string;
  webUrl?: string;
  text: string;
}

export interface DocumentClient {
  ensure(caseId: string): Promise<string>;
  list(containerId: string): Promise<CaseFile[]>;
  preview(containerId: string, itemId: string): Promise<string>;
  download(containerId: string, itemId: string): Promise<string>;
  search(containerId: string, query: string): Promise<SearchHit[]>;
}

export interface SharePointEmbeddedProps {
  containerId: string;
  caseId: string;
  client: DocumentClient;
  onContainerId: (containerId: string) => void;
}

function formatSize(size?: number): string {
  if (size === undefined) {
    return "";
  }
  if (size < 1024) {
    return `${size} B`;
  }
  if (size < 1024 * 1024) {
    return `${Math.round(size / 1024)} KB`;
  }
  return `${(size / (1024 * 1024)).toFixed(1)} MB`;
}

export const SharePointEmbedded: React.FC<SharePointEmbeddedProps> = ({
  containerId,
  caseId,
  client,
  onContainerId,
}) => {
  const [files, setFiles] = React.useState<CaseFile[]>([]);
  const [hits, setHits] = React.useState<SearchHit[]>([]);
  const [query, setQuery] = React.useState("");
  const [previewUrl, setPreviewUrl] = React.useState<string | null>(null);
  const [status, setStatus] = React.useState<string | null>(null);
  const [error, setError] = React.useState<string | null>(null);
  const [busy, setBusy] = React.useState(false);

  const load = React.useCallback(async () => {
    if (!containerId) {
      setFiles([]);
      return;
    }
    setBusy(true);
    setError(null);
    try {
      setFiles(await client.list(containerId));
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : "Could not load documents.");
    } finally {
      setBusy(false);
    }
  }, [client, containerId]);

  React.useEffect(() => {
    void load();
  }, [load]);

  const run = async (work: () => Promise<void>) => {
    setBusy(true);
    setError(null);
    setStatus(null);
    try {
      await work();
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : "The document request failed.");
    } finally {
      setBusy(false);
    }
  };

  React.useEffect(() => {
    if (containerId || !caseId) {
      return;
    }
    void run(async () => {
      onContainerId(await client.ensure(caseId));
    });
  }, [caseId, containerId]);

  return (
    <div className="spe-docs">
      <h2>Case documents</h2>
      {!containerId && !error && (
        <p className="spe-message">
          {caseId
            ? "Creating a SharePoint Embedded container for this case."
            : "Save the case first. A container is created for each saved case."}
        </p>
      )}
      {error && <p className="spe-message spe-error">{error}</p>}
      {status && <p className="spe-message">{status}</p>}

      <form
        className="spe-search"
        onSubmit={(event) => {
          event.preventDefault();
          if (!containerId || !query.trim()) {
            return;
          }
          void run(async () => {
            setHits(await client.search(containerId, query.trim()));
            setPreviewUrl(null);
          });
        }}
      >
        <input
          aria-label="Search documents with Copilot"
          value={query}
          placeholder="Ask Copilot about the documents on this case"
          onChange={(event) => setQuery(event.target.value)}
          disabled={!containerId || busy}
        />
        <button type="submit" disabled={!containerId || busy || !query.trim()}>
          Search with Copilot
        </button>
        <button
          type="button"
          className="secondary"
          onClick={() => {
            if (!containerId) {
              void run(async () => {
                if (!caseId) {
                  throw new Error("Save the case first. A container is created for each saved case.");
                }
                onContainerId(await client.ensure(caseId));
              });
              return;
            }
            void load();
          }}
          disabled={busy || (!containerId && !caseId)}
        >
          Refresh
        </button>
      </form>

      {hits.length > 0 && (
        <ul className="spe-hits">
          {hits.map((hit, index) => (
            <li key={`${hit.webUrl ?? hit.itemId ?? "hit"}-${index}`}>
              <div className="spe-row">
                <strong>{hit.name ?? hit.webUrl ?? "Matching passage"}</strong>
                {hit.itemId && (
                  <button
                    type="button"
                    className="secondary"
                    disabled={busy}
                    onClick={() =>
                      void run(async () => {
                        setPreviewUrl(await client.preview(containerId, hit.itemId!));
                      })
                    }
                  >
                    Preview
                  </button>
                )}
              </div>
              <p className="spe-extract">{hit.text}</p>
            </li>
          ))}
        </ul>
      )}

      <ul className="spe-files">
        {files.map((file) => (
          <li key={file.id}>
            <div className="spe-row">
              <div>
                <strong>{file.name}</strong>
                <div>
                  {[formatSize(file.size), file.modifiedOn].filter(Boolean).join(" · ")}
                </div>
              </div>
              <div className="spe-actions">
                <button
                  type="button"
                  className="secondary"
                  disabled={busy}
                  onClick={() =>
                    void run(async () => {
                      setPreviewUrl(await client.preview(containerId, file.id));
                    })
                  }
                >
                  Preview
                </button>
                <button
                  type="button"
                  disabled={busy}
                  onClick={() =>
                    void run(async () => {
                      const url = await client.download(containerId, file.id);
                      window.open(url, "_blank", "noopener");
                      setStatus(`Download started for ${file.name}.`);
                    })
                  }
                >
                  Download
                </button>
              </div>
            </div>
          </li>
        ))}
      </ul>
      {containerId && !busy && files.length === 0 && !error && (
        <p className="spe-message">No documents in this case container.</p>
      )}
      {previewUrl && (
        <iframe className="spe-preview" title="Document preview" src={previewUrl} />
      )}
    </div>
  );
};
