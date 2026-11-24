import * as React from "react";
import { CaseAccess, CaseAccessPanel } from "./CaseAccessPanel";
import { CaseFile, CaseFiles, isSignInRequired, SearchHit } from "./CaseDocumentsGraph";

export interface CaseContainer {
  containerId: string;
  role: string;
}

export interface SharePointEmbeddedProps {
  caseId: string;
  files: CaseFiles;
  access: CaseAccess;
  ensureContainer: (caseId: string) => Promise<CaseContainer>;
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

export const SharePointEmbedded: React.FC<SharePointEmbeddedProps> = ({ caseId, files: caseFiles, access, ensureContainer }) => {
  const [container, setContainer] = React.useState<CaseContainer | null>(null);
  const [files, setFiles] = React.useState<CaseFile[] | null>(null);
  const [hits, setHits] = React.useState<SearchHit[]>([]);
  const [query, setQuery] = React.useState("");
  const [previewUrl, setPreviewUrl] = React.useState<string | null>(null);
  const [needsSignIn, setNeedsSignIn] = React.useState(false);
  const [status, setStatus] = React.useState<string | null>(null);
  const [error, setError] = React.useState<string | null>(null);
  const [busy, setBusy] = React.useState(false);

  const run = async (work: () => Promise<void>) => {
    setBusy(true);
    setError(null);
    setStatus(null);
    try {
      await work();
    } catch (caught) {
      if (isSignInRequired(caught)) {
        setNeedsSignIn(true);
      } else {
        setError(caught instanceof Error ? caught.message : "The document request failed.");
      }
    } finally {
      setBusy(false);
    }
  };

  const containerId = container?.containerId ?? "";
  const canEdit = container?.role === "writer";

  React.useEffect(() => {
    setContainer(null);
    setFiles(null);
    setHits([]);
    setPreviewUrl(null);
    if (caseId) {
      void run(async () => setContainer(await ensureContainer(caseId)));
    }
  }, [caseId]);

  const load = () =>
    run(async () => {
      setNeedsSignIn(false);
      setFiles(await caseFiles.list(containerId));
    });

  React.useEffect(() => {
    if (containerId) {
      void load();
    }
  }, [containerId]);

  if (!caseId) {
    return (
      <div className="spe-docs">
        <h2>Case documents</h2>
        <p className="spe-message">Save the case first. A document container is created for each saved case.</p>
      </div>
    );
  }

  return (
    <div className="spe-docs">
      <h2>Case documents</h2>
      {!container && !error && <p className="spe-message">Opening the document container for this case.</p>}
      {error && <p className="spe-message spe-error">{error}</p>}
      {status && <p className="spe-message">{status}</p>}
      {needsSignIn && (
        <div className="spe-message spe-row">
          <span>Sign in to SharePoint to work with the documents on this case.</span>
          <button
            type="button"
            onClick={() =>
              void run(async () => {
                await caseFiles.signIn();
                await load();
              })
            }
          >
            Sign in
          </button>
        </div>
      )}

      {containerId && !needsSignIn && (
        <>
          <form
            className="spe-search"
            onSubmit={(event) => {
              event.preventDefault();
              void run(async () => {
                const found = await caseFiles.search(containerId, query.trim());
                setHits(found);
                setPreviewUrl(null);
                if (found.length === 0) {
                  setStatus("Copilot found nothing in this case's documents.");
                }
              });
            }}
          >
            <input
              aria-label="Search documents with Copilot"
              value={query}
              placeholder="Ask Copilot about the documents on this case"
              onChange={(event) => setQuery(event.target.value)}
              disabled={busy}
            />
            <button type="submit" disabled={busy || !query.trim()}>
              Search with Copilot
            </button>
            <button type="button" className="secondary" onClick={() => void load()} disabled={busy}>
              Refresh
            </button>
            {canEdit && (
              <label className={`spe-upload${busy ? " disabled" : ""}`}>
                Upload
                <input
                  type="file"
                  multiple
                  disabled={busy}
                  onChange={(event) => {
                    const chosen = Array.from(event.target.files ?? []);
                    event.target.value = "";
                    void run(async () => {
                      for (const file of chosen) {
                        setStatus(`Uploading ${file.name}.`);
                        await caseFiles.upload(containerId, file);
                      }
                      setFiles(await caseFiles.list(containerId));
                      setStatus(`Uploaded ${chosen.length} ${chosen.length === 1 ? "file" : "files"}.`);
                    });
                  }}
                />
              </label>
            )}
          </form>

          {hits.length > 0 && (
            <ul className="spe-hits">
              {hits.map((hit, index) => (
                <li key={`${hit.itemId}-${index}`}>
                  <div className="spe-row">
                    <strong>{hit.name}</strong>
                    <button
                      type="button"
                      className="secondary"
                      disabled={busy}
                      onClick={() => void run(async () => setPreviewUrl(await caseFiles.preview(containerId, hit.itemId)))}
                    >
                      Preview
                    </button>
                  </div>
                  <p className="spe-extract">{hit.text}</p>
                </li>
              ))}
            </ul>
          )}

          <ul className="spe-files">
            {(files ?? []).map((file) => (
              <li key={file.id}>
                <div className="spe-row">
                  <div>
                    <strong>{file.name}</strong>
                    <div>{[formatSize(file.size), file.modifiedOn].filter(Boolean).join(" · ")}</div>
                  </div>
                  <div className="spe-actions">
                    <button
                      type="button"
                      className="secondary"
                      disabled={busy}
                      onClick={() => void run(async () => setPreviewUrl(await caseFiles.preview(containerId, file.id)))}
                    >
                      Preview
                    </button>
                    <button
                      type="button"
                      disabled={busy}
                      onClick={() =>
                        void run(async () => {
                          // A pop-up opened after an await is blocked; the download URL is served as an attachment.
                          const link = document.createElement("a");
                          link.href = await caseFiles.download(containerId, file.id);
                          link.click();
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
          {files === null && busy && <p className="spe-message">Loading documents.</p>}
          {files?.length === 0 && <p className="spe-message">No documents on this case yet.</p>}
          {previewUrl && <iframe className="spe-preview" title="Document preview" src={previewUrl} />}
        </>
      )}

      {container && <CaseAccessPanel caseId={caseId} access={access} canEdit={canEdit} />}
    </div>
  );
};
