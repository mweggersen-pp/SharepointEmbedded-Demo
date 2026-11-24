import * as React from "react";
import { CaseAccess, CaseAccessPanel } from "./CaseAccessPanel";
import { CaseFile, CaseFiles, errorMessage, isSignInRequired, SearchHit } from "./CaseDocumentsGraph";

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

interface CopilotTurn {
  id: number;
  question: string;
  hits?: SearchHit[];
  error?: string;
}

interface OpenPreview {
  name: string;
  url: string;
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
  const [preview, setPreview] = React.useState<OpenPreview | null>(null);
  const [turns, setTurns] = React.useState<CopilotTurn[]>([]);
  const [question, setQuestion] = React.useState("");
  const [needsSignIn, setNeedsSignIn] = React.useState(false);
  const [status, setStatus] = React.useState<string | null>(null);
  const [error, setError] = React.useState<string | null>(null);
  const [busy, setBusy] = React.useState(false);
  const transcriptEnd = React.useRef<HTMLDivElement>(null);

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
        setError(errorMessage(caught, "The document request failed."));
      }
    } finally {
      setBusy(false);
    }
  };

  const containerId = container?.containerId ?? "";
  const canEdit = container?.role === "writer";
  const asking = turns.some((turn) => !turn.hits && !turn.error);

  React.useEffect(() => {
    setContainer(null);
    setFiles(null);
    setTurns([]);
    setPreview(null);
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

  React.useEffect(() => {
    transcriptEnd.current?.scrollIntoView({ block: "nearest" });
  }, [turns]);

  const openPreview = (itemId: string, name: string) =>
    run(async () => setPreview({ name, url: await caseFiles.preview(containerId, itemId) }));

  const ask = () => {
    const text = question.trim();
    if (!text || asking) {
      return;
    }
    const id = Date.now();
    setQuestion("");
    setTurns((current) => [...current, { id, question: text }]);
    caseFiles.search(containerId, text).then(
      (hits) => setTurns((current) => current.map((turn) => (turn.id === id ? { ...turn, hits } : turn))),
      (caught: unknown) => {
        if (isSignInRequired(caught)) {
          setNeedsSignIn(true);
        }
        const message = errorMessage(caught, "Copilot could not answer.");
        setTurns((current) => current.map((turn) => (turn.id === id ? { ...turn, error: message } : turn)));
      }
    );
  };

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
        <div className="spe-layout">
          <div className="spe-main">
            <section className="spe-panel">
              <div className="spe-row">
                <h3>Files</h3>
                <div className="spe-actions">
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
                  <button type="button" className="secondary" onClick={() => void load()} disabled={busy}>
                    Refresh
                  </button>
                </div>
              </div>
              {status && <p className="spe-message">{status}</p>}

              <ul className="spe-files">
                {(files ?? []).map((file) => (
                  <li key={file.id} className={preview?.name === file.name ? "selected" : undefined}>
                    <div className="spe-row">
                      <div className="spe-file-name">
                        <strong>{file.name}</strong>
                        <div>{[formatSize(file.size), file.modifiedOn].filter(Boolean).join(" · ")}</div>
                      </div>
                      <div className="spe-actions">
                        <button type="button" className="secondary" disabled={busy} onClick={() => void openPreview(file.id, file.name)}>
                          Preview
                        </button>
                        <button
                          type="button"
                          className="secondary"
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
            </section>

            {preview && (
              <section className="spe-panel">
                <div className="spe-row">
                  <h3 className="spe-file-name">{preview.name}</h3>
                  <div className="spe-actions">
                    <a className="spe-link" href={preview.url} target="_blank" rel="noopener noreferrer">
                      Open in new tab
                    </a>
                    <button type="button" className="secondary" onClick={() => setPreview(null)}>
                      Close preview
                    </button>
                  </div>
                </div>
                <iframe className="spe-preview" title={`Preview of ${preview.name}`} src={preview.url} />
              </section>
            )}

            <section className="spe-panel">
              <CaseAccessPanel caseId={caseId} access={access} canEdit={canEdit} />
            </section>
          </div>

          <section className="spe-panel spe-copilot">
            <h3>Ask Copilot</h3>
            <div className="spe-transcript">
              {turns.length === 0 && (
                <p className="spe-hint">Ask a question about the documents on this case. Copilot answers with the passages that match.</p>
              )}
              {turns.map((turn) => (
                <div key={turn.id} className="spe-turn">
                  <p className="spe-bubble spe-question">{turn.question}</p>
                  <div className="spe-bubble spe-answer">
                    {!turn.hits && !turn.error && <p>Searching the case documents…</p>}
                    {turn.error && <p className="spe-error-text">{turn.error}</p>}
                    {turn.hits?.length === 0 && <p>I found nothing about that in this case's documents.</p>}
                    {turn.hits?.map((hit, index) => (
                      <div key={`${hit.itemId}-${index}`} className="spe-passage">
                        <p className="spe-extract">{hit.text || "This file matches, but has no text passage."}</p>
                        <button type="button" className="spe-source" onClick={() => void openPreview(hit.itemId, hit.name)}>
                          {hit.name}
                        </button>
                      </div>
                    ))}
                  </div>
                </div>
              ))}
              <div ref={transcriptEnd} />
            </div>
            <form
              className="spe-ask"
              onSubmit={(event) => {
                event.preventDefault();
                ask();
              }}
            >
              <textarea
                aria-label="Question for Copilot"
                value={question}
                rows={2}
                placeholder="Ask about this case's documents"
                onChange={(event) => setQuestion(event.target.value)}
                onKeyDown={(event) => {
                  if (event.key === "Enter" && !event.shiftKey) {
                    event.preventDefault();
                    ask();
                  }
                }}
              />
              <button type="submit" disabled={asking || !question.trim()}>
                Ask
              </button>
            </form>
          </section>
        </div>
      )}
    </div>
  );
};
