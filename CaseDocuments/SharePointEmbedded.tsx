import * as React from "react";
import { CaseAccess, CaseAccessPanel } from "./CaseAccessPanel";
import { CaseFile, CaseFiles, errorMessage, isSignInRequired, SearchHit } from "./CaseDocumentsGraph";
import { FileSharePanel } from "./FileSharePanel";
import { ProgressSpinner } from "./ProgressSpinner";

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
  itemId: string;
  name: string;
  url: string;
}

const suggestedQuestions = ["Summarise the documents on this case", "Which deadlines are mentioned?", "Who are the parties involved?"];

const fileKinds: Record<string, string> = {
  doc: "word",
  docx: "word",
  xls: "excel",
  xlsx: "excel",
  csv: "excel",
  ppt: "powerpoint",
  pptx: "powerpoint",
  pdf: "pdf",
  png: "image",
  jpg: "image",
  jpeg: "image",
  gif: "image",
  msg: "mail",
  eml: "mail",
};

function fileBadge(name: string): { label: string; kind: string } {
  const extension = name.includes(".") ? name.split(".").pop()?.toLowerCase() ?? "" : "";
  return { label: extension.slice(0, 4).toUpperCase() || "FILE", kind: fileKinds[extension] ?? "other" };
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
  const [progress, setProgress] = React.useState<string | null>(null);
  const [status, setStatus] = React.useState<string | null>(null);
  const [error, setError] = React.useState<string | null>(null);
  const [dragging, setDragging] = React.useState(false);
  const [confirmDelete, setConfirmDelete] = React.useState<string | null>(null);
  const [sharing, setSharing] = React.useState<string | null>(null);
  const transcriptEnd = React.useRef<HTMLDivElement>(null);
  const filePicker = React.useRef<HTMLInputElement>(null);
  const previewPanel = React.useRef<HTMLDivElement>(null);

  const run = async (label: string, work: () => Promise<void>) => {
    setProgress(label);
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
      setProgress(null);
    }
  };

  const busy = progress !== null;
  const containerId = container?.containerId ?? "";
  const canEdit = container?.role === "writer";
  const asking = turns.some((turn) => !turn.hits && !turn.error);

  React.useEffect(() => {
    setContainer(null);
    setFiles(null);
    setTurns([]);
    setPreview(null);
    if (caseId) {
      void run("Preparing the document container", async () => setContainer(await ensureContainer(caseId)));
    }
  }, [caseId]);

  const load = () =>
    run("Loading documents", async () => {
      setNeedsSignIn(false);
      setFiles(await caseFiles.list(containerId));
    });

  React.useEffect(() => {
    if (containerId) {
      void load();
    }
  }, [containerId]);

  React.useEffect(() => {
    previewPanel.current?.scrollIntoView({ block: "nearest", behavior: "smooth" });
  }, [preview?.itemId]);

  React.useEffect(() => {
    transcriptEnd.current?.scrollIntoView({ block: "nearest" });
  }, [turns]);

  const uploadFiles = (chosen: File[]) => {
    if (chosen.length === 0 || busy) {
      return;
    }
    void run(`Uploading ${chosen.length === 1 ? chosen[0].name : `${chosen.length} files`}`, async () => {
      for (let index = 0; index < chosen.length; index++) {
        const file = chosen[index];
        setProgress(chosen.length === 1 ? `Uploading ${file.name}` : `Uploading ${index + 1} of ${chosen.length}: ${file.name}`);
        await caseFiles.upload(containerId, file);
      }
      setFiles(await caseFiles.list(containerId));
      setStatus(`Uploaded ${chosen.length} ${chosen.length === 1 ? "file" : "files"}.`);
    });
  };

  const openPreview = (itemId: string, name: string) =>
    run(`Opening ${name}`, async () => setPreview({ itemId, name, url: await caseFiles.preview(containerId, itemId) }));

  const ask = async (text: string) => {
    const trimmed = text.trim();
    if (!trimmed || asking) {
      return;
    }
    const id = Date.now();
    setQuestion("");
    setTurns((current) => [...current, { id, question: trimmed }]);
    try {
      const hits = await caseFiles.search(containerId, trimmed);
      setTurns((current) => current.map((turn) => (turn.id === id ? { ...turn, hits } : turn)));
    } catch (caught) {
      if (isSignInRequired(caught)) {
        setNeedsSignIn(true);
      }
      const message = errorMessage(caught, "Copilot could not answer.");
      setTurns((current) => current.map((turn) => (turn.id === id ? { ...turn, error: message } : turn)));
    }
  };

  if (!caseId) {
    return (
      <div className="spe-docs">
        <div className="spe-empty-card">
          <span className="spe-empty-icon" aria-hidden="true">🗂️</span>
          <h3>Save the case to start collecting documents</h3>
          <p>Each saved case gets its own secure document container.</p>
        </div>
      </div>
    );
  }

  return (
    <div className="spe-docs">
      <header className="spe-hero">
        <div>
          <p className="spe-eyebrow">SharePoint Embedded</p>
          <h2>Case documents</h2>
        </div>
        <div className="spe-hero-chips">
          {files && (
            <span className="spe-chip">
              {files.length} {files.length === 1 ? "file" : "files"}
            </span>
          )}
          {container && <span className="spe-chip">{canEdit ? "You can edit" : "Read only"}</span>}
        </div>
      </header>

      {!container && !error && (
        <div className="spe-stage">
          <ProgressSpinner size="large" label="Preparing the document container" detail="New cases take a few seconds while SharePoint sets things up." />
        </div>
      )}
      {error && <p className="spe-banner spe-banner-error">{error}</p>}
      {needsSignIn && (
        <div className="spe-banner spe-row">
          <span>Sign in to SharePoint to work with the documents on this case.</span>
          <button
            type="button"
            onClick={() =>
              void run("Signing in", async () => {
                await caseFiles.signIn();
                setNeedsSignIn(false);
                setFiles(await caseFiles.list(containerId));
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
                <button type="button" className="spe-ghost" onClick={() => void load()} disabled={busy}>
                  ↻ Refresh
                </button>
              </div>

              {canEdit && (
                <div
                  className={`spe-dropzone${dragging ? " active" : ""}${busy ? " disabled" : ""}`}
                  role="button"
                  tabIndex={0}
                  onClick={() => !busy && filePicker.current?.click()}
                  onKeyDown={(event) => {
                    if ((event.key === "Enter" || event.key === " ") && !busy) {
                      event.preventDefault();
                      filePicker.current?.click();
                    }
                  }}
                  onDragOver={(event) => {
                    event.preventDefault();
                    setDragging(true);
                  }}
                  onDragLeave={() => setDragging(false)}
                  onDrop={(event) => {
                    event.preventDefault();
                    setDragging(false);
                    uploadFiles(Array.from(event.dataTransfer.files));
                  }}
                >
                  {progress?.startsWith("Uploading") ? (
                    <ProgressSpinner label={progress} />
                  ) : (
                    <>
                      <span className="spe-drop-icon" aria-hidden="true">⬆</span>
                      <span>
                        <strong>Drop files here</strong> or click to upload
                      </span>
                    </>
                  )}
                  <input
                    ref={filePicker}
                    type="file"
                    multiple
                    hidden
                    onChange={(event) => {
                      const chosen = Array.from(event.target.files ?? []);
                      event.target.value = "";
                      uploadFiles(chosen);
                    }}
                  />
                </div>
              )}

              {status && <p className="spe-banner spe-banner-ok">{status}</p>}
              {progress && !progress.startsWith("Uploading") && container && <ProgressSpinner label={progress} />}

              <ul className="spe-files">
                {(files ?? []).map((file) => {
                  const badge = fileBadge(file.name);
                  return (
                    <li key={file.id} className={preview?.itemId === file.id ? "selected" : undefined}>
                      <span className={`spe-badge spe-badge-${badge.kind}`}>{badge.label}</span>
                      <div className="spe-file-name">
                        <strong title={file.name}>{file.name}</strong>
                        <span>{[formatSize(file.size), file.modifiedOn].filter(Boolean).join(" · ")}</span>
                      </div>
                      {confirmDelete === file.id ? (
                        <div className="spe-actions spe-confirm">
                          <span>Delete for everyone?</span>
                          <button
                            type="button"
                            className="spe-danger"
                            disabled={busy}
                            onClick={() =>
                              void run(`Deleting ${file.name}`, async () => {
                                await caseFiles.remove(containerId, file.id);
                                setConfirmDelete(null);
                                if (preview?.itemId === file.id) {
                                  setPreview(null);
                                }
                                setFiles(await caseFiles.list(containerId));
                                setStatus(`${file.name} was moved to the container's recycle bin.`);
                              })
                            }
                          >
                            Delete
                          </button>
                          <button type="button" className="spe-ghost" disabled={busy} onClick={() => setConfirmDelete(null)}>
                            Cancel
                          </button>
                        </div>
                      ) : (
                      <div className="spe-actions">
                        <button type="button" className="spe-ghost" disabled={busy} onClick={() => void openPreview(file.id, file.name)}>
                          Preview
                        </button>
                        <button
                          type="button"
                          className="spe-ghost"
                          disabled={busy}
                          onClick={() =>
                            void run(`Preparing ${file.name}`, async () => {
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
                        {canEdit && (
                          <button
                            type="button"
                            className={`spe-ghost${sharing === file.id ? " on" : ""}`}
                            disabled={busy}
                            onClick={() => setSharing(sharing === file.id ? null : file.id)}
                          >
                            Share
                          </button>
                        )}
                        {canEdit && (
                          <button
                            type="button"
                            className="spe-close"
                            aria-label={`Delete ${file.name}`}
                            title="Delete"
                            disabled={busy}
                            onClick={() => setConfirmDelete(file.id)}
                          >
                            🗑
                          </button>
                        )}
                      </div>
                      )}
                      {preview?.itemId === file.id && (
                        <div className="spe-preview-panel" ref={previewPanel}>
                          <div className="spe-preview-bar">
                            <strong className="spe-file-name" title={preview.name}>
                              Preview · {preview.name}
                            </strong>
                            <div className="spe-actions">
                              <a className="spe-preview-link" href={preview.url} target="_blank" rel="noopener noreferrer">
                                Open in new tab ↗
                              </a>
                              <button type="button" className="spe-close" aria-label="Close preview" onClick={() => setPreview(null)}>
                                ✕
                              </button>
                            </div>
                          </div>
                          <iframe className="spe-preview" title={`Preview of ${preview.name}`} src={preview.url} />
                        </div>
                      )}
                      {sharing === file.id && (
                        <FileSharePanel containerId={containerId} file={file} files={caseFiles} access={access} onClose={() => setSharing(null)} />
                      )}
                    </li>
                  );
                })}
              </ul>
              {files?.length === 0 && (
                <div className="spe-empty-card">
                  <span className="spe-empty-icon" aria-hidden="true">📄</span>
                  <h3>No documents yet</h3>
                  <p>{canEdit ? "Drop the first file above to get started." : "Documents added to this case will show up here."}</p>
                </div>
              )}
            </section>

            <section className="spe-panel">
              <CaseAccessPanel caseId={caseId} access={access} canEdit={canEdit} />
            </section>
          </div>

          <section className="spe-panel spe-copilot">
            <div className="spe-copilot-head">
              <span className="spe-sparkle" aria-hidden="true">✦</span>
              <div>
                <h3>Ask Copilot</h3>
                <p>Answers come from this case&apos;s documents only.</p>
              </div>
            </div>
            <div className="spe-transcript">
              {turns.length === 0 && (
                <div className="spe-suggestions">
                  {suggestedQuestions.map((suggestion) => (
                    <button key={suggestion} type="button" className="spe-suggestion" onClick={() => void ask(suggestion)}>
                      {suggestion}
                    </button>
                  ))}
                </div>
              )}
              {turns.map((turn) => (
                <div key={turn.id} className="spe-turn">
                  <p className="spe-bubble spe-question">{turn.question}</p>
                  <div className="spe-bubble spe-answer">
                    {!turn.hits && !turn.error && (
                      <span className="spe-typing" aria-label="Copilot is searching">
                        <span />
                        <span />
                        <span />
                      </span>
                    )}
                    {turn.error && <p className="spe-error-text">{turn.error}</p>}
                    {turn.hits?.length === 0 && <p>I found nothing about that in this case&apos;s documents.</p>}
                    {turn.hits?.map((hit, index) => (
                      <div key={`${hit.itemId}-${index}`} className="spe-passage">
                        <p>{hit.text || "This file matches, but has no text passage."}</p>
                        <button type="button" className="spe-source" onClick={() => void openPreview(hit.itemId, hit.name)}>
                          📎 {hit.name}
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
                void ask(question);
              }}
            >
              <textarea
                aria-label="Question for Copilot"
                value={question}
                rows={2}
                placeholder="Ask about this case's documents…"
                onChange={(event) => setQuestion(event.target.value)}
                onKeyDown={(event) => {
                  if (event.key === "Enter" && !event.shiftKey) {
                    event.preventDefault();
                    void ask(question);
                  }
                }}
              />
              <button type="submit" className="spe-send" aria-label="Ask Copilot" disabled={asking || !question.trim()}>
                ➤
              </button>
            </form>
          </section>
        </div>
      )}
    </div>
  );
};
