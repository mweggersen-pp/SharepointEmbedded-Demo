import * as React from "react";
import { CaseAccess, initials } from "./CaseAccessPanel";
import { CaseFile, CaseFiles, errorMessage, FileShare, FileShareRole } from "./CaseDocumentsGraph";
import { ProgressSpinner } from "./ProgressSpinner";

interface FileSharePanelProps {
  containerId: string;
  file: CaseFile;
  files: CaseFiles;
  access: CaseAccess;
  onClose: () => void;
}

const validEmail = /^[^@\s]+@[^@\s]+\.[^@\s]+$/;

function linkEmail(file: CaseFile, to: string): string {
  const body = `Hi,\n\nI have shared ${file.name} with you:\n${file.webUrl ?? ""}\n\nSign in as ${to} to open it.`;
  return `mailto:${encodeURIComponent(to)}?subject=${encodeURIComponent(file.name)}&body=${encodeURIComponent(body)}`;
}

export const FileSharePanel: React.FC<FileSharePanelProps> = ({ containerId, file, files, access, onClose }) => {
  const [shares, setShares] = React.useState<FileShare[] | null>(null);
  const [email, setEmail] = React.useState("");
  const [role, setRole] = React.useState<FileShareRole>("read");
  const [error, setError] = React.useState<string | null>(null);
  const [sharedWith, setSharedWith] = React.useState<string | null>(null);
  const [copied, setCopied] = React.useState(false);
  const [progress, setProgress] = React.useState<string | null>(null);

  const run = async (label: string, work: () => Promise<void>) => {
    setProgress(label);
    setError(null);
    setSharedWith(null);
    try {
      await work();
      setShares(await files.listShares(containerId, file.id));
    } catch (caught) {
      setError(errorMessage(caught, "The sharing request failed."));
    } finally {
      setProgress(null);
    }
  };

  const busy = progress !== null;

  React.useEffect(() => {
    void run("Loading who this file is shared with", () => Promise.resolve());
  }, [file.id]);

  const chooseContact = async () => {
    try {
      const picked = await access.pickContact();
      if (!picked) {
        return;
      }
      if (!picked.email) {
        setError(`${picked.name} has no email address in CRM.`);
        return;
      }
      setEmail(picked.email);
    } catch {
      setError("Could not open the contact lookup.");
    }
  };

  const copyLink = async () => {
    try {
      await navigator.clipboard.writeText(file.webUrl ?? "");
      setCopied(true);
    } catch {
      setError("Could not copy the link.");
    }
  };

  const address = email.trim();

  return (
    <div className="spe-file-share">
      <div className="spe-row">
        <span className="spe-grant-label">Share only this file</span>
        <button type="button" className="spe-close" aria-label="Close sharing" onClick={onClose}>
          ✕
        </button>
      </div>
      {error && <p className="spe-banner spe-banner-error">{error}</p>}
      {sharedWith && (
        <div className="spe-banner spe-banner-ok spe-row">
          <span>{sharedWith} can now open this file. Send them the link:</span>
          <div className="spe-actions">
            <a className="spe-link-button" href={linkEmail(file, sharedWith)}>
              ✉ Email link
            </a>
            {file.webUrl && (
              <button
                type="button"
                className="spe-ghost"
                onClick={() => void copyLink()}
              >
                {copied ? "Copied" : "Copy link"}
              </button>
            )}
          </div>
        </div>
      )}
      {progress && <ProgressSpinner label={progress} />}

      {shares && shares.length > 0 && (
        <ul className="spe-people">
          {shares.map((share) => (
            <li key={share.key}>
              <span className="spe-avatar external small" aria-hidden="true">
                {initials(share.name || share.email)}
              </span>
              <div className="spe-file-name">
                <strong>{share.name || share.email}</strong>
                <span>{share.name ? share.email : "Invitation sent"}</span>
              </div>
              <span className={`spe-role spe-role-${share.role === "write" ? "writer" : "reader"}`}>
                {share.role === "write" ? "Edit" : "Read"}
              </span>
              {share.email && (
                <a className="spe-link-button" href={linkEmail(file, share.email)} title={`Email the link to ${share.email}`}>
                  ✉
                </a>
              )}
              <button
                type="button"
                className="spe-close"
                aria-label={`Stop sharing with ${share.name || share.email}`}
                disabled={busy}
                onClick={() => void run("Removing access to the file", () => files.unshare(containerId, file.id, share.permissionId))}
              >
                ✕
              </button>
            </li>
          ))}
        </ul>
      )}
      {shares?.length === 0 && <p className="spe-hint">This file is not shared with anyone outside the case yet.</p>}

      <form
        className="spe-grant-row"
        onSubmit={(event) => {
          event.preventDefault();
          if (!validEmail.test(address)) {
            return;
          }
          void run(`Sharing ${file.name}`, async () => {
            await files.share(containerId, file.id, address, role);
            setEmail("");
            setCopied(false);
            setSharedWith(address);
          });
        }}
      >
        <input
          type="email"
          className="spe-email"
          aria-label="Email address"
          placeholder="name@company.com"
          value={email}
          disabled={busy}
          onChange={(event) => setEmail(event.target.value)}
        />
        <button type="button" className="spe-contact-pick" disabled={busy} onClick={() => void chooseContact()}>
          ＋ CRM contact
        </button>
        <select aria-label="Access" value={role} disabled={busy} onChange={(event) => setRole(event.target.value as FileShareRole)}>
          <option value="read">Read</option>
          <option value="write">Edit</option>
        </select>
        <button type="submit" disabled={busy || !validEmail.test(address)}>
          Share file
        </button>
      </form>
      <span className="spe-hint">They can open this file only, signed in as that email address. Microsoft does not email them, so send the link yourself.</span>
    </div>
  );
};
