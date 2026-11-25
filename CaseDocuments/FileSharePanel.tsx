import * as React from "react";
import { CaseAccess, initials } from "./CaseAccessPanel";
import { CaseFile, CaseFiles, errorMessage, FileShare, FileShareRole } from "./CaseDocumentsGraph";
import { ProgressSpinner } from "./ProgressSpinner";

interface FileSharePanelProps {
  caseId: string;
  containerId: string;
  file: CaseFile;
  files: CaseFiles;
  access: CaseAccess;
  onClose: () => void;
}

const validEmail = /^[^@\s]+@[^@\s]+\.[^@\s]+$/;

const wait = (seconds: number) => new Promise<void>((resolve) => window.setTimeout(resolve, seconds * 1000));

function FileRoleSelect({ value, disabled, onChange }: { value: FileShareRole; disabled: boolean; onChange: (role: FileShareRole) => void }) {
  return (
    <select aria-label="Access" value={value} disabled={disabled} onChange={(event) => onChange(event.target.value as FileShareRole)}>
      <option value="read">Read</option>
      <option value="write">Edit</option>
    </select>
  );
}

export const FileSharePanel: React.FC<FileSharePanelProps> = ({ caseId, containerId, file, files, access, onClose }) => {
  const [shares, setShares] = React.useState<FileShare[] | null>(null);
  const [email, setEmail] = React.useState("");
  const [role, setRole] = React.useState<FileShareRole>("read");
  const [error, setError] = React.useState<string | null>(null);
  const [status, setStatus] = React.useState<string | null>(null);
  const [copiedFor, setCopiedFor] = React.useState<string | null>(null);
  const [progress, setProgress] = React.useState<string | null>(null);

  const run = async (label: string, work: () => Promise<void>) => {
    setProgress(label);
    setError(null);
    setStatus(null);
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

  const copyLink = async (key: string) => {
    try {
      await navigator.clipboard.writeText(file.webUrl ?? "");
      setCopiedFor(key);
      window.setTimeout(() => setCopiedFor(null), 2000);
    } catch {
      setError("Could not copy the link.");
    }
  };

  const shareWith = (address: string) =>
    run(`Sharing ${file.name}`, async () => {
      const created = await access.ensureGuest(caseId, address);
      // SharePoint sees a brand-new guest account only after Entra replicates it.
      for (let attempt = 1; ; attempt++) {
        try {
          await files.share(containerId, file.id, address, role);
          break;
        } catch (caught) {
          if (!created || attempt === 5) {
            throw caught;
          }
          setProgress(`Setting up a guest account for ${address}`);
          await wait(5 * attempt);
        }
      }
      setEmail("");
      setProgress(`Emailing the link to ${address}`);
      try {
        await files.emailLink(address, file);
        setStatus(`${address} can now open this file and got an email with the link.`);
      } catch (caught) {
        setError(`${address} can open the file, but the email was not sent: ${errorMessage(caught, "unknown error")} Use Copy link instead.`);
      }
    });

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
      {status && <p className="spe-banner spe-banner-ok">{status}</p>}
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
              {file.webUrl && (
                <button type="button" className="spe-ghost" onClick={() => void copyLink(share.key)}>
                  {copiedFor === share.key ? "Copied ✓" : "Copy link"}
                </button>
              )}
              <FileRoleSelect
                value={share.role}
                disabled={busy}
                onChange={(next) => void run("Updating access to the file", () => files.updateShare(containerId, file.id, share.permissionId, next))}
              />
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
          if (validEmail.test(address)) {
            void shareWith(address);
          }
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
        <FileRoleSelect value={role} disabled={busy} onChange={setRole} />
        <button type="submit" disabled={busy || !validEmail.test(address)}>
          Share file
        </button>
      </form>
      <span className="spe-hint">They get an email from you with a link to this file only, and sign in as that email address to open it.</span>
    </div>
  );
};
