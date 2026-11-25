import * as React from "react";
import { CaseAccess, initials } from "./CaseAccessPanel";
import { CaseText, useCaseText } from "./case-documents-text";
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

function escapeHtml(value: string): string {
  return value.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

function linkEmailHtml(t: CaseText, file: CaseFile, to: string): string {
  const name = escapeHtml(file.name);
  return (
    `<p>${t.emailGreeting}</p><p>${t.emailShared(`<strong>${name}</strong>`)}</p>` +
    `<p><a href="${escapeHtml(file.webUrl ?? "")}">${t.emailOpen(name)}</a></p>` +
    `<p>${t.emailSignIn(escapeHtml(to))}</p>`
  );
}

const wait = (seconds: number) => new Promise<void>((resolve) => window.setTimeout(resolve, seconds * 1000));

function FileRoleSelect({ value, disabled, onChange }: { value: FileShareRole; disabled: boolean; onChange: (role: FileShareRole) => void }) {
  const t = useCaseText();
  return (
    <select aria-label={t.access} value={value} disabled={disabled} onChange={(event) => onChange(event.target.value as FileShareRole)}>
      <option value="read">{t.roles.reader}</option>
      <option value="write">{t.roles.writer}</option>
    </select>
  );
}

export const FileSharePanel: React.FC<FileSharePanelProps> = ({ caseId, containerId, file, files, access, onClose }) => {
  const t = useCaseText();
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
      setError(errorMessage(caught, t.sharingFailed));
    } finally {
      setProgress(null);
    }
  };

  const busy = progress !== null;

  React.useEffect(() => {
    void run(t.loadingShares, () => Promise.resolve());
  }, [file.id]);

  const chooseContact = async () => {
    try {
      const picked = await access.pickContact();
      if (!picked) {
        return;
      }
      if (!picked.email) {
        setError(t.noEmailInCrm(picked.name));
        return;
      }
      setEmail(picked.email);
    } catch {
      setError(t.contactLookupFailed);
    }
  };

  const copyLink = async (key: string) => {
    try {
      await navigator.clipboard.writeText(file.webUrl ?? "");
      setCopiedFor(key);
      window.setTimeout(() => setCopiedFor(null), 2000);
    } catch {
      setError(t.copyFailed);
    }
  };

  const shareWith = (address: string) =>
    run(t.sharingFile(file.name), async () => {
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
          setProgress(t.settingUpGuest(address));
          await wait(5 * attempt);
        }
      }
      setEmail("");
      setProgress(t.emailingLink(address));
      try {
        if (!file.webUrl) {
          throw new Error(t.copyFailed);
        }
        await files.sendEmail(address, t.emailSubject(file.name), linkEmailHtml(t, file, address));
        setStatus(t.sharedAndEmailed(address));
      } catch (caught) {
        setError(t.emailFailed(address, errorMessage(caught, t.unknownError)));
      }
    });

  const address = email.trim();

  return (
    <div className="spe-file-share">
      <div className="spe-row">
        <span className="spe-grant-label">{t.shareOnlyThisFile}</span>
        <button type="button" className="spe-close" aria-label={t.closeSharing} onClick={onClose}>
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
                <span>{share.name ? share.email : t.invitationSent}</span>
              </div>
              {file.webUrl && (
                <button type="button" className="spe-ghost" onClick={() => void copyLink(share.key)}>
                  {copiedFor === share.key ? t.copied : t.copyLink}
                </button>
              )}
              <FileRoleSelect
                value={share.role}
                disabled={busy}
                onChange={(next) => void run(t.updatingFileAccess, () => files.updateShare(containerId, file.id, share.permissionId, next))}
              />
              <button
                type="button"
                className="spe-close"
                aria-label={t.stopSharing(share.name || share.email)}
                disabled={busy}
                onClick={() => void run(t.removingFileAccess, () => files.unshare(containerId, file.id, share.permissionId))}
              >
                ✕
              </button>
            </li>
          ))}
        </ul>
      )}
      {shares?.length === 0 && <p className="spe-hint">{t.fileNotShared}</p>}

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
          aria-label={t.emailAddress}
          placeholder={t.emailPlaceholder}
          value={email}
          disabled={busy}
          onChange={(event) => setEmail(event.target.value)}
        />
        <button type="button" className="spe-contact-pick" disabled={busy} onClick={() => void chooseContact()}>
          {t.crmContact}
        </button>
        <FileRoleSelect value={role} disabled={busy} onChange={setRole} />
        <button type="submit" disabled={busy || !validEmail.test(address)}>
          {t.shareFile}
        </button>
      </form>
      <span className="spe-hint">{t.fileShareHint}</span>
    </div>
  );
};
