import * as React from "react";
import { useCaseText } from "./case-documents-text";
import { errorMessage } from "./CaseDocumentsGraph";
import { ProgressSpinner } from "./ProgressSpinner";

export type ContactRole = "reader" | "writer";

export interface AccessEntry {
  id: string;
  name: string;
  email: string;
  role: string;
  external: boolean;
}

export interface ContactChoice {
  id: string;
  name: string;
  email: string;
}

export interface CaseAccess {
  list(caseId: string): Promise<AccessEntry[]>;
  pickContact(): Promise<ContactChoice | null>;
  grant(caseId: string, contactId: string, role: ContactRole): Promise<void>;
  shareByEmail(caseId: string, email: string, role: ContactRole): Promise<void>;
  ensureGuest(caseId: string, email: string): Promise<boolean>;
  update(caseId: string, permissionId: string, role: ContactRole): Promise<void>;
  remove(caseId: string, permissionId: string): Promise<void>;
}

interface CaseAccessPanelProps {
  caseId: string;
  access: CaseAccess;
  canEdit: boolean;
}

export function initials(name: string): string {
  const parts = name.trim().split(/\s+/).filter(Boolean);
  return ((parts[0]?.[0] ?? "?") + (parts.length > 1 ? parts[parts.length - 1][0] : "")).toUpperCase();
}

function RoleSelect({ value, disabled, onChange }: { value: string; disabled: boolean; onChange: (role: ContactRole) => void }) {
  const t = useCaseText();
  return (
    <select aria-label={t.access} value={value} disabled={disabled} onChange={(event) => onChange(event.target.value as ContactRole)}>
      <option value="reader">{t.roles.reader}</option>
      <option value="writer">{t.roles.writer}</option>
    </select>
  );
}

export const CaseAccessPanel: React.FC<CaseAccessPanelProps> = ({ caseId, access, canEdit }) => {
  const t = useCaseText();
  const [entries, setEntries] = React.useState<AccessEntry[] | null>(null);
  const [contact, setContact] = React.useState<ContactChoice | null>(null);
  const [shareBy, setShareBy] = React.useState<"contact" | "email">("contact");
  const [email, setEmail] = React.useState("");
  const [role, setRole] = React.useState<ContactRole>("reader");
  const [error, setError] = React.useState<string | null>(null);
  const [progress, setProgress] = React.useState<string | null>(null);

  const run = async (label: string, work: () => Promise<void>) => {
    setProgress(label);
    setError(null);
    try {
      await work();
      setEntries(await access.list(caseId));
    } catch (caught) {
      setError(errorMessage(caught, t.accessRequestFailed));
    } finally {
      setProgress(null);
    }
  };

  const busy = progress !== null;

  const chooseContact = async () => {
    try {
      const picked = await access.pickContact();
      if (picked) {
        setContact(picked);
      }
    } catch {
      setError(t.contactLookupFailed);
    }
  };

  React.useEffect(() => {
    void run(t.loadingAccess, () => Promise.resolve());
  }, [caseId]);

  return (
    <div className="spe-access">
      <div className="spe-row">
        <h3>{t.peopleWithAccess}</h3>
        {entries && <span className="spe-count">{entries.length}</span>}
      </div>
      {error && <p className="spe-banner spe-banner-error">{error}</p>}
      {progress && <ProgressSpinner label={progress} />}

      {entries && entries.length > 0 && (
        <ul className="spe-people">
          {entries.map((entry) => (
            <li key={entry.id}>
              <span className={`spe-avatar${entry.external ? " external" : ""}`} aria-hidden="true">
                {initials(entry.name || entry.email)}
              </span>
              <div className="spe-file-name">
                <strong>{entry.name || entry.email}</strong>
                <span>{entry.external ? entry.email : t.colleague}</span>
              </div>
              {entry.external && canEdit ? (
                <div className="spe-actions">
                  <RoleSelect
                    value={entry.role}
                    disabled={busy}
                    onChange={(next) => void run(t.updatingAccess, () => access.update(caseId, entry.id, next))}
                  />
                  <button
                    type="button"
                    className="spe-close"
                    aria-label={t.remove(entry.name || entry.email)}
                    disabled={busy}
                    onClick={() => void run(t.removingAccess, () => access.remove(caseId, entry.id))}
                  >
                    ✕
                  </button>
                </div>
              ) : (
                <span className={`spe-role spe-role-${entry.role}`}>{t.roles[entry.role] ?? entry.role}</span>
              )}
            </li>
          ))}
        </ul>
      )}
      {entries?.length === 0 && <p className="spe-hint">{t.nobodyHasAccess}</p>}

      {canEdit && (
        <form
          className="spe-grant"
          onSubmit={(event) => {
            event.preventDefault();
            if (shareBy === "contact" && contact) {
              void run(t.givingAccess(contact.name), async () => {
                await access.grant(caseId, contact.id, role);
                setContact(null);
              });
            } else if (shareBy === "email" && email.trim()) {
              const address = email.trim();
              void run(t.sharingWith(address), async () => {
                await access.shareByEmail(caseId, address, role);
                setEmail("");
              });
            }
          }}
        >
          <div className="spe-row">
            <span className="spe-grant-label">{t.shareWholeCase}</span>
            <div className="spe-segmented" role="tablist" aria-label={t.shareWith}>
              <button type="button" role="tab" aria-selected={shareBy === "contact"} className={shareBy === "contact" ? "on" : ""} onClick={() => setShareBy("contact")}>
                {t.contactInCrm}
              </button>
              <button type="button" role="tab" aria-selected={shareBy === "email"} className={shareBy === "email" ? "on" : ""} onClick={() => setShareBy("email")}>
                {t.anyoneByEmail}
              </button>
            </div>
          </div>
          <div className="spe-grant-row">
            {shareBy === "contact" ? (
              <button
                type="button"
                className={`spe-contact-pick${contact ? " chosen" : ""}`}
                disabled={busy}
                  onClick={() => void chooseContact()}
              >
                {contact ? (
                  <>
                    <span className="spe-avatar external small" aria-hidden="true">
                      {initials(contact.name)}
                    </span>
                    {contact.name}
                  </>
                ) : (
                  t.chooseContact
                )}
              </button>
            ) : (
              <input
                type="email"
                className="spe-email"
                aria-label={t.emailAddress}
                placeholder={t.emailPlaceholder}
                value={email}
                disabled={busy}
                onChange={(event) => setEmail(event.target.value)}
              />
            )}
            <RoleSelect value={role} disabled={busy} onChange={setRole} />
            <button type="submit" disabled={busy || (shareBy === "contact" ? !contact : !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email.trim()))}>
              {shareBy === "contact" ? t.giveAccess : t.share}
            </button>
          </div>
          <span className="spe-hint">{t.wholeCaseHint}</span>
        </form>
      )}
    </div>
  );
};
