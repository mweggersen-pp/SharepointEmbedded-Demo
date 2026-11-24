import * as React from "react";
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
}

export interface CaseAccess {
  list(caseId: string): Promise<AccessEntry[]>;
  pickContact(): Promise<ContactChoice | null>;
  grant(caseId: string, contactId: string, role: ContactRole): Promise<void>;
  shareByEmail(caseId: string, email: string, role: ContactRole): Promise<void>;
  update(caseId: string, permissionId: string, role: ContactRole): Promise<void>;
  remove(caseId: string, permissionId: string): Promise<void>;
}

interface CaseAccessPanelProps {
  caseId: string;
  access: CaseAccess;
  canEdit: boolean;
}

const roleLabels: Record<string, string> = { reader: "Read", writer: "Edit", manager: "Manage", owner: "Owner" };

function initials(name: string): string {
  const parts = name.trim().split(/\s+/).filter(Boolean);
  return ((parts[0]?.[0] ?? "?") + (parts.length > 1 ? parts[parts.length - 1][0] : "")).toUpperCase();
}

function RoleSelect({ value, disabled, onChange }: { value: string; disabled: boolean; onChange: (role: ContactRole) => void }) {
  return (
    <select aria-label="Access" value={value} disabled={disabled} onChange={(event) => onChange(event.target.value as ContactRole)}>
      <option value="reader">Read</option>
      <option value="writer">Edit</option>
    </select>
  );
}

export const CaseAccessPanel: React.FC<CaseAccessPanelProps> = ({ caseId, access, canEdit }) => {
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
      setError(errorMessage(caught, "The access request failed."));
    } finally {
      setProgress(null);
    }
  };

  const busy = progress !== null;

  React.useEffect(() => {
    void run("Loading who has access", () => Promise.resolve());
  }, [caseId]);

  return (
    <div className="spe-access">
      <div className="spe-row">
        <h3>People with access</h3>
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
                <span>{entry.external ? entry.email : "Colleague · follows Dynamics access"}</span>
              </div>
              {entry.external && canEdit ? (
                <div className="spe-actions">
                  <RoleSelect
                    value={entry.role}
                    disabled={busy}
                    onChange={(next) => void run("Updating access", () => access.update(caseId, entry.id, next))}
                  />
                  <button
                    type="button"
                    className="spe-close"
                    aria-label={`Remove ${entry.name || entry.email}`}
                    disabled={busy}
                    onClick={() => void run("Removing access", () => access.remove(caseId, entry.id))}
                  >
                    ✕
                  </button>
                </div>
              ) : (
                <span className={`spe-role spe-role-${entry.role}`}>{roleLabels[entry.role] ?? entry.role}</span>
              )}
            </li>
          ))}
        </ul>
      )}
      {entries?.length === 0 && <p className="spe-hint">Nobody has access to this case container yet.</p>}

      {canEdit && (
        <form
          className="spe-grant"
          onSubmit={(event) => {
            event.preventDefault();
            if (shareBy === "contact" && contact) {
              void run(`Giving ${contact.name} access`, async () => {
                await access.grant(caseId, contact.id, role);
                setContact(null);
              });
            } else if (shareBy === "email" && email.trim()) {
              const address = email.trim();
              void run(`Sharing with ${address}`, async () => {
                await access.shareByEmail(caseId, address, role);
                setEmail("");
              });
            }
          }}
        >
          <div className="spe-row">
            <span className="spe-grant-label">Share with someone outside the company</span>
            <div className="spe-segmented" role="tablist" aria-label="Share with">
              <button type="button" role="tab" aria-selected={shareBy === "contact"} className={shareBy === "contact" ? "on" : ""} onClick={() => setShareBy("contact")}>
                Contact in CRM
              </button>
              <button type="button" role="tab" aria-selected={shareBy === "email"} className={shareBy === "email" ? "on" : ""} onClick={() => setShareBy("email")}>
                Anyone by email
              </button>
            </div>
          </div>
          <div className="spe-grant-row">
            {shareBy === "contact" ? (
              <button
                type="button"
                className={`spe-contact-pick${contact ? " chosen" : ""}`}
                disabled={busy}
                onClick={() =>
                  void access.pickContact().then(
                    (picked) => picked && setContact(picked),
                    () => setError("Could not open the contact lookup.")
                  )
                }
              >
                {contact ? (
                  <>
                    <span className="spe-avatar external small" aria-hidden="true">
                      {initials(contact.name)}
                    </span>
                    {contact.name}
                  </>
                ) : (
                  "＋ Choose contact"
                )}
              </button>
            ) : (
              <input
                type="email"
                className="spe-email"
                aria-label="Email address"
                placeholder="name@company.com"
                value={email}
                disabled={busy}
                onChange={(event) => setEmail(event.target.value)}
              />
            )}
            <RoleSelect value={role} disabled={busy} onChange={setRole} />
            <button type="submit" disabled={busy || (shareBy === "contact" ? !contact : !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email.trim()))}>
              {shareBy === "contact" ? "Give access" : "Share"}
            </button>
          </div>
          <span className="spe-hint">They get an email invitation and sign in with their own work or Microsoft account.</span>
        </form>
      )}
    </div>
  );
};
