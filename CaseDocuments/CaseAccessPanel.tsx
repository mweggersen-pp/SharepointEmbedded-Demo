import * as React from "react";

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
  update(caseId: string, permissionId: string, role: ContactRole): Promise<void>;
  remove(caseId: string, permissionId: string): Promise<void>;
}

interface CaseAccessPanelProps {
  caseId: string;
  access: CaseAccess;
  canEdit: boolean;
}

const roleLabels: Record<string, string> = { reader: "Read", writer: "Edit", manager: "Manage", owner: "Owner" };

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
  const [role, setRole] = React.useState<ContactRole>("reader");
  const [error, setError] = React.useState<string | null>(null);
  const [busy, setBusy] = React.useState(false);

  const run = async (work: () => Promise<void>) => {
    setBusy(true);
    setError(null);
    try {
      await work();
      setEntries(await access.list(caseId));
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : "The access request failed.");
    } finally {
      setBusy(false);
    }
  };

  React.useEffect(() => {
    void run(() => Promise.resolve());
  }, [caseId]);

  return (
    <section className="spe-access">
      <h3>People with access</h3>
      {error && <p className="spe-message spe-error">{error}</p>}
      {entries === null && !error && <p className="spe-message">Loading who has access.</p>}

      {entries && (
        <ul className="spe-files">
          {entries.map((entry) => (
            <li key={entry.id}>
              <div className="spe-row">
                <div>
                  <strong>{entry.name || entry.email}</strong>
                  <div>{entry.external ? `${entry.email} · External contact` : "Colleague · follows Dynamics access"}</div>
                </div>
                {entry.external && canEdit ? (
                  <div className="spe-actions">
                    <RoleSelect
                      value={entry.role}
                      disabled={busy}
                      onChange={(next) => void run(() => access.update(caseId, entry.id, next))}
                    />
                    <button type="button" className="secondary" disabled={busy} onClick={() => void run(() => access.remove(caseId, entry.id))}>
                      Remove
                    </button>
                  </div>
                ) : (
                  <span>{roleLabels[entry.role] ?? entry.role}</span>
                )}
              </div>
            </li>
          ))}
        </ul>
      )}
      {entries?.length === 0 && <p className="spe-message">Nobody has access to this case container yet.</p>}

      {canEdit && (
        <div className="spe-row spe-grant">
          <button
            type="button"
            className="secondary"
            disabled={busy}
            onClick={() =>
              void access.pickContact().then(
                (picked) => picked && setContact(picked),
                () => setError("Could not open the contact lookup.")
              )
            }
          >
            {contact ? contact.name : "Choose contact"}
          </button>
          <RoleSelect value={role} disabled={busy} onChange={setRole} />
          <button
            type="button"
            disabled={busy || !contact}
            onClick={() =>
              void run(async () => {
                if (contact) {
                  await access.grant(caseId, contact.id, role);
                  setContact(null);
                }
              })
            }
          >
            Give access
          </button>
        </div>
      )}
    </section>
  );
};
