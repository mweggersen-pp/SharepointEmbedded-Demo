import * as React from "react";

interface ProgressSpinnerProps {
  label: string;
  detail?: string;
  size?: "large" | "small";
}

export const ProgressSpinner: React.FC<ProgressSpinnerProps> = ({ label, detail, size = "small" }) => (
  <div className={`spe-progress spe-progress-${size}`} role="status" aria-live="polite">
    <span className="spe-wheel" aria-hidden="true" />
    <span className="spe-progress-text">
      <span>{label}</span>
      {detail && <span className="spe-progress-detail">{detail}</span>}
    </span>
  </div>
);
