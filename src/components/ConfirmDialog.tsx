"use client";

import { useEffect, useRef } from "react";

interface ConfirmDialogProps {
  title: string;
  message: string;
  confirmLabel?: string;
  cancelLabel?: string;
  variant?: "default" | "danger";
  onConfirm: () => void;
  onCancel: () => void;
}

export function ConfirmDialog({
  title,
  message,
  confirmLabel = "Confirm",
  cancelLabel = "Cancel",
  variant = "default",
  onConfirm,
  onCancel,
}: ConfirmDialogProps) {
  const dialogRef = useRef<HTMLDialogElement>(null);

  useEffect(() => {
    dialogRef.current?.showModal();
  }, []);

  return (
    <dialog ref={dialogRef} className="modal" onClose={onCancel}>
      <div className="modal-box" style={{ fontSize: 13 }}>
        <h3 style={{ fontSize: 14, fontWeight: 600 }}>{title}</h3>
        <p style={{ padding: "12px 0", color: "var(--text-muted)", lineHeight: 1.5 }}>{message}</p>
        <div style={{ display: "flex", gap: 8, justifyContent: "flex-end" }}>
          <form method="dialog">
            <button className="btn btn-ghost btn-sm" style={{ color: "var(--text-muted)" }}>{cancelLabel}</button>
          </form>
          <button
            className={`btn btn-sm ${variant === "danger" ? "btn-error" : "btn-primary"}`}
            onClick={onConfirm}
          >
            {confirmLabel}
          </button>
        </div>
      </div>
      <form method="dialog" className="modal-backdrop">
        <button>close</button>
      </form>
    </dialog>
  );
}
