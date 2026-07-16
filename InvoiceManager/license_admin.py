from __future__ import annotations

import base64
import json
import uuid
from datetime import date
from pathlib import Path
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from cryptography.hazmat.primitives.asymmetric.ed25519 import Ed25519PrivateKey

from licensing import canonical_payload


KEY_FILE_NAME = "FinishLicenseAdmin.key"


def private_key_path() -> Path:
    base = Path(sys.executable).parent if getattr(sys, "frozen", False) else Path(__file__).parent
    return base / KEY_FILE_NAME


def load_private_key() -> Ed25519PrivateKey:
    path = private_key_path()
    if not path.exists():
        raise FileNotFoundError(
            f"Private key file not found:\n{path}\n\n"
            "FinishLicenseAdmin.key ko Admin EXE ke same folder mein rakhein."
        )
    raw = base64.b64decode(path.read_text(encoding="ascii").strip(), validate=True)
    return Ed25519PrivateKey.from_private_bytes(raw)


class LicenseAdmin(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Finish Invoice Manager - License Generator")
        self.geometry("650x410")
        self.resizable(False, False)
        self.configure(bg="#F3F6F9")
        self.vars = {
            "customer": tk.StringVar(),
            "machine_code": tk.StringVar(),
            "expires_on": tk.StringVar(),
        }
        self._build()

    def _build(self):
        header = tk.Frame(self, bg="#123B4A", height=85)
        header.pack(fill="x")
        header.pack_propagate(False)
        tk.Label(header, text="License Generator", bg="#123B4A", fg="white",
                 font=("Segoe UI Semibold", 20)).pack(anchor="w", padx=28, pady=(14, 0))
        tk.Label(header, text="Keep this application with the software owner only.",
                 bg="#123B4A", fg="#C8E1E8", font=("Segoe UI", 10)).pack(anchor="w", padx=28)

        body = ttk.Frame(self, padding=28)
        body.pack(fill="both", expand=True)
        fields = (
            ("Customer / Company", "customer"),
            ("Client Machine Code", "machine_code"),
            ("Expiry (YYYY-MM-DD)", "expires_on"),
        )
        for row, (label, key) in enumerate(fields):
            ttk.Label(body, text=label).grid(row=row, column=0, sticky="w", pady=10)
            ttk.Entry(body, textvariable=self.vars[key], width=48).grid(row=row, column=1, padx=(18, 0), pady=10)
        ttk.Label(body, text="Expiry blank chhor dein to license lifetime hoga.").grid(
            row=3, column=1, sticky="w", padx=(18, 0))
        ttk.Button(body, text="Generate .lic File", command=self.generate).grid(
            row=4, column=1, sticky="e", padx=(18, 0), pady=30)

    def generate(self):
        customer = self.vars["customer"].get().strip()
        machine = self.vars["machine_code"].get().strip().upper()
        expires = self.vars["expires_on"].get().strip()
        if not customer or not machine:
            messagebox.showerror("Missing information", "Customer aur Machine Code required hain.")
            return
        if not machine.startswith("FIM-"):
            messagebox.showerror("Invalid Machine Code", "Client app ka FIM- wala Machine Code paste karein.")
            return
        if expires:
            try:
                date.fromisoformat(expires)
            except ValueError:
                messagebox.showerror("Invalid expiry", "Expiry YYYY-MM-DD format mein likhein.")
                return

        payload = {
            "product": "FinishInvoiceManager",
            "license_id": "LIC-" + uuid.uuid4().hex[:12].upper(),
            "customer": customer,
            "machine_code": machine,
            "issued_on": date.today().isoformat(),
            "expires_on": expires or None,
        }
        try:
            private_key = load_private_key()
        except (OSError, ValueError) as exc:
            messagebox.showerror("Private key unavailable", str(exc))
            return
        document = {
            "payload": payload,
            "signature": base64.b64encode(private_key.sign(canonical_payload(payload))).decode("ascii"),
        }
        safe_name = "".join(c if c.isalnum() or c in "-_" else "-" for c in customer).strip("-")
        target = filedialog.asksaveasfilename(
            title="Save client license",
            initialfile=f"{safe_name or 'client'}.lic",
            defaultextension=".lic",
            filetypes=[("License file", "*.lic")],
        )
        if not target:
            return
        Path(target).write_text(json.dumps(document, indent=2), encoding="utf-8")
        messagebox.showinfo("License ready", f"WhatsApp par Document ke taur par bhejein:\n{target}")


if __name__ == "__main__":
    LicenseAdmin().mainloop()
