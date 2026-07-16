from __future__ import annotations

import base64
import hashlib
import json
import os
import platform
import subprocess
from datetime import date
from pathlib import Path

from cryptography.exceptions import InvalidSignature
from cryptography.hazmat.primitives.asymmetric.ed25519 import Ed25519PublicKey


APP_FOLDER = "FinishInvoiceManager"
PUBLIC_KEY_B64 = "wcFxj5VY76uuzG0IgERzJmG/sJVNt/7K2RBdjwXX7zw="


def data_dir() -> Path:
    path = Path(os.getenv("LOCALAPPDATA", Path.home())) / APP_FOLDER
    path.mkdir(parents=True, exist_ok=True)
    return path


def license_path() -> Path:
    return data_dir() / "license.lic"


def _windows_machine_guid() -> str:
    if os.name != "nt":
        return ""
    try:
        import winreg

        flags = winreg.KEY_READ | getattr(winreg, "KEY_WOW64_64KEY", 0)
        with winreg.OpenKey(
            winreg.HKEY_LOCAL_MACHINE,
            r"SOFTWARE\Microsoft\Cryptography",
            0,
            flags,
        ) as key:
            return str(winreg.QueryValueEx(key, "MachineGuid")[0]).strip()
    except OSError:
        return ""


def _bios_uuid() -> str:
    if os.name != "nt":
        return ""
    try:
        creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
        result = subprocess.run(
            ["powershell.exe", "-NoProfile", "-NonInteractive", "-Command",
             "(Get-CimInstance Win32_ComputerSystemProduct).UUID"],
            capture_output=True,
            text=True,
            timeout=8,
            creationflags=creationflags,
            check=False,
        )
        return result.stdout.strip()
    except (OSError, subprocess.SubprocessError):
        return ""


def machine_code() -> str:
    # MachineGuid is stable for a Windows installation; BIOS UUID keeps copied
    # Windows images from producing the same code on different computers.
    parts = [_windows_machine_guid(), _bios_uuid()]
    stable = "|".join(part.upper() for part in parts if part)
    if not stable:
        stable = f"{platform.node()}|{platform.machine()}|{platform.system()}"
    digest = hashlib.sha256(("FinishInvoiceManager|" + stable).encode("utf-8")).hexdigest().upper()
    return "FIM-" + "-".join(digest[i:i + 4] for i in range(0, 20, 4))


def canonical_payload(payload: dict) -> bytes:
    return json.dumps(payload, sort_keys=True, separators=(",", ":"), ensure_ascii=False).encode("utf-8")


def read_and_validate_license(path: Path | None = None) -> tuple[bool, str, dict | None]:
    path = path or license_path()
    if not path.exists():
        return False, "License file is not installed.", None
    try:
        document = json.loads(path.read_text(encoding="utf-8"))
        payload = document["payload"]
        signature = base64.b64decode(document["signature"], validate=True)
        public_key = Ed25519PublicKey.from_public_bytes(base64.b64decode(PUBLIC_KEY_B64))
        public_key.verify(signature, canonical_payload(payload))
    except (OSError, ValueError, KeyError, TypeError, json.JSONDecodeError, InvalidSignature):
        return False, "License file is invalid or has been modified.", None

    if payload.get("product") != "FinishInvoiceManager":
        return False, "This license belongs to a different product.", None
    if str(payload.get("machine_code", "")).upper() != machine_code().upper():
        return False, "This license was issued for a different computer.", None
    expires_on = payload.get("expires_on")
    if expires_on:
        try:
            if date.today() > date.fromisoformat(expires_on):
                return False, f"License expired on {expires_on}.", None
        except ValueError:
            return False, "License expiry date is invalid.", None
    return True, "License is valid.", payload


def install_license(source: str | Path) -> tuple[bool, str]:
    source = Path(source)
    valid, message, _ = read_and_validate_license(source)
    if not valid:
        return False, message
    license_path().write_bytes(source.read_bytes())
    return True, "License activated successfully."
