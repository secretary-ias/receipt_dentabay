"""Configuration helpers for the receipt application."""
from __future__ import annotations

import base64
import json
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any, Dict, Optional

DEFAULT_LOGO_BASE64 = (
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9p2zBUwAAAAASUVORK5CYII="
)


@dataclass
class ClinicInfo:
    name: str = ""
    address: str = ""
    phone: str = ""
    email: str = ""
    logo_path: str = ""


@dataclass
class DatabaseInfo:
    source: str = "mysql"
    backup_path: str = ""
    cache_path: str = "data/clinic_cache.sqlite"
    mysql_host: str = "localhost"
    mysql_port: int = 3306
    mysql_user: str = ""
    mysql_password: str = ""
    mysql_database: str = "clinicdb"


@dataclass
class ReceiptOptions:
    output_directory: str = "receipts"


@dataclass
class EmailSettings:
    sender: str = ""
    app_password: str = ""
    subject: str = "Clinic Receipt"
    body: str = ("Dear {patient_name},\n\nPlease find attached your receipt {receipt_id}.\n\nThank you.")

@dataclass
class AppSettings:
    clinic: ClinicInfo = field(default_factory=ClinicInfo)
    database: DatabaseInfo = field(default_factory=DatabaseInfo)
    receipt: ReceiptOptions = field(default_factory=ReceiptOptions)
    email: EmailSettings = field(default_factory=EmailSettings)


class ConfigManager:
    """Load and persist application configuration."""

    def __init__(self, config_path: Path) -> None:
        self.config_path = Path(config_path)
        self.settings = AppSettings()
        self.load()

    def load(self) -> None:
        if not self.config_path.exists():
            return
        data = json.loads(self.config_path.read_text(encoding="utf-8"))
        self.settings = self._from_dict(data)
        self._normalise_paths()

    def save(self) -> None:
        self.config_path.parent.mkdir(parents=True, exist_ok=True)
        self._normalise_paths()
        payload = self._to_dict()
        self.config_path.write_text(json.dumps(payload, indent=2), encoding="utf-8")

    def _to_dict(self) -> Dict[str, Any]:
        return {
            "clinic": asdict(self.settings.clinic),
            "database": asdict(self.settings.database),
            "receipt": asdict(self.settings.receipt),
            "email": asdict(self.settings.email),
        }

    def _from_dict(self, data: Dict[str, Any]) -> AppSettings:
        clinic_data = data.get("clinic", {})
        database_data = data.get("database", {})
        receipt_data = data.get("receipt", {})
        email_data = data.get("email", {})
        return AppSettings(
            clinic=ClinicInfo(**{**asdict(ClinicInfo()), **clinic_data}),
            database=DatabaseInfo(**{**asdict(DatabaseInfo()), **database_data}),
            receipt=ReceiptOptions(**{**asdict(ReceiptOptions()), **receipt_data}),
            email=EmailSettings(**{**asdict(EmailSettings()), **email_data}),
        )

    def _normalise_paths(self) -> None:
        self.settings.receipt.output_directory = self._normalise_output_directory(
            self.settings.receipt.output_directory
        )

    def _normalise_output_directory(self, value: str) -> str:
        text_value = (value or "").strip()
        if not text_value:
            return "receipts"
        candidate = Path(text_value)
        base_dir = self.config_path.parent.resolve()

        if candidate.is_absolute():
            resolved = candidate.resolve(strict=False)
            try:
                relative = resolved.relative_to(base_dir)
            except ValueError:
                parts = [part for part in candidate.parts if part not in (candidate.anchor, "")]
                if not parts:
                    return "receipts"
                candidate = Path(parts[-1])
            else:
                candidate = relative
        else:
            parts = []
            for part in candidate.parts:
                if part in ("", "."):
                    continue
                if part == "..":
                    if parts:
                        parts.pop()
                    continue
                parts.append(part)
            if not parts:
                return "receipts"
            candidate = Path(*parts)

        normalised = candidate.as_posix()
        if not normalised or normalised == ".":
            return "receipts"
        return normalised


    def update_clinic(self, **kwargs: Any) -> None:
        for key, value in kwargs.items():
            if hasattr(self.settings.clinic, key):
                setattr(self.settings.clinic, key, value)

    def update_database(self, **kwargs: Any) -> None:
        for key, value in kwargs.items():
            if hasattr(self.settings.database, key):
                setattr(self.settings.database, key, value)

    def update_email(self, **kwargs: Any) -> None:
        for key, value in kwargs.items():
            if hasattr(self.settings.email, key):
                setattr(self.settings.email, key, value)

    def update_receipt(self, **kwargs: Any) -> None:
        for key, value in kwargs.items():
            if hasattr(self.settings.receipt, key):
                if key == "output_directory":
                    value = self._normalise_output_directory(str(value))
                setattr(self.settings.receipt, key, value)

    def resolve_cache_path(self) -> Path:
        cache_path = Path(self.settings.database.cache_path)
        if not cache_path.is_absolute():
            cache_path = self.config_path.parent / cache_path
        return cache_path

    @staticmethod
    def _default_logo_bytes() -> bytes:
        return base64.b64decode(DEFAULT_LOGO_BASE64)

    def resolve_logo_path(self) -> Optional[Path]:
        logo_value = self.settings.clinic.logo_path.strip()
        if logo_value:
            candidate = Path(logo_value)
            if not candidate.is_absolute():
                candidate = self.config_path.parent / candidate
            if candidate.exists():
                return candidate

        assets_dir = self.config_path.parent / 'assets'
        try:
            assets_dir.mkdir(parents=True, exist_ok=True)
        except Exception:
            pass

        for filename in ('clinic_logo.png', 'clinic_logo.jpg', 'clinic_logo.jpeg'):
            candidate = assets_dir / filename
            if candidate.exists():
                return candidate

        default_logo = assets_dir / 'clinic_logo.png'
        if default_logo.exists():
            return default_logo

        try:
            default_logo.write_bytes(self._default_logo_bytes())
            return default_logo
        except Exception:
            return None

    def resolve_output_dir(self) -> Path:
        output_value = self._normalise_output_directory(self.settings.receipt.output_directory)
        if output_value != self.settings.receipt.output_directory:
            self.settings.receipt.output_directory = output_value
        out_dir = (self.config_path.parent / Path(output_value)).resolve()
        out_dir.mkdir(parents=True, exist_ok=True)
        return out_dir

    def mysql_settings(self) -> Dict[str, Any]:
        db = self.settings.database
        return {
            "host": db.mysql_host,
            "port": db.mysql_port,
            "user": db.mysql_user,
            "password": db.mysql_password,
            "database": db.mysql_database,
        }


