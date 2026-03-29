from __future__ import annotations

from io import BytesIO
from pathlib import Path
from typing import Optional

import msoffcrypto
import pandas as pd


def _decrypt_office_file(file_path: str | Path, password: Optional[str]) -> BytesIO:
    """
    Decrypt password-protected legacy Office file into memory.
    """
    source = Path(file_path)
    decrypted = BytesIO()
    with source.open("rb") as f:
        office_file = msoffcrypto.OfficeFile(f)
        office_file.load_key(password=password or "")
        office_file.decrypt(decrypted)
    decrypted.seek(0)
    return decrypted


def load_payroll_xls(file_path: str | Path, password: Optional[str] = None) -> pd.DataFrame:
    """
    Load payroll source (.xls) to DataFrame. Supports encrypted files via msoffcrypto.
    """
    payload = _decrypt_office_file(file_path=file_path, password=password)
    df = pd.read_excel(payload, engine="xlrd")
    return df


def load_payroll_xls_from_bytes(file_bytes: bytes, password: Optional[str] = None) -> pd.DataFrame:
    """
    Load payroll source from uploaded bytes (supports encrypted xls).
    """
    decrypted = BytesIO()
    office_file = msoffcrypto.OfficeFile(BytesIO(file_bytes))
    office_file.load_key(password=password or "")
    office_file.decrypt(decrypted)
    decrypted.seek(0)
    return pd.read_excel(decrypted, engine="xlrd")


def payroll_file_is_encrypted(file_path: str | Path) -> bool:
    with Path(file_path).open("rb") as f:
        office_file = msoffcrypto.OfficeFile(f)
        return bool(office_file.is_encrypted())


def payroll_bytes_is_encrypted(file_bytes: bytes) -> bool:
    office_file = msoffcrypto.OfficeFile(BytesIO(file_bytes))
    return bool(office_file.is_encrypted())
