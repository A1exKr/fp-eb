"""Asset file storage.

Persists CV / experience files. Uses the exaBase File Manager when configured
(``FPGEN_FILE_MANAGER_URL``); otherwise stores to a local/persistent-volume
directory. The File Manager HTTP upload API for this workspace is not yet
confirmed, so volume storage is the default and the File Manager call is left
as a clearly marked hook.
"""
import mimetypes
from pathlib import Path

from app.config import settings


def save_asset_file(asset_id: str, filename: str, content: bytes) -> dict:
    """Store the file and return metadata (storage_ref, mime_type, size_bytes)."""
    mime = mimetypes.guess_type(filename)[0] or "application/octet-stream"
    storage_ref = _save_to_volume(asset_id, filename, content)
    return {"storage_ref": storage_ref, "mime_type": mime, "size_bytes": len(content)}


def _save_to_volume(asset_id: str, filename: str, content: bytes) -> str:
    settings.asset_storage_dir.mkdir(parents=True, exist_ok=True)
    suffix = Path(filename).suffix
    safe_id = "".join(c if c.isalnum() or c in "-_" else "_" for c in asset_id) or "asset"
    dest = settings.asset_storage_dir / f"{safe_id}{suffix}"
    dest.write_bytes(content)
    return str(dest)


# NOTE: Once the exaBase File Manager upload API is confirmed, add a
# `_save_via_file_manager(...)` that POSTs to `settings.file_manager_url` and
# returns the file reference/URL, and prefer it when that URL is configured.
