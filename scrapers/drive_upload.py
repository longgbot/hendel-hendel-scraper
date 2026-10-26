import argparse
import json
import mimetypes
import os
import sys
from pathlib import Path

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

SCOPES = ["https://www.googleapis.com/auth/drive"]


def get_service():
    raw = os.environ.get("GDRIVE_SERVICE_ACCOUNT_JSON")
    if not raw:
        sys.exit("GDRIVE_SERVICE_ACCOUNT_JSON is not set")
    info = json.loads(raw)
    creds = service_account.Credentials.from_service_account_info(info, scopes=SCOPES)
    return build("drive", "v3", credentials=creds, cache_discovery=False)


def escape(name):
    return name.replace("\\", "\\\\").replace("'", "\\'")


def find_child(service, parent_id, name, folder):
    mime_clause = "mimeType = 'application/vnd.google-apps.folder'" if folder else "mimeType != 'application/vnd.google-apps.folder'"
    q = f"name = '{escape(name)}' and '{parent_id}' in parents and trashed = false and {mime_clause}"
    res = service.files().list(q=q, fields="files(id, name, md5Checksum, size)", pageSize=1, supportsAllDrives=True, includeItemsFromAllDrives=True).execute()
    files = res.get("files", [])
    return files[0] if files else None


def ensure_folder(service, parent_id, name, cache):
    key = (parent_id, name)
    if key in cache:
        return cache[key]
    existing = find_child(service, parent_id, name, folder=True)
    if existing:
        cache[key] = existing["id"]
        return existing["id"]
    meta = {"name": name, "mimeType": "application/vnd.google-apps.folder", "parents": [parent_id]}
    created = service.files().create(body=meta, fields="id", supportsAllDrives=True).execute()
    cache[key] = created["id"]
    return created["id"]


def upload_file(service, parent_id, path, overwrite):
    existing = find_child(service, parent_id, path.name, folder=False)
    if existing and not overwrite:
        return "skipped"
    mime = mimetypes.guess_type(str(path))[0] or "application/octet-stream"
    media = MediaFileUpload(str(path), mimetype=mime, resumable=path.stat().st_size > 5_000_000)
    if existing:
        service.files().update(fileId=existing["id"], media_body=media, supportsAllDrives=True).execute()
        return "updated"
    service.files().create(body={"name": path.name, "parents": [parent_id]}, media_body=media, fields="id", supportsAllDrives=True).execute()
    return "created"


def main():
    ap = argparse.ArgumentParser(description="Mirror a local folder tree into a Google Drive folder")
    ap.add_argument("--src", required=True)
    ap.add_argument("--folder-id", default=os.environ.get("DRIVE_FOLDER_ID", ""))
    ap.add_argument("--overwrite-ext", default=".xlsx,.csv,.json,.txt,.html", help="Extensions re-uploaded every run; everything else is uploaded only if missing")
    args = ap.parse_args()
    if not args.folder_id:
        sys.exit("DRIVE_FOLDER_ID is not set")

    overwrite_ext = {e.strip().lower() for e in args.overwrite_ext.split(",") if e.strip()}
    service = get_service()
    src = Path(args.src)
    cache = {}
    counts = {"created": 0, "updated": 0, "skipped": 0}

    for path in sorted(src.rglob("*")):
        if path.is_dir():
            continue
        rel = path.relative_to(src)
        parent = args.folder_id
        for part in rel.parts[:-1]:
            parent = ensure_folder(service, parent, part, cache)
        status = upload_file(service, parent, path, overwrite=path.suffix.lower() in overwrite_ext)
        counts[status] += 1
        if counts["created"] % 25 == 0 and status == "created":
            print(f"  {counts['created']} files created so far", flush=True)

    print(f"Drive sync complete: {counts}")


if __name__ == "__main__":
    main()
