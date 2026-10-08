"""
Zip JJ's Setlist (browser edition) into one package that works on Windows,
Mac and Linux.  Run:  python make_package.py   (or Make Package.bat)

    python make_package.py            for your band: includes the song
                                      spreadsheet, the saved setlists and the
                                      Google Drive link
    python make_package.py --public   for anyone: the app only, no personal files

The zip records Unix permissions and the right line endings for each system,
so the Mac and Linux launchers stay double-clickable / runnable after unzipping.
"""

import json
import sys
import time
import zipfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
OUT_DIR = HERE / "packages"
TOP = "JJs Setlist"                                   # the folder inside the zip

APP_FILES = [
    "serve.py",
    "Start JJ's Setlist.bat",
    "Start JJ's Setlist.command",
    "start-jjs-setlist.sh",
    "READ ME FIRST.txt",
    "README.md",
    "LICENSE",
]
# Must be executable on Mac / Linux (double-clickable / runnable).
EXECUTABLE = {"Start JJ's Setlist.command", "start-jjs-setlist.sh"}
LF_ONLY = (".command", ".sh", ".py", ".js", ".css", ".html", ".md", ".json")   # Unix line endings
CRLF = (".bat", ".txt")                                                     # Windows line endings
TEXT = LF_ONLY + CRLF + (".webmanifest",)


def config():
    try:
        return json.loads((HERE / "config.json").read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}


def song_database():
    """The song spreadsheet in this folder, if there's exactly one obvious one."""
    sheets = [p for p in HERE.glob("*.xlsx") if not p.name.startswith(("~$", "JJ-SETLIST-TEMPLATE"))]
    return sheets[0].name if len(sheets) == 1 else None


def file_bytes(name):
    data = (HERE / name).read_bytes()
    if name.endswith(TEXT) and not name.startswith("web/lib/"):
        text = data.decode("utf-8").replace("\r\n", "\n")
        if name.endswith(CRLF):
            text = text.replace("\n", "\r\n")
        data = text.encode("utf-8")
    return data


def add(z, arcname, data, mode=0o644, mtime=None):
    info = zipfile.ZipInfo(f"{TOP}/{arcname}", time.localtime(mtime or time.time())[:6])
    info.create_system = 3                              # 3 = Unix, so the permissions are used
    info.external_attr = (0o100000 | mode) << 16
    info.compress_type = zipfile.ZIP_DEFLATED
    z.writestr(info, data)


def add_folder(z, arcname):
    info = zipfile.ZipInfo(f"{TOP}/{arcname}/")
    info.create_system = 3
    info.external_attr = ((0o040000 | 0o755) << 16) | 0x10
    z.writestr(info, b"")


def main():
    public = "--public" in sys.argv
    files = list(APP_FILES)
    files += sorted(str(p.relative_to(HERE)).replace("\\", "/")
                    for p in (HERE / "web").rglob("*") if p.is_file())
    personal = []
    if not public:
        db = song_database()
        if db:
            personal.append(db)
        personal += sorted(f"setlists/{p.name}" for p in (HERE / "setlists").glob("*.json"))
    missing = [f for f in files + personal if not (HERE / f).is_file()]
    if missing:
        print("Missing:", ", ".join(missing))
        return 1

    OUT_DIR.mkdir(exist_ok=True)
    zip_path = OUT_DIR / ("JJs Setlist - Browser Edition (public).zip" if public
                          else "JJs Setlist - Browser Edition.zip")
    tmp = zip_path.with_suffix(".zip.part")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
        for name in files + personal:
            mode = 0o755 if name in EXECUTABLE else 0o644
            add(z, name, file_bytes(name), mode, (HERE / name).stat().st_mtime)
            print(f"  added   {name}" + ("   (executable)" if mode == 0o755 else ""))
        if public or not any(n.startswith("setlists/") for n in personal):
            add_folder(z, "setlists")                   # an empty folder, ready to use
        # The Google Drive link (not the colours etc.), so the song database
        # loads from Drive straight away.
        link = None if public else config().get("backup_url")
        if link:
            settings = {"backup_url": link, "drive_first": bool(config().get("drive_first", True))}
            add(z, "config.json", (json.dumps(settings, indent=2) + "\n").encode("utf-8"))
            print("  added   config.json   (Google Drive link of the song database)")
    tmp.replace(zip_path)
    print(f"\n  Package:  {zip_path}")
    print("  Unzip it on Windows, Mac or Linux and follow READ ME FIRST.txt.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
