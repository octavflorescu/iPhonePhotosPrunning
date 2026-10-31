#!/usr/bin/env python3
"""
Delete earlier iPhone photos over USB (MTP) when the next photo in the
**same calendar month** was taken within N seconds. Month comes from the
album folder name (e.g. 201705_a = May 2017); photos are never compared
across months.

Runs on Windows with the phone unlocked and trusted. Uses Shell COM (not a drive letter).
Default is dry-run; pass --execute to delete on the device.

Caches EXIF timestamps in exif_cache_YYYY.csv next to this script (see --exif-cache-dir).
MTP/USB work is single-threaded during scan; delete can use --delete-parallel for albums.

Requirements:
  pip install pywin32 piexif pillow pillow-heif
"""

from __future__ import annotations

import argparse
import csv
import os
import re
import signal
import sys
import tempfile
import time
from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Iterable

SCRIPT_DIR = Path(__file__).resolve().parent
EXIF_CACHE_FILENAME = "exif_cache_{year}.csv"
EXIF_CACHE_FIELDS = ("album", "name", "type_label", "taken", "precision")

# iPhone DCIM folders: 201705_a → May 2017, 202510_a → Oct 2025
ALBUM_YEAR_MONTH = re.compile(r"^(\d{4})(\d{2})")

try:
    import pythoncom
    import win32com.client
except ImportError as exc:
    raise SystemExit("Install pywin32: pip install pywin32") from exc

try:
    import piexif
except ImportError as exc:
    raise SystemExit("Install piexif: pip install piexif") from exc

try:
    import pillow_heif

    pillow_heif.register_heif_opener()
except ImportError:
    pass

from PIL import Image
from PIL.ExifTags import TAGS
from win32com.shell import shell, shellcon

# FileOperation flags: silent, no confirm, no error UI, no connected elements
FOF_FLAGS = 0x0004 | 0x0010 | 0x0400 | 0x0100

IMAGE_TYPE_HINTS = ("JPG", "JPEG", "HEIC", "PNG", "DNG", "TIFF", "WEBP")
VIDEO_TYPE_HINTS = ("MOV", "M4V")

# MTP CopyHere can block for minutes; Ctrl+C often does not fire until it returns.
DEFAULT_COPY_TIMEOUT_SEC = 90.0


def log(*parts: object) -> None:
    print(*parts, flush=True)


def install_hard_interrupt() -> None:
    def _on_sigint(_signum: int, _frame: object) -> None:
        log("\nCtrl+C received. If the process does not exit, kill python.exe in Task Manager.")
        os._exit(130)

    signal.signal(signal.SIGINT, _on_sigint)
    if hasattr(signal, "SIGBREAK"):
        signal.signal(signal.SIGBREAK, _on_sigint)


@dataclass
class PhotoEntry:
    album: str
    name: str
    type_label: str
    taken: datetime
    item: Any


def parse_shell_datetime(raw: str) -> datetime | None:
    raw = raw.strip()
    if not raw:
        return None
    patterns = (
        "%d-%b-%y %H:%M",
        "%d-%b-%Y %H:%M",
        "%m/%d/%Y %H:%M",
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d %H:%M",
    )
    for fmt in patterns:
        try:
            return datetime.strptime(raw, fmt)
        except ValueError:
            continue
    return None


def build_detail_column_map(shell: Any) -> dict[str, int]:
    meta = shell.NameSpace(0)
    columns: dict[str, int] = {}
    for idx in range(512):
        label = meta.GetDetailsOf(None, idx)
        if not label:
            break
        columns[str(label).strip().lower()] = idx
    return columns


def copy_shell_item_to_dir(
    shell: Any,
    dest_dir: str,
    item: Any,
    timeout: float = DEFAULT_COPY_TIMEOUT_SEC,
) -> Path | None:
    dest_ns = shell.NameSpace(dest_dir)
    if dest_ns is None:
        return None
    before = set(os.listdir(dest_dir))
    dest_ns.CopyHere(item, FOF_FLAGS)
    deadline = time.time() + timeout
    while time.time() < deadline:
        after = set(os.listdir(dest_dir))
        new_names = after - before
        if new_names:
            name = sorted(new_names)[0]
            return Path(dest_dir) / name
        time.sleep(0.25)
    return None


def exif_datetime_from_path(path: Path) -> datetime | None:
    if path.suffix.lower() in {".jpg", ".jpeg", ".heic"}:
        try:
            exif = piexif.load(str(path))
            raw = exif.get("Exif", {}).get(piexif.ExifIFD.DateTimeOriginal)
            if not raw:
                raw = exif.get("0th", {}).get(piexif.ImageIFD.DateTime)
            if raw:
                return datetime.strptime(raw.decode("ascii"), "%Y:%m:%d %H:%M:%S")
        except Exception:
            pass
    try:
        with Image.open(path) as im:
            exif = im.getexif()
            if exif:
                for tag_id, value in exif.items():
                    if TAGS.get(tag_id) in ("DateTimeOriginal", "DateTime"):
                        return datetime.strptime(value, "%Y:%m:%d %H:%M:%S")
    except Exception:
        pass
    return None


def item_datetime_shell(
    album_folder: Any, item: Any, columns: dict[str, int]
) -> datetime | None:
    for key in ("date taken", "date created", "date modified", "date"):
        idx = columns.get(key)
        if idx is None:
            continue
        raw = album_folder.GetDetailsOf(item, idx)
        if not raw:
            continue
        dt = parse_shell_datetime(str(raw))
        if dt and dt.year >= 2000:
            return dt
    return None


def is_image_type(type_label: str) -> bool:
    upper = type_label.upper()
    if any(v in upper for v in VIDEO_TYPE_HINTS) and not any(
        h in upper for h in IMAGE_TYPE_HINTS
    ):
        return False
    return any(h in upper for h in IMAGE_TYPE_HINTS)


def is_mov_type(type_label: str) -> bool:
    upper = type_label.upper()
    return "MOV" in upper or "M4V" in upper


def parse_album_year_month(album: str) -> tuple[int, int] | None:
    m = ALBUM_YEAR_MONTH.match(album)
    if not m:
        return None
    year, month = int(m.group(1)), int(m.group(2))
    if 1 <= month <= 12:
        return year, month
    return None


def photo_year_month(photo: PhotoEntry) -> tuple[int, int]:
    ym = parse_album_year_month(photo.album)
    if ym is not None:
        return ym
    return photo.taken.year, photo.taken.month


def list_album_entries(storage_folder: Any) -> list[tuple[str, Any]]:
    """List DCIM album folder names without opening every folder (much faster than GetFolder×N)."""
    log("Listing album folder names on device…")
    entries: list[tuple[str, Any]] = []
    try:
        count = storage_folder.Items().Count
    except Exception as exc:
        raise RuntimeError(f"Cannot enumerate Internal Storage: {exc}") from exc
    for i in range(count):
        album_item = storage_folder.Items().Item(i)
        if album_item:
            entries.append((str(album_item.Name), album_item))
    entries.sort(key=lambda t: (parse_album_year_month(t[0]) or (9999, 99), t[0]))
    log(f"Found {len(entries)} album folders.")
    return entries


def iter_album_items(storage_folder: Any) -> Iterable[tuple[str, Any, Any]]:
    for i in range(storage_folder.Items().Count):
        album_item = storage_folder.Items().Item(i)
        if not album_item:
            continue
        try:
            album = album_item.GetFolder
        except Exception:
            continue
        if album is None:
            continue
        yield album_item.Name, album_item, album


def find_device_storage(
    shell: Any,
    device_name: str,
    storage_name: str,
    *,
    wait_sec: float = 45.0,
    poll_sec: float = 2.0,
) -> Any:
    """Wait for MTP to expose Internal Storage (often delayed until unlock + Trust)."""
    deadline = time.time() + wait_sec
    storage_key = storage_name.strip().lower()
    device_key = device_name.strip().lower()
    last_diag = ""

    while time.time() < deadline:
        computer = shell.NameSpace(0x11)  # CSIDL_DRIVES / This PC
        if computer is None:
            raise RuntimeError("Could not open This PC in Shell.")

        device_names: list[str] = []
        matched_dev = None
        for i in range(computer.Items().Count):
            dev = computer.Items().Item(i)
            if not dev:
                continue
            name = str(dev.Name)
            device_names.append(name)
            nlow = name.lower()
            if (
                name == device_name
                or nlow == device_key
                or device_key in nlow
                or ("iphone" in device_key and "iphone" in nlow)
            ):
                matched_dev = dev
                break

        if matched_dev is None:
            last_diag = f"no iPhone under This PC; seen {device_names!r}"
        else:
            try:
                device_folder = matched_dev.GetFolder
                sub_names: list[str] = []
                for j in range(device_folder.Items().Count):
                    st = device_folder.Items().Item(j)
                    if not st:
                        continue
                    sub = str(st.Name)
                    sub_names.append(sub)
                    if sub.strip().lower() == storage_key:
                        folder = st.GetFolder
                        if folder is not None:
                            return folder
                last_diag = (
                    f"found {matched_dev.Name!r} but not {storage_name!r}; "
                    f"volumes={sub_names!r}"
                )
            except Exception as exc:
                last_diag = f"{matched_dev.Name!r} not readable: {exc}"

        remaining = max(0, int(deadline - time.time()))
        log(f"Waiting for {storage_name!r} on {device_name!r} ({remaining}s left) — {last_diag}")
        time.sleep(poll_sec)

    raise RuntimeError(
        f"Device {device_name!r} / {storage_name!r} not available after {wait_sec:.0f}s. "
        f"Last status: {last_diag}. "
        "Unlock the iPhone, keep it awake, tap Trust / Allow Full Access to Photos, "
        "then unplug and replug USB (close Apple Devices / Explorer on the phone if open)."
    )


def photo_key(photo: PhotoEntry) -> tuple[str, str]:
    return photo.album, photo.name


@dataclass
class CacheRow:
    album: str
    name: str
    type_label: str
    taken: datetime
    precision: str  # "exif" (seconds) or "shell" (minute)


class ExifYearCache:
    """Persistent CSV cache: one file per calendar year under cache_dir."""

    def __init__(self, cache_dir: Path) -> None:
        self.cache_dir = cache_dir
        self.cache_dir.mkdir(parents=True, exist_ok=True)
        self._by_year: dict[int, dict[tuple[str, str], CacheRow]] = {}
        self._dirty_years: set[int] = set()

    def _path(self, year: int) -> Path:
        return self.cache_dir / EXIF_CACHE_FILENAME.format(year=year)

    def _year_for_album(self, album: str) -> int | None:
        ym = parse_album_year_month(album)
        return ym[0] if ym else None

    def _load_year(self, year: int) -> dict[tuple[str, str], CacheRow]:
        if year in self._by_year:
            return self._by_year[year]
        rows: dict[tuple[str, str], CacheRow] = {}
        path = self._path(year)
        if path.is_file():
            with path.open(newline="", encoding="utf-8") as fh:
                for rec in csv.DictReader(fh):
                    try:
                        taken = datetime.fromisoformat(rec["taken"])
                    except (KeyError, ValueError):
                        continue
                    row = CacheRow(
                        album=rec["album"],
                        name=rec["name"],
                        type_label=rec.get("type_label", ""),
                        taken=taken,
                        precision=rec.get("precision", "exif"),
                    )
                    rows[(row.album, row.name)] = row
        self._by_year[year] = rows
        return rows

    def get(self, album: str, name: str) -> CacheRow | None:
        year = self._year_for_album(album)
        if year is None:
            return None
        return self._load_year(year).get((album, name))

    def has_exif(self, album: str, name: str) -> bool:
        row = self.get(album, name)
        return row is not None and row.precision == "exif"

    def put(
        self,
        album: str,
        name: str,
        type_label: str,
        taken: datetime,
        *,
        precision: str,
    ) -> None:
        year = self._year_for_album(album)
        if year is None:
            return
        bucket = self._load_year(year)
        bucket[(album, name)] = CacheRow(
            album=album,
            name=name,
            type_label=type_label,
            taken=taken,
            precision=precision,
        )
        self._dirty_years.add(year)

    def flush(self) -> None:
        for year in sorted(self._dirty_years):
            path = self._path(year)
            bucket = self._by_year.get(year, {})
            with path.open("w", newline="", encoding="utf-8") as fh:
                writer = csv.DictWriter(fh, fieldnames=EXIF_CACHE_FIELDS)
                writer.writeheader()
                for row in sorted(
                    bucket.values(), key=lambda r: (r.album, r.name)
                ):
                    writer.writerow(
                        {
                            "album": row.album,
                            "name": row.name,
                            "type_label": row.type_label,
                            "taken": row.taken.isoformat(sep=" "),
                            "precision": row.precision,
                        }
                    )
        self._dirty_years.clear()


def keys_needing_exif_refinement(
    photos: list[PhotoEntry],
    threshold_sec: float,
    exif_cache: ExifYearCache | None,
) -> set[tuple[str, str]]:
    """Shell dates are minute-only; copy EXIF only for pairs that might be within threshold."""
    by_month: dict[tuple[int, int], list[PhotoEntry]] = defaultdict(list)
    for photo in photos:
        by_month[photo_year_month(photo)].append(photo)

    need: set[tuple[str, str]] = set()
    # Worst case: 10:03:59 vs 10:04:09 → 10s apart but shell shows 60s between minutes.
    shell_slack = 60.0 + threshold_sec
    for _ym in sorted(by_month.keys()):
        ordered = sorted(
            by_month[_ym], key=lambda p: (p.taken, p.album, p.name)
        )
        for i in range(len(ordered) - 1):
            cur, nxt = ordered[i], ordered[i + 1]
            delta = (nxt.taken - cur.taken).total_seconds()
            if delta <= shell_slack:
                if exif_cache is None or not exif_cache.has_exif(cur.album, cur.name):
                    need.add(photo_key(cur))
                if exif_cache is None or not exif_cache.has_exif(nxt.album, nxt.name):
                    need.add(photo_key(nxt))
    return need


def refine_timestamps_from_exif(
    shell: Any,
    photos: list[PhotoEntry],
    keys: set[tuple[str, str]],
    cache_dir: str,
    album_folders: dict[str, Any],
    columns: dict[str, int],
    exif_cache: ExifYearCache | None,
) -> None:
    total = len(keys)
    log(f"Refining {total} photo(s) with EXIF over MTP (not all {len(photos)}).")
    done = 0
    for photo in photos:
        if photo_key(photo) not in keys:
            continue
        done += 1
        log(f"  EXIF {done}/{total}: {photo.album}/{photo.name} (MTP copy, up to {DEFAULT_COPY_TIMEOUT_SEC:.0f}s)…")
        local = copy_shell_item_to_dir(shell, cache_dir, photo.item)
        if local is None:
            log("    copy failed or timed out, keeping shell date")
            continue
        taken = exif_datetime_from_path(local)
        try:
            local.unlink(missing_ok=True)
        except OSError:
            pass
        if taken is None:
            folder = album_folders.get(photo.album)
            if folder is not None:
                taken = item_datetime_shell(folder, photo.item, columns)
        if taken is not None:
            photo.taken = taken
            if exif_cache is not None:
                exif_cache.put(
                    photo.album,
                    photo.name,
                    photo.type_label,
                    taken,
                    precision="exif",
                )
    if exif_cache is not None:
        exif_cache.flush()


def collect_photos(
    shell: Any,
    storage: Any,
    columns: dict[str, int],
    *,
    album_filter: re.Pattern[str] | None,
    limit: int | None,
    exif_mode: str,
    cache_dir: str,
    threshold_sec: float,
    exif_cache: ExifYearCache | None,
) -> tuple[list[PhotoEntry], dict[tuple[str, str], Any]]:
    images: list[PhotoEntry] = []
    mov_by_key: dict[tuple[str, str], Any] = {}
    album_folders: dict[str, Any] = {}
    albums = list_album_entries(storage)
    total_albums = len(albums)

    for album_idx, (album_name, album_item) in enumerate(albums):
        if album_filter and not album_filter.search(album_name):
            continue
        log(f"Scanning album {album_idx + 1}/{total_albums}: {album_name} …")
        try:
            album_folder = album_item.GetFolder
        except Exception as exc:
            log(f"  skip album (cannot open): {exc}")
            continue
        if album_folder is None:
            log("  skip album (empty folder handle)")
            continue
        album_folders[album_name] = album_folder
        album_images = 0
        try:
            count = album_folder.Items().Count
        except Exception as exc:
            log(f"  skip album (cannot list items): {exc}")
            continue
        for k in range(count):
            if k == 0 or (k + 1) % 25 == 0:
                log(f"  reading item {k + 1}/{count} …")
            item = album_folder.Items().Item(k)
            if not item:
                continue
            type_label = str(item.Type or "")
            name = str(item.Name or "")
            if is_mov_type(type_label):
                mov_by_key[(album_name, name)] = item
                continue
            if not is_image_type(type_label):
                continue
            taken: datetime | None = None
            if exif_cache is not None:
                cached = exif_cache.get(album_name, name)
                if cached is not None:
                    taken = cached.taken
            if taken is None:
                taken = item_datetime_shell(album_folder, item, columns)
                if taken is not None and exif_cache is not None:
                    exif_cache.put(
                        album_name, name, type_label, taken, precision="shell"
                    )
            if taken is None:
                log(f"  skip (no date): {album_name}/{name} [{type_label}]")
                continue
            album_images += 1
            images.append(
                PhotoEntry(
                    album=album_name,
                    name=name,
                    type_label=type_label,
                    taken=taken,
                    item=item,
                )
            )
            if exif_mode == "all" and (exif_cache is None or not exif_cache.has_exif(album_name, name)):
                log(f"  EXIF copy {album_name}/{name} …")
                local = copy_shell_item_to_dir(shell, cache_dir, item)
                if local is None:
                    log(f"    skip (copy failed): {album_name}/{name}")
                else:
                    exif_t = exif_datetime_from_path(local)
                    try:
                        local.unlink(missing_ok=True)
                    except OSError:
                        pass
                    if exif_t is not None:
                        images[-1].taken = exif_t
                        if exif_cache is not None:
                            exif_cache.put(
                                album_name,
                                name,
                                type_label,
                                exif_t,
                                precision="exif",
                            )
            if limit is not None and len(images) >= limit:
                if exif_mode == "smart":
                    refine_timestamps_from_exif(
                        shell,
                        images,
                        keys_needing_exif_refinement(
                            images, threshold_sec, exif_cache
                        ),
                        cache_dir,
                        album_folders,
                        columns,
                        exif_cache,
                    )
                if exif_cache is not None:
                    exif_cache.flush()
                return images, mov_by_key
        if album_images:
            log(f"  {album_images} image(s) indexed (shell dates).")

    if exif_mode == "smart":
        refine_timestamps_from_exif(
            shell,
            images,
            keys_needing_exif_refinement(images, threshold_sec, exif_cache),
            cache_dir,
            album_folders,
            columns,
            exif_cache,
        )
    if exif_cache is not None:
        exif_cache.flush()
    return images, mov_by_key


def choose_deletions(
    photos: list[PhotoEntry], threshold_sec: float
) -> list[PhotoEntry]:
    """Only compare consecutive photos within the same calendar month (from album YYYYMM)."""
    by_month: dict[tuple[int, int], list[PhotoEntry]] = defaultdict(list)
    for photo in photos:
        by_month[photo_year_month(photo)].append(photo)

    victims: list[PhotoEntry] = []
    for _ym in sorted(by_month.keys()):
        ordered = sorted(
            by_month[_ym], key=lambda p: (p.taken, p.album, p.name)
        )
        for i in range(len(ordered) - 1):
            cur = ordered[i]
            nxt = ordered[i + 1]
            delta = (nxt.taken - cur.taken).total_seconds()
            if 0 <= delta <= threshold_sec:
                victims.append(cur)
    return victims


FILEOP_FLAGS = shellcon.FOF_NO_UI


def index_album_items(
    album_folder: Any,
) -> tuple[dict[str, Any], dict[str, Any]]:
    """One MTP pass per album: name -> FolderItem for images and MOVs."""
    images: dict[str, Any] = {}
    movs: dict[str, Any] = {}
    try:
        count = album_folder.Items().Count
    except Exception:
        return images, movs
    for k in range(count):
        item = album_folder.Items().Item(k)
        if not item:
            continue
        name = str(item.Name)
        type_label = str(item.Type or "")
        if is_mov_type(type_label):
            movs[name] = item
        elif is_image_type(type_label):
            images[name] = item
    return images, movs


def delete_items_via_fileop(items: list[Any]) -> None:
    if not items:
        return
    fo = pythoncom.CoCreateInstance(
        shell.CLSID_FileOperation,
        None,
        pythoncom.CLSCTX_INPROC_SERVER,
        shell.IID_IFileOperation,
    )
    fo.SetOperationFlags(FILEOP_FLAGS)
    for item in items:
        pidl = shell.SHGetIDListFromObject(item)
        psi = shell.SHCreateItemFromIDList(pidl, shell.IID_IShellItem)
        fo.DeleteItem(psi, None)
    fo.PerformOperations()
    if fo.GetAnyOperationsAborted():
        raise OSError("IFileOperation aborted")


def delete_item_via_invoke(item: Any) -> None:
    """Fast on MTP but may show a confirmation dialog per file."""
    item.InvokeVerb("delete")


@dataclass(frozen=True)
class DeleteTarget:
    album: str
    name: str
    is_mov: bool


def index_album_folders(storage: Any) -> dict[str, Any]:
    folders: dict[str, Any] = {}
    for i in range(storage.Items().Count):
        album_item = storage.Items().Item(i)
        if not album_item:
            continue
        try:
            folder = album_item.GetFolder
        except Exception:
            continue
        if folder is not None:
            folders[str(album_item.Name)] = folder
    return folders


def _delete_album_batch(
    album_folder: Any,
    album_name: str,
    targets: list[DeleteTarget],
    use_invoke_verb: bool,
) -> tuple[int, int]:
    images, movs = index_album_items(album_folder)
    items_to_delete: list[Any] = []
    missing = 0
    for target in targets:
        item = movs.get(target.name) if target.is_mov else images.get(target.name)
        if item is None:
            missing += 1
        else:
            items_to_delete.append(item)

    if not items_to_delete:
        return 0, missing

    if use_invoke_verb:
        ok = 0
        failed = missing
        for item in items_to_delete:
            try:
                delete_item_via_invoke(item)
                ok += 1
            except Exception:
                failed += 1
        return ok, failed

    try:
        delete_items_via_fileop(items_to_delete)
        return len(items_to_delete), missing
    except Exception:
        ok = 0
        failed = missing
        for item in items_to_delete:
            try:
                delete_items_via_fileop([item])
                ok += 1
            except Exception:
                failed += 1
        return ok, failed


def _delete_album_worker(
    device: str,
    storage_name: str,
    connect_wait: float,
    album_name: str,
    targets: list[DeleteTarget],
    use_invoke_verb: bool,
) -> tuple[int, int]:
    pythoncom.CoInitialize()
    try:
        sh = win32com.client.Dispatch("Shell.Application")
        storage = find_device_storage(
            sh, device, storage_name, wait_sec=min(connect_wait, 15.0)
        )
        album_folder = index_album_folders(storage).get(album_name)
        if album_folder is None:
            return 0, len(targets)
        return _delete_album_batch(
            album_folder, album_name, targets, use_invoke_verb
        )
    finally:
        pythoncom.CoUninitialize()


def delete_targets_on_device(
    storage: Any,
    targets: list[DeleteTarget],
    *,
    use_invoke_verb: bool,
    parallel: int,
    device: str,
    storage_name: str,
    connect_wait: float,
) -> tuple[int, int]:
    """Delete by album; optional parallel workers (separate MTP session each)."""
    by_album: dict[str, list[DeleteTarget]] = defaultdict(list)
    for target in targets:
        by_album[target.album].append(target)

    ok = 0
    failed = 0
    total_albums = len(by_album)

    if parallel > 1 and total_albums > 1:
        log(f"Deleting across {total_albums} album(s) with {parallel} parallel worker(s)…")
        with ThreadPoolExecutor(max_workers=parallel) as pool:
            futures = {
                pool.submit(
                    _delete_album_worker,
                    device,
                    storage_name,
                    connect_wait,
                    album_name,
                    album_targets,
                    use_invoke_verb,
                ): album_name
                for album_name, album_targets in by_album.items()
            }
            done_albums = 0
            for fut in as_completed(futures):
                album_name = futures[fut]
                done_albums += 1
                try:
                    aok, afail = fut.result()
                    ok += aok
                    failed += afail
                    log(
                        f"  [{done_albums}/{total_albums}] {album_name}: "
                        f"deleted {aok}, failed {afail}"
                    )
                except Exception as exc:
                    n = len(by_album[album_name])
                    failed += n
                    log(f"  [{done_albums}/{total_albums}] {album_name} ERROR: {exc}")
        return ok, failed

    album_folders = index_album_folders(storage)
    for album_idx, (album_name, album_targets) in enumerate(
        sorted(by_album.items()), start=1
    ):
        album_folder = album_folders.get(album_name)
        if album_folder is None:
            n = len(album_targets)
            failed += n
            log(f"  FAILED missing album {album_name!r} ({n} file(s))")
            continue
        log(
            f"  [{album_idx}/{total_albums}] deleting {len(album_targets)} "
            f"target(s) in {album_name} …"
        )
        aok, afail = _delete_album_batch(
            album_folder, album_name, album_targets, use_invoke_verb
        )
        ok += aok
        failed += afail
        log(f"  [{album_idx}/{total_albums}] {album_name}: deleted {aok}, failed {afail}")

    return ok, failed


def main() -> int:
    parser = argparse.ArgumentParser(
        description="MTP burst cleanup on a connected iPhone (Windows Shell)."
    )
    parser.add_argument(
        "--device", default="Apple iPhone", help="Name under This PC"
    )
    parser.add_argument(
        "--storage", default="Internal Storage", help="Storage volume name"
    )
    parser.add_argument(
        "--seconds",
        type=float,
        default=10.0,
        help="Max gap to next photo to delete the earlier one",
    )
    parser.add_argument(
        "--album-regex",
        default="",
        help="Only scan albums whose name matches this regex",
    )
    parser.add_argument(
        "--limit",
        type=int,
        default=None,
        help="Stop after this many photos (not albums/months)",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="List files that would be deleted (default)",
    )
    parser.add_argument(
        "--execute",
        action="store_true",
        help="Delete on the iPhone over MTP",
    )
    parser.add_argument(
        "--no-live-photo-mov",
        action="store_true",
        help="Do not delete .MOV paired with the same name in the same album",
    )
    parser.add_argument(
        "--connect-wait",
        type=float,
        default=45.0,
        help="Seconds to retry until Internal Storage appears (default 45)",
    )
    parser.add_argument(
        "--fast-shell-dates",
        action="store_true",
        help="Never copy for EXIF (minute precision only — misses many 10s bursts)",
    )
    parser.add_argument(
        "--full-exif",
        action="store_true",
        help="Copy every photo over MTP for EXIF (slowest, most accurate)",
    )
    parser.add_argument(
        "--invoke-verb-delete",
        action="store_true",
        help="Fast delete via Explorer verb (you may get Yes/No per file)",
    )
    parser.add_argument(
        "--exif-cache-dir",
        type=Path,
        default=SCRIPT_DIR,
        help=f"Directory for exif_cache_YEAR.csv files (default: {SCRIPT_DIR})",
    )
    parser.add_argument(
        "--no-exif-cache",
        action="store_true",
        help="Do not read or write EXIF CSV caches",
    )
    parser.add_argument(
        "--delete-parallel",
        type=int,
        default=2,
        metavar="N",
        help="Delete up to N albums in parallel (separate MTP sessions; default 2)",
    )
    args = parser.parse_args()

    if args.delete_parallel < 1:
        parser.error("--delete-parallel must be >= 1")

    if args.fast_shell_dates and args.full_exif:
        parser.error("Use either --fast-shell-dates or --full-exif, not both.")

    if args.full_exif:
        exif_mode = "all"
    elif args.fast_shell_dates:
        exif_mode = "none"
    else:
        exif_mode = "smart"

    if args.execute and args.dry_run:
        parser.error("Use either --dry-run or --execute, not both.")
    if not args.execute:
        args.dry_run = True

    album_filter = re.compile(args.album_regex) if args.album_regex else None

    install_hard_interrupt()
    pythoncom.CoInitialize()
    shell = win32com.client.Dispatch("Shell.Application")
    columns = build_detail_column_map(shell)
    storage = find_device_storage(
        shell,
        args.device,
        args.storage,
        wait_sec=args.connect_wait,
    )

    exif_cache = None if args.no_exif_cache else ExifYearCache(args.exif_cache_dir)
    if exif_cache is not None:
        log(f"EXIF cache directory: {args.exif_cache_dir}")

    cache_dir = tempfile.mkdtemp(prefix="iphone_mtp_exif_")
    try:
        if exif_mode == "all":
            log("Copying every photo over MTP for EXIF. Very slow on large libraries.")
        elif exif_mode == "smart":
            log(
                "Pass 1: shell dates (fast). Pass 2: MTP copy only for likely burst pairs."
            )
        else:
            log("WARNING: --fast-shell-dates uses minute precision only.")
        photos, mov_map = collect_photos(
            shell,
            storage,
            columns,
            album_filter=album_filter,
            limit=args.limit,
            exif_mode=exif_mode,
            cache_dir=cache_dir,
            threshold_sec=args.seconds,
            exif_cache=exif_cache,
        )
    finally:
        try:
            os.rmdir(cache_dir)
        except OSError:
            pass
    print(f"Found {len(photos)} images with timestamps.")

    victims = choose_deletions(photos, args.seconds)
    extras: list[tuple[str, str, Any]] = []
    if not args.no_live_photo_mov:
        for v in victims:
            mov = mov_map.get((v.album, v.name))
            if mov is not None:
                extras.append((v.album, v.name, mov))

    print(
        f"Would delete {len(victims)} image(s)"
        + (f" and {len(extras)} Live Photo .MOV(s)" if extras else "")
        + f" (threshold {args.seconds}s)."
    )

    for v in victims:
        tag = "DELETE" if args.execute else "WOULD DELETE"
        print(
            f"  {tag}  {v.taken:%Y-%m-%d %H:%M:%S}  {v.album}/{v.name}  [{v.type_label}]"
        )
    for album, name, _mov in extras:
        tag = "DELETE" if args.execute else "WOULD DELETE"
        print(f"  {tag}  (Live Photo MOV)  {album}/{name}.MOV")

    if args.dry_run:
        print("\nDry run only. Re-run with --execute to delete on the phone.")
        return 0

    if not victims and not extras:
        print("Nothing to delete.")
        return 0

    log("\nDeleting in 5 seconds. Ctrl+C to abort…")
    time.sleep(5)

    storage = find_device_storage(
        shell,
        args.device,
        args.storage,
        wait_sec=args.connect_wait,
    )
    targets: list[DeleteTarget] = [
        DeleteTarget(v.album, v.name, False) for v in victims
    ]
    targets.extend(
        DeleteTarget(album, name, True) for album, name, _mov in extras
    )
    ok, failed = delete_targets_on_device(
        storage,
        targets,
        use_invoke_verb=args.invoke_verb_delete,
        parallel=args.delete_parallel,
        device=args.device,
        storage_name=args.storage,
        connect_wait=args.connect_wait,
    )
    if ok == 0 and failed:
        log(f"Delete failed for all {failed} item(s).")
        return 1
    log(
        f"Deleted {ok} item(s)"
        + (f", {failed} failed" if failed else "")
        + ". Check the phone / Recently Deleted."
    )
    return 0 if failed == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
