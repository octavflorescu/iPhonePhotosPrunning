# iPhone burst photo cleanup (Windows + USB)

Removes **earlier** photos in a burst when the **next** photo in the **same calendar month** was taken within N seconds (default 10). Works over USB using Windows MTP (Shell COM)—there is no drive letter like `D:\`.

Album folders on the phone look like `201805_a` → May **2018**. Photos are never compared across months.

---

## Requirements

- **Windows** PC
- **Python 3.10+** (3.11 or 3.12 recommended) — [Download Python for Windows](https://www.python.org/downloads/windows/) (run the installer and enable **“Add python.exe to PATH”**)
- iPhone connected by USB, unlocked, trusted
- **Full photo access** when iOS asks (not “Limited Access” only)

Verify Python after install:

```powershell
python --version
pip --version
```

---

## 1. One-time setup

Open PowerShell in the folder that contains `iphone_mtp_burst_cleanup.py` (for example: **Shift + right-click** the folder → **Open in terminal**, or `cd` to that folder yourself).

```powershell
cd path\to\iPhonePhotosPrunning
```

Create a virtual environment (optional but recommended):

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

Install dependencies:

```powershell
pip install -r requirements.txt
```

---

## 2. Prepare the iPhone

1. Unlock the iPhone and keep the screen on during long runs.
2. Connect USB; if prompted, tap **Trust** on the phone and confirm on the PC.
3. On iOS 18+, choose **Allow Full Access** to photos if asked.
4. In **File Explorer** → **This PC** → **Apple iPhone** → **Internal Storage**, confirm you see DCIM-style folders (`201805_a`, …).  
   If **Internal Storage** is empty, fix that in Explorer before running the script (`--connect-wait 90` helps after replugging).
5. Close Explorer windows that are browsing the iPhone (they can lock MTP).

---

## 3. Dry run (always do this first)

Default mode lists what would be deleted; nothing is removed.

Run from the folder that contains the script:

```powershell
python -u iphone_mtp_burst_cleanup.py --dry-run
```

If you run from a parent directory instead, point at the script:

```powershell
python -u path\to\iPhonePhotosPrunning\iphone_mtp_burst_cleanup.py --dry-run
```

**Try one year only** (album names start with `2018`):

```powershell
python -u iphone_mtp_burst_cleanup.py --dry-run --album-regex "^2018"
```

**Cap how many photos are scanned** (for testing—not “months”, actual photo count):

```powershell
python -u iphone_mtp_burst_cleanup.py --dry-run --limit 200
```

Review the `WOULD DELETE` lines. When the list looks right, continue to step 4.

---

## 4. Delete on the phone

```powershell
python -u iphone_mtp_burst_cleanup.py --execute --album-regex "^2018"
```

You get a **5 second countdown**; Ctrl+C aborts before deletes start. If Ctrl+C does not respond during an MTP copy, end **python.exe** in Task Manager.

**Faster silent deletes** (default: 2 albums at a time):

```powershell
python -u iphone_mtp_burst_cleanup.py --execute --album-regex "^2018" --delete-parallel 3
```

**Fastest deletes, but you may get a Yes/No dialog per file:**

```powershell
python -u iphone_mtp_burst_cleanup.py --execute --album-regex "^2018" --invoke-verb-delete
```

---

## 5. EXIF cache (speed on repeat runs)

The script writes CSV files here by default:

- `exif_cache_2018.csv`
- `exif_cache_2019.csv`
- …

Each row stores `album`, `name`, `taken`, and `precision` (`exif` = second-accurate, `shell` = minute-only). After the first full pass, later runs skip MTP copies for photos already cached with `exif`.

Custom cache directory:

```powershell
python -u iphone_mtp_burst_cleanup.py --dry-run --exif-cache-dir .\my_exif_cache
```

Disable cache:

```powershell
python -u iphone_mtp_burst_cleanup.py --dry-run --no-exif-cache
```

---

## 6. Useful options

| Option | Purpose |
|--------|--------|
| `--seconds 10` | Max gap to the **next** photo to delete the earlier one |
| `--album-regex "^202510"` | Only albums whose folder name matches regex |
| `--limit 500` | Stop after N photos indexed |
| `--connect-wait 90` | Wait up to 90s for Internal Storage to appear |
| `--no-live-photo-mov` | Do not delete paired Live Photo `.MOV` files |
| `--fast-shell-dates` | No EXIF copy (fast, bad for 10s rule) |
| `--full-exif` | EXIF copy for every photo (slow) |

Default timestamp mode is **smart**: shell dates first, EXIF copy only for likely burst pairs.

---

## 7. What to expect

- **Scanning** over USB is slow the first time; cache makes re-runs much faster.
- **Deletes** may not behave exactly like the Photos app (check **Recently Deleted** on the iPhone).
- **Back up** important photos (iCloud or copy in Explorer) before `--execute`.
- MTP + COM must stay **single-threaded** during scan; only **delete** uses `--delete-parallel`.

---

## 8. Example workflow (one year)

```powershell
cd path\to\iPhonePhotosPrunning
.\.venv\Scripts\Activate.ps1

python -u iphone_mtp_burst_cleanup.py --dry-run --album-regex "^2018"
python -u iphone_mtp_burst_cleanup.py --execute --album-regex "^2018" --delete-parallel 2
```

---

## Troubleshooting

| Problem | What to try |
|---------|-------------|
| `Internal Storage not found` | Unlock phone, replug USB, open Internal Storage in Explorer once, `--connect-wait 90` |
| Script “stuck”, Ctrl+C dead | Kill **python.exe**; close Explorer on iPhone; replug |
| Too many false positives in dry run | Use default smart EXIF (not `--fast-shell-dates`); ensure cache is building |
| Deletes fail / aborted | Retry with `--delete-parallel 1`; or `--invoke-verb-delete` |

---

## Files

| File | Role |
|------|------|
| `iphone_mtp_burst_cleanup.py` | Main script |
| `requirements.txt` | Python packages |
| `exif_cache_YYYY.csv` | Created automatically; safe to delete to force rescan |
