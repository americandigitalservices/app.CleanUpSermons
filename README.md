# ProPresenter Sermon Outline Cleaner

Converts sermon outlines (`.docx`, `.pdf`, `.eml`, or `.txt`) into clean `.txt` files on your `~/Desktop` for ProPresenter lower thirds:
* `MainSpeaker.txt` (speaker name, no trailing line break)
* `SermonTitle.txt` (extracted from document header, email subject, or filename, no trailing line break)
* `SermonText.txt` (smart-wrapped at 105 characters per line across both scriptures and commentary, with standardized `Book Ch:V (TRANS) - ` prefixes)
* `SermonScriptures.txt` (ordered list of canonical scripture references and translations)

---

## Key Features

* **Multi-Format & Apple Mail (`.eml`) Support:** Reads Word (`.docx`), PDF (`.pdf`), plain text (`.txt`), and Apple Mail (`.eml`) files, automatically extracting sender names (`From:`) and subjects while unwrapping hard-wrapped email lines.
* **Canonical Scripture & Translation Normalization:** Expands book abbreviations (`Luk` -> `Luke`, `II Cor` -> `2 Corinthians`, `Philip.` -> `Philippians`) and full translation names (`Amplified Bible, Classic Edition` -> `AMPC`, `New Living Translation` / `NLT2` -> `NLT`, `GOD'S WORD Translation` -> `GW`) into a unified `Book Ch:V (TRANS) - ` prefix on every line.
* **Smart Lower-Third Line Wrapping:** Uses bracket-aware, clause-balanced wrapping (`MAX_LINE_CHARS = 105`) for both scripture verses and commentary notes—keeping `(...)` and `[...]` amplifications intact and preventing orphaned words at the ends of lines.
* **Typography & Study Note Cleanup:** Replaces em/en-dashes with standard hyphens (` - `), converts smart quotes, strips BibleGateway footnote markers (`[a]`, `[b]`), separates multi-part Hebrew/transliteration study lines, and splits combined `Topic: Reference` headers onto separate lines.

---

## Option 1: Run via AppleScript (macOS)

Keep `Launcher.applescript` (or your exported `.app`), `clean_sermon.py`, and `requirements.txt` in the **same folder**. On first run, the launcher automatically uses Homebrew `python3` to create the `venv` and install `requirements.txt`.

### How to Use
* **Double-click** the app (or click **Run** in Script Editor) -> Select your outline file (`.docx`, `.pdf`, `.eml`, or `.txt`) -> Confirm the Speaker Name -> View the summary dialog.
* **Or Drag-and-Drop** your outline file directly onto the app icon -> Confirm the Speaker Name -> View the summary dialog.

### Exporting as a macOS App & Signing Notes
1. Open `Launcher.applescript` in **Script Editor** and choose **File -> Export...**
2. Set **File Format** to **Application** and save it in the **same folder** as `clean_sermon.py`.
3. **Run-Only & Code Signing Options:**
   * **Export as Run-Only + Signed (Recommended for Production/Sharing):** Check **Run-only** and select your Developer certificate under **Code Sign**. This locks the script source and prevents macOS Gatekeeper warnings or repeated permission prompts on every run.
   * **Sign to Run Locally (Ad-Hoc Signing):** Select **Sign to Run Locally** in the Code Sign dropdown (or run `codesign --force --deep --sign - "AppName.app"` in Terminal). This allows the app to run smoothly on the current Mac without an Apple Developer account.

---

## Option 2: Manual Setup & CLI Run (Mac / PC)

### 1. Setup Virtual Environment (`venv`) & Requirements

Open your terminal in the project folder:

* **macOS / Linux:**
```bash
  python3 -m venv venv
  source venv/bin/activate
  pip install -r requirements.txt

```

* **Windows (Command Prompt / PowerShell):**
```cmd
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt

```



### 2. Run via CLI

* **Interactive Speaker Prompt:**
```bash
python clean_sermon.py "Outline.docx"

```


*(Press **Enter** to accept the default `Pastor Ernest Fink` or detected `.eml` sender, or type a custom name.)*
* **Pass Speaker Name Directly:**
```bash
python clean_sermon.py "Outline.eml" "Pastor John Doe"

```



---

## Configuration

To change the character limit, default speaker, or output folder, edit the top of `clean_sermon.py`:

```python
MAX_LINE_CHARS = 105
DEFAULT_SPEAKER = "Pastor Ernest Fink"
OUTPUT_DIR = Path("~/Desktop").expanduser()

```
