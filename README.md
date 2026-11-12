# unlockppt

Got a PPTX that opens with a password but is **read-only** (no editing, no Save As, no PDF export)? This tool removes those locks and exports a PDF, so you can study from it.

> ⚠️ **You must already know the password.** This tool does not crack, guess, or bypass anything. It only decrypts files you have legitimate access to. Use it only on files you're allowed to have a personal copy of.

## What it does

- Decrypts the password-protected PPTX (AES) using the password you provide
- Strips "read-only recommended", "password to modify", and "Mark as Final" flags
- Exports to PDF (PowerPoint on Windows, LibreOffice on macOS/Linux)
- Works on a single file or a whole folder

## Install

```bash
git clone https://github.com/abdulrehman1610/unlockppt.git
cd unlockppt
pip install -r requirements.txt
```

**For PDF export you need one of:**
- Microsoft PowerPoint (Windows), or
- [LibreOffice](https://www.libreoffice.org/) (Windows/macOS/Linux)

If neither is installed, the unlocked PPTX is still created and you can use Save As PDF manually.

## Usage

```bash
# Single file (prompts for password)
python unlockppt.py "Lecture W2.pptx"

# Whole folder, password inline
python unlockppt.py lectures/ -p yourpassword

# Unlock only, no PDF
python unlockppt.py "Lecture W2.pptx" --no-pdf

# Custom output folder
python unlockppt.py lectures/ -o my_notes
```

| Option | Description |
|---|---|
| `inputs` | One or more PPTX files or folders |
| `-p, --password` | Password (omit to be prompted securely) |
| `-o, --out` | Output folder (default: `unlocked`) |
| `--no-pdf` | Skip PDF export |

Output lands in the `unlocked/` folder.

## Optional: Markdown export

```bash
pip install markitdown
markitdown unlocked/lecture.pptx > lecture.md
```

## Limitations

- **Wrong password:** the file won't decrypt.
- **IRM / "Restrict Access" files** (tied to a Microsoft account) can't be unlocked by any script. Ask the file owner for an unrestricted copy.
- PDF fidelity is best with PowerPoint; LibreOffice may shift fonts or animations slightly.

## Privacy

Everything runs locally on your machine. Nothing is uploaded anywhere.

## License

MIT
