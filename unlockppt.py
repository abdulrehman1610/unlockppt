#!/usr/bin/env python3
"""unlockppt: remove open/modify locks from a PPTX (password required) and export to PDF."""
import argparse, getpass, io, os, re, shutil, subprocess, sys, tempfile, zipfile
from pathlib import Path
import msoffcrypto


def unlock(src: Path, dst: Path, password: str):
    buf = io.BytesIO()
    with open(src, "rb") as f:
        of = msoffcrypto.OfficeFile(f)
        if of.is_encrypted():
            of.load_key(password=password)
            of.decrypt(buf)
        else:
            buf.write(f.read())
    buf.seek(0)

    with zipfile.ZipFile(buf) as zin, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "ppt/presentation.xml":
                t = data.decode("utf-8")
                t = re.sub(r'<p:modifyVerifier[^>]*?(?:/>|></p:modifyVerifier>)', '', t)
                t = re.sub(r'\s*readOnlyRecommended="[^"]*"', '', t)
                data = t.encode("utf-8")
            elif item.filename == "docProps/custom.xml":
                t = data.decode("utf-8")
                t = re.sub(r'<property[^>]*name="_MarkAsFinal"[^>]*>.*?</property>', '', t, flags=re.DOTALL)
                data = t.encode("utf-8")
            zout.writestr(item, data)


def to_pdf(pptx: Path, out_dir: Path) -> bool:
    pdf = out_dir / (pptx.stem + ".pdf")
    if os.name == "nt":
        try:
            import win32com.client
            app = win32com.client.Dispatch("PowerPoint.Application")
            try:
                deck = app.Presentations.Open(str(pptx.resolve()), WithWindow=False)
                deck.SaveAs(str(pdf.resolve()), 32)
                deck.Close()
                return True
            finally:
                app.Quit()
        except Exception as e:
            print(f"  PowerPoint failed ({e}), trying LibreOffice...")
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        print("  No PowerPoint/LibreOffice found, PDF skipped.")
        return False
    subprocess.run([soffice, "--headless", "--convert-to", "pdf", str(pptx), "--outdir", str(out_dir)],
                   check=True, stdout=subprocess.DEVNULL)
    return True


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("inputs", nargs="+", help="PPTX files or folders")
    ap.add_argument("-p", "--password", help="Password (omit to be prompted)")
    ap.add_argument("-o", "--out", default="unlocked", help="Output folder")
    ap.add_argument("--no-pdf", action="store_true", help="Only unlock, skip PDF")
    args = ap.parse_args()

    files = []
    for i in args.inputs:
        p = Path(i)
        files += sorted(p.glob("*.pptx")) if p.is_dir() else [p]

    password = args.password or getpass.getpass("Password: ")
    out = Path(args.out); out.mkdir(exist_ok=True)

    for f in files:
        print(f"[*] {f.name}")
        try:
            dst = out / f.name
            unlock(f, dst, password)
            if not args.no_pdf:
                to_pdf(dst, out)
            print("  done")
        except Exception as e:
            print(f"  [!] failed: {e}")


if __name__ == "__main__":
    main()
