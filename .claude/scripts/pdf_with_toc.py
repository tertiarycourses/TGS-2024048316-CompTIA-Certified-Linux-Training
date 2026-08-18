#!/usr/bin/env python3
"""pdf_with_toc.py — convert a DOCX to PDF with the TOC/index fields REFRESHED.

Plain `soffice --convert-to pdf` exports the stale TOC field text ("Right-click
and choose Update Field…"). This script drives LibreOffice over UNO instead:
load the DOCX headless, refresh every text field and document index (the TOC),
then export to PDF.

Usage:  /Applications/LibreOffice.app/Contents/Resources/python pdf_with_toc.py <in.docx> [out.pdf]

Run with LibreOffice's BUNDLED python (it ships the `uno` module). The script
starts its own soffice listener on a private pipe, so no other LibreOffice
instance may be running.
"""
import os
import subprocess
import sys
import time

SOFFICE = "/Applications/LibreOffice.app/Contents/MacOS/soffice"
PIPE = "pdf_with_toc"


def url(path):
    from urllib.parse import quote
    return "file://" + quote(os.path.abspath(path))


def main():
    src = os.path.abspath(sys.argv[1])
    dst = os.path.abspath(sys.argv[2]) if len(sys.argv) > 2 else os.path.splitext(src)[0] + ".pdf"

    proc = subprocess.Popen([
        SOFFICE, "--headless", "--invisible", "--norestore", "--nologo",
        f"--accept=pipe,name={PIPE};urp;StarOffice.ComponentContext"])
    try:
        import uno
        from com.sun.star.beans import PropertyValue

        ctx = None
        resolver_ctx = uno.getComponentContext()
        resolver = resolver_ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver", resolver_ctx)
        for _ in range(60):
            try:
                ctx = resolver.resolve(
                    f"uno:pipe,name={PIPE};urp;StarOffice.ComponentContext")
                break
            except Exception:
                time.sleep(0.5)
        if ctx is None:
            raise SystemExit("could not connect to soffice over UNO")

        smgr = ctx.ServiceManager
        desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)

        def prop(name, value):
            p = PropertyValue(); p.Name = name; p.Value = value; return p

        doc = desktop.loadComponentFromURL(url(src), "_blank", 0, (prop("Hidden", True),))
        try:
            doc.refresh()
            fields = doc.getTextFields()
            fields.refresh()
            indexes = doc.getDocumentIndexes()
            for i in range(indexes.getCount()):
                indexes.getByIndex(i).update()
            doc.refresh()
            doc.storeToURL(url(dst), (prop("FilterName", "writer_pdf_Export"),))
        finally:
            doc.close(False)
        print("wrote", dst)
    finally:
        proc.terminate()
        proc.wait()


if __name__ == "__main__":
    main()
