"""Met à jour la table des matières d'un .docx avec LibreOffice (UNO) et réenregistre en .docx.

Usage : python3 maj_sommaire.py entree.docx sortie.docx
"""
import os
import subprocess
import sys
import time

import uno
from com.sun.star.beans import PropertyValue


def prop(name, value):
    p = PropertyValue()
    p.Name, p.Value = name, value
    return p


def main(src, dst):
    port = 2002
    office = subprocess.Popen([
        "soffice", "-env:UserInstallation=file:///tmp/lo_profile_uno", "--headless", "--invisible",
        "--norestore", f"--accept=socket,host=localhost,port={port};urp;",
    ])
    try:
        ctx_local = uno.getComponentContext()
        resolver = ctx_local.ServiceManager.createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver", ctx_local)
        for _ in range(60):
            try:
                ctx = resolver.resolve(f"uno:socket,host=localhost,port={port};urp;StarOffice.ComponentContext")
                break
            except Exception:
                time.sleep(0.5)
        desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
        doc = desktop.loadComponentFromURL(uno.systemPathToFileUrl(os.path.abspath(src)), "_blank", 0,
                                           (prop("Hidden", True),))
        # remplace le repère [[SOMMAIRE]] par une table des matières générée à partir des titres
        search = doc.createSearchDescriptor()
        search.SearchString = "[[SOMMAIRE]]"
        found = doc.findFirst(search)
        if found is not None:
            index = doc.createInstance("com.sun.star.text.ContentIndex")
            index.CreateFromOutline = True
            index.Level = 2
            index.Title = ""
            found.setString("")
            doc.getText().insertTextContent(found, index, True)
        # table des matières compacte (tient sur une page)
        styles = doc.getStyleFamilies().getByName("ParagraphStyles")
        for name in ("Contents 1", "Contents 2"):
            if styles.hasByName(name):
                st = styles.getByName(name)
                st.CharHeight = 11
                st.CharFontName = "Times New Roman"
                st.ParaTopMargin = 0
                st.ParaBottomMargin = 60 if name == "Contents 1" else 0
                ls = st.ParaLineSpacing
                ls.Mode, ls.Height = 0, 100
                st.ParaLineSpacing = ls
                if name == "Contents 1":
                    st.CharWeight = 150.0
        indexes = doc.getDocumentIndexes()
        for i in range(indexes.getCount()):
            indexes.getByIndex(i).update()
        doc.storeToURL(uno.systemPathToFileUrl(os.path.abspath(dst)),
                       (prop("FilterName", "MS Word 2007 XML"),))
        doc.close(True)
    finally:
        office.terminate()
        office.wait(timeout=30)


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
