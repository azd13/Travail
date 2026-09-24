"""Convertit les parties rédigées (Markdown simple) en document Word mis en forme."""
import re
import sys

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

FONT = "Times New Roman"
AUTEUR = "Azdine ZARGUI-NAJI"
ENTETE = "Orchestration stratégique RH – Zargui-Naji"


def set_font(style, size, bold=None, italic=None, color=None):
    style.font.name = FONT
    style.font.size = Pt(size)
    style.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    if bold is not None:
        style.font.bold = bold
    if italic is not None:
        style.font.italic = italic
    if color is not None:
        style.font.color.rgb = color
    # supprime les polices de thème (sinon Word/LibreOffice affichent une police sans empattement)
    rfonts = style.element.rPr.rFonts
    for att in ("w:asciiTheme", "w:hAnsiTheme", "w:eastAsiaTheme", "w:cstheme"):
        if rfonts.get(qn(att)) is not None:
            del rfonts.attrib[qn(att)]
    rfonts.set(qn("w:ascii"), FONT)
    rfonts.set(qn("w:hAnsi"), FONT)


def add_field(paragraph, instr):
    run = paragraph.add_run()
    for tag, text in (("begin", None), (None, instr), ("separate", None), ("end", None)):
        if tag:
            el = OxmlElement("w:fldChar")
            el.set(qn("w:fldCharType"), tag)
        else:
            el = OxmlElement("w:instrText")
            el.set(qn("xml:space"), "preserve")
            el.text = text
        run._r.append(el)
    return run


NBSP = "\u00a0"


def typo_fr(text):
    """Espaces insécables de la typographie française."""
    text = text.replace("« ", "«" + NBSP).replace(" »", NBSP + "»")
    return re.sub(r" ([:;?!])", NBSP + r"\1", text)


def add_inline(paragraph, text):
    """Gère **gras** et *italique* (non imbriqués au-delà de ***)."""
    text = typo_fr(text)
    for tok in re.split(r"(\*\*[^*]+\*\*|\*[^*]+\*)", text):
        if not tok:
            continue
        if tok.startswith("**"):
            r = paragraph.add_run(tok[2:-2])
            r.bold = True
        elif tok.startswith("*"):
            r = paragraph.add_run(tok[1:-1])
            r.italic = True
        else:
            paragraph.add_run(tok)


def build(md_path, out_path, titre_partie):
    doc = Document()
    sec = doc.sections[0]
    sec.page_height, sec.page_width = Cm(29.7), Cm(21)
    for m in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
        setattr(sec, m, Cm(2.5))

    normal = doc.styles["Normal"]
    set_font(normal, 12)
    pf = normal.paragraph_format
    pf.line_spacing = 1.15
    pf.space_after = Pt(6)
    pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    bleu = RGBColor(0x1F, 0x3A, 0x5F)
    for name, size in (("Heading 1", 14), ("Heading 2", 13), ("Heading 3", 12)):
        st = doc.styles[name]
        set_font(st, size, bold=True, italic=False, color=bleu)
        st.paragraph_format.space_before = Pt(18 if name == "Heading 1" else 12)
        st.paragraph_format.space_after = Pt(6)
        st.paragraph_format.keep_with_next = True
        st.paragraph_format.line_spacing = 1.15
        st.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT

    # --- Page de titre ---
    def centre(text, size, bold=False, before=0, italic=False):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(before)
        r = p.add_run(text)
        r.font.size = Pt(size)
        r.bold = bold
        r.italic = italic
        return p

    centre("CAS Stratégie et organisation – MRHC", 12, before=60)
    centre("Université de Lausanne", 12)
    centre("ORCHESTRATION STRATÉGIQUE RH", 18, bold=True, before=90)
    centre("Audit organisationnel de l'orchestration stratégique RH", 15, before=12)
    centre("La Centrale téléphonique des médecins de garde (CTMG) – Unisanté", 13, italic=True)
    centre(titre_partie, 11, italic=True, before=12)
    centre("Dr Joëlle BÉDAT", 12, before=90)
    centre(AUTEUR, 13, bold=True, before=60)
    centre("Septembre 2026", 12, before=6)

    # --- Nouvelle section avec en-tête et pagination ---
    body = doc.add_section(WD_SECTION.NEW_PAGE)
    body.different_first_page_header_footer = False
    body.header.is_linked_to_previous = False
    body.footer.is_linked_to_previous = False
    hp = body.header.paragraphs[0]
    hp.text = ENTETE
    hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    hp.runs[0].font.size = Pt(9)
    hp.runs[0].italic = True
    fp = body.footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_field(fp, "PAGE").font.size = Pt(10)
    # la page de titre n'a ni en-tête ni numéro
    doc.sections[0].header.is_linked_to_previous = False
    doc.sections[0].footer.is_linked_to_previous = False

    # --- Sommaire ---
    p = doc.add_paragraph()
    r = p.add_run("Table des matières")
    r.bold = True
    r.font.size = Pt(14)
    r.font.color.rgb = bleu
    toc = doc.add_paragraph()
    add_field(toc, 'TOC \\o "1-3" \\h \\z \\u')
    note = doc.add_paragraph()
    rn = note.add_run("(Clic droit sur la table puis « Mettre à jour les champs » pour l'afficher dans Word.)")
    rn.italic = True
    rn.font.size = Pt(9)
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)

    # --- Corps ---
    lines = open(md_path, encoding="utf-8").read().splitlines()
    # Saute le bloc de titre Markdown (déjà sur la page de titre)
    start = next(i for i, l in enumerate(lines) if l.startswith("## Introduction"))
    in_refs = False
    for line in lines[start:]:
        s = line.rstrip()
        if not s or s == "---":
            continue
        if s.startswith("### "):
            doc.add_heading(typo_fr(s[4:]), level=3 if not in_refs else 2)
        elif s.startswith("## "):
            t = s[3:]
            in_refs = t.lower().startswith("références")
            if in_refs:
                doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
            doc.add_heading(typo_fr(t), level=1)
        else:
            p = doc.add_paragraph()
            add_inline(p, s)
            if in_refs:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                p.paragraph_format.left_indent = Cm(1.25)
                p.paragraph_format.first_line_indent = Cm(-1.25)
                p.paragraph_format.line_spacing = 1.15

    doc.save(out_path)


if __name__ == "__main__":
    build(sys.argv[1], sys.argv[2], sys.argv[3] if len(sys.argv) > 3 else "")
