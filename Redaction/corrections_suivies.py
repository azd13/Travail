"""Applique des corrections en mode révision (modifications suivies) dans un .docx.

Usage : python3 corrections_suivies.py entree.docx sortie.docx
"""
import copy
import sys
from datetime import datetime, timezone

from docx import Document
from docx.oxml.ns import qn

AUTEUR = "Claude (relecture)"
DATE = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
NBSP = " "

# (début du paragraphe pour le repérer, texte fautif, texte corrigé)
CORRECTIONS = [
    ("Le présent audit", "Le présent audit se basera donc sur ce décalage.", "Le présent audit part de ce décalage."),
    ("Le cadre d", "grille d’évaluation du module", "grille d’audit du module"),
    ("Le cadre d", "Une fuite de l’expertise avec le départ des personnes n'est plus orchestrée : elle est perdue.",
     "Lorsque l’expertise part avec les personnes, elle n'est plus orchestrée : elle est perdue."),
    ("L’analyse est basée", "L’analyse est basée sur la perspective de la fonction de chef de salle à la CTMG.",
     "L’analyse adopte la perspective de l’auteur, chef de salle à la CTMG."),
    ("Un chef de salle voit", "Ne disposant pas de l’accès à ces données, ils ne peuvent pas être mesurés, et nous les présentons comme tels.",
     "Faute d’accès à ces données, nous ne pouvons pas les mesurer, et nous les présentons comme tels."),
    ("Le matériau mobilisé", "Comportements Organisationnels", "Comportements organisationnels"),
    ("Le matériau mobilisé", "Ce choix découle toujours de la contrainte", "Ce choix découle de la contrainte"),
    ("Cette proximité expose", "afin de rester dans le cadre d’un audit adressé à un conseil ayant vocation à décrire des mécanismes de l’organisation",
     "car un audit adressé à un conseil a vocation à décrire des mécanismes de l’organisation"),
    ("La CTMG est une ligne", "Police", "police"),
    ("Depuis le 1er juillet 2024", "la CTMG intègre comme secteur à part entière du Département",
     "la CTMG est intégrée, comme secteur à part entière, au Département"),
    ("La CTMG compte", "pendant jusqu'à trois mois (E2)", "pendant une période pouvant aller jusqu'à trois mois (E2),"),
    ("L'enjeu principal", "24/7", "24 heures sur 24 et 7 jours sur 7"),
    ("L'enjeu principal", "repose notamment par la qualité", "repose notamment sur la qualité"),
    ("L'enjeu principal", "de santé publique de régulateurs", "de santé publique des régulateurs"),
    ("L'enjeu principal", "outils de supports", "outils de support"),
    ("L'enjeu principal", "afin d’analyser et orienter les appelants caractérisées par une grande variété de situations et problématiques",
     "afin d’analyser et d’orienter des appels caractérisés par une grande variété de situations et de problématiques"),
    ("Une gouvernance à distance", " qui peut être assimilé à la", ", ce qui renvoie à la"),
    ("Au sein de la CTMG, une ligne",
     "Une position rendant ces derniers transparents dans la communication ascendante et silencieux dans les prises de décisions.",
     "Cette position les rend transparents dans la communication ascendante et sans voix dans les prises de décision."),
    ("À la CTMG, les canaux", "ainsi que d’observation et gestion managériale limitée", "ainsi que d’observations managériales limitées"),
    ("Les facteurs facilitants", "infirmiers expérimentés et autonomes et formés", "infirmiers expérimentés, autonomes et formés"),
    ("L'expertise des régulateurs", "Ces faits exposent donc un risque", "Ces faits exposent donc l’organisation à un risque"),
    ("Une centrale en sous-effectif", "Ces conditions exposent un risque", "Ces conditions exposent à un risque"),
    ("Ce travail repose", "sur les observations en qualité de chef de salle", "sur nos observations en qualité de chef de salle"),
]


def norm(s):
    return s.replace(NBSP, " ").replace(" ", " ")


def make_run(template, text, deleted=False):
    r = copy.deepcopy(template)
    # ne garde que la mise en forme (w:rPr) : évite de dupliquer sauts de ligne, tabulations, etc.
    for child in list(r):
        if child.tag != qn("w:rPr"):
            r.remove(child)
    el = r.makeelement(qn("w:delText" if deleted else "w:t"), {})
    el.text = text
    el.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    r.append(el)
    return r


def wrap(tag, run, rid):
    w = run.makeelement(qn(tag), {qn("w:id"): str(rid), qn("w:author"): AUTEUR, qn("w:date"): DATE})
    w.append(run)
    return w


def track_replace(paragraph, old, new, rid):
    for run in paragraph.runs:
        txt = run.text
        i = norm(txt).find(norm(old))
        if i < 0:
            continue
        r = run._r
        before, target, after = txt[:i], txt[i:i + len(old)], txt[i + len(old):]
        parent = r.getparent()
        pos = parent.index(r)
        items = []
        if before:
            items.append(make_run(r, before))
        items.append(wrap("w:del", make_run(r, target, deleted=True), rid))
        items.append(wrap("w:ins", make_run(r, new), rid + 1))
        if after:
            items.append(make_run(r, after))
        parent.remove(r)
        for k, it in enumerate(items):
            parent.insert(pos + k, it)
        return True
    return False


def main(src, dst):
    doc = Document(src)
    rid = 9000
    ok, ko = 0, []
    for start, old, new in CORRECTIONS:
        cibles = [p for p in doc.paragraphs if norm(p.text).lstrip().startswith(norm(start)) and norm(old) in norm(p.text)]
        if cibles and track_replace(cibles[0], old, new, rid):
            ok += 1
            rid += 2
        else:
            ko.append(old[:60])
    doc.save(dst)
    print(ok, "corrections suivies appliquées")
    for k in ko:
        print("NON APPLIQUÉE :", k)


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
