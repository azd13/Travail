"""Schéma : structure de la CTMG jusqu'à Unisanté et orchestration stratégique RH observée."""
import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyArrowPatch, FancyBboxPatch, Rectangle

plt.rcParams["font.family"] = "Liberation Serif"

NAVY = "#1F3A5F"
NAVY_L = "#DCE3EC"
GREY = "#F2F2F2"
GREY_D = "#6B6B6B"
ACCENT = "#B5462C"
ACCENT_L = "#F6E3DD"

fig, ax = plt.subplots(figsize=(7.4, 10.45))
ax.set_xlim(0, 104)
ax.set_ylim(0, 148)
ax.axis("off")


def box(x, y, w, h, title, sub="", fc=GREY, ec=GREY_D, tc="black", ts=8.2, ss=6.6, lw=0.9, style="round,pad=0.4"):
    ax.add_patch(FancyBboxPatch((x, y), w, h, boxstyle=style, fc=fc, ec=ec, lw=lw))
    if sub:
        ax.text(x + w / 2, y + h * 0.66, title, ha="center", va="center", fontsize=ts, weight="bold", color=tc)
        ax.text(x + w / 2, y + h * 0.28, sub, ha="center", va="center", fontsize=ss, color="#222222",
                linespacing=1.15)
    else:
        ax.text(x + w / 2, y + h / 2, title, ha="center", va="center", fontsize=ts, weight="bold", color=tc)


def arrow(p1, p2, color=NAVY, ls="-", lw=1.1, style="-|>", ms=8, rad=0.0):
    ax.add_patch(FancyArrowPatch(p1, p2, arrowstyle=style, mutation_scale=ms, color=color, lw=lw,
                                 linestyle=ls, connectionstyle=f"arc3,rad={rad}", shrinkA=0, shrinkB=0))


def cross(x, y, s=1.6):
    ax.plot([x - s, x + s], [y - s, y + s], color=ACCENT, lw=1.8)
    ax.plot([x - s, x + s], [y + s, y - s], color=ACCENT, lw=1.8)


# ---------- Environnement institutionnel ----------
ax.text(50, 146, "Environnement institutionnel (configuration « administration publique »)", ha="center",
        fontsize=7.4, style="italic", color=GREY_D)
box(3, 134, 27, 8.5, "Cantons mandants", "Vaud et Neuchâtel\n(mandat, financement)", fc="white")
box(36.5, 134, 27, 8.5, "Conseil d'État vaudois", "désigne le Conseil et le DG,\nratifie le plan stratégique",
    fc=NAVY_L, ec=NAVY)
box(70, 134, 27, 8.5, "Unisanté", "sous la surveillance\ndu DSAS et de l'UNIL", fc="white")
# DGS (Direction générale de la santé), rattachée au DSAS
box(37, 123.3, 26, 6.6, "DGS", "Direction générale de la santé (DSAS)\nsiège au Conseil d'Unisanté", fc="white",
    ec=NAVY, tc=NAVY, ts=7.6, ss=5.9)
ax.plot([50, 50], [134, 130.5], color=NAVY, lw=0.9)

# ---------- Unisanté : architecture de type dualiste ----------
ax.add_patch(Rectangle((1.5, 88.5), 97, 29, fill=False, ec=NAVY, lw=1.2, ls=(0, (5, 3))))
ax.text(9, 90.2, "UNISANTÉ – établissement autonome de droit public :\narchitecture de type dualiste",
        fontsize=7.2, weight="bold", color=NAVY, va="bottom")
box(4, 101, 36, 11, "Conseil d'Unisanté", "SURVEILLANCE ET ORIENTATION\nplan stratégique, budget,\nobjectifs annuels, comptes",
    fc="white", ec=NAVY, tc=NAVY)
ax.text(22, 99.2, "dont la directrice générale de la santé (DGS)", ha="center", fontsize=6.1, style="italic",
        color=GREY_D)
box(56, 101, 40, 11, "Direction générale", "EXÉCUTION\ncomité de direction : finances, RH,\nmédical, soins, 7 départements",
    fc="white", ec=NAVY, tc=NAVY)
box(64, 90.5, 24, 5.8, "Direction des RH", "", fc=NAVY_L, ec=NAVY, tc=NAVY, ts=7.4)
arrow((76, 101), (76, 96.8), color=NAVY, lw=0.8, style="-", ms=6)
# séparation surveillance / gestion
ax.plot([48, 48], [99, 104.8], color=NAVY, lw=1.0, ls=(0, (2, 2)))
ax.text(48, 98.4, "séparation surveillance / gestion", ha="center", va="top", fontsize=5.8, color=NAVY,
        style="italic")
ax.annotate("", xy=(56, 106.5), xytext=(40, 106.5),
            arrowprops=dict(arrowstyle="<|-", color=NAVY, lw=0.9, ls="--"))
ax.text(48, 108.3, "comptes rendus\n(information filtrée)", ha="center", fontsize=5.8, color=NAVY)

# double principal
arrow((38, 134), (22, 112.4), color=NAVY, lw=1.1)
arrow((62, 134), (78, 112.4), color=NAVY, lw=1.1)
ax.text(26.5, 124.5, "désigne", fontsize=6.2, color=NAVY, rotation=53)
ax.text(69.5, 126.5, "désigne", fontsize=6.2, color=NAVY, rotation=-53)
ax.text(50, 119.6, "double principal", ha="center", fontsize=6.4, weight="bold", color=ACCENT)

# ---------- DUSC ----------
box(22, 80, 56, 5.8, "DUSC – Département urgences et santé communautaires", "", fc=GREY, ec=GREY_D, ts=7.3)
arrow((59, 101), (56, 86.2), color=NAVY, lw=1.0)

# ---------- CTMG ----------
ax.add_patch(Rectangle((1.5, 14), 97, 62.5, fill=False, ec=GREY_D, lw=1.2))
ax.text(10, 74.2, "CTMG – secteur du DUSC depuis le 01.07.2024", fontsize=7.6, weight="bold", color="#333333")
arrow((50, 80), (50, 71.6), color=NAVY, lw=1.0)
box(34, 65.5, 32, 5.8, "Cheffe de service", "", fc="white", ec=GREY_D, ts=7.6)

ax.text(12, 60.5, "Ligne managériale", ha="left", fontsize=6.8, style="italic", color=GREY_D)
ax.text(88, 60.5, "Ligne médicale", ha="right", fontsize=6.8, style="italic", color=GREY_D)
box(12, 50, 28, 8, "3 cadres", "gestion RH courante déléguée\n« entre la direction et les équipes »",
    fc="white", ec=GREY_D)
box(60, 50, 28, 8, "Médecin responsable", "lien avec les partenaires\nsanitaires", fc="white", ec=GREY_D)
arrow((42, 65.5), (28, 58.4), color=GREY_D, lw=0.9)
arrow((58, 65.5), (72, 58.4), color=GREY_D, lw=0.9)

box(12, 36, 28, 8.5, "7 chefs de salle", "coordination en temps réel,\nrelais informel sans mandat", fc=ACCENT_L,
    ec=ACCENT)
box(60, 36, 28, 8.5, "Médecins formateurs", "qualité, validation clinique\nde l'autonomie", fc="white", ec=GREY_D)
arrow((34, 50), (34, 44.9), color=GREY_D, lw=0.9)
arrow((74, 50), (74, 44.9), color=GREY_D, lw=0.9)
ax.plot([40.6, 59.4], [40.2, 40.2], color=GREY_D, lw=0.8, ls=":")
ax.text(50, 41.2, "intersection des deux lignes", ha="center", fontsize=5.8, style="italic", color=GREY_D)

box(8, 18, 84, 11, "Équipes opérationnelles",
    "33 régulateurs sanitaires (dont 5 praticiens formateurs) · 16 agents de front-office · 2 administratifs\n"
    "expertise clinique largement tacite – autonomie élevée – turnover observé sur le terrain",
    fc="white", ec=GREY_D)
arrow((26, 36), (26, 29.6), color=GREY_D, lw=0.9)

# ---------- Flux ----------
# pilotage descendant par indicateurs
arrow((6.5, 99), (6.5, 31), color=NAVY, lw=2.2, ms=12)
ax.text(4.3, 52, "pilotage descendant par indicateurs de volume", rotation=90, ha="center", va="center",
        fontsize=6.2, color=NAVY, weight="bold")
# information humaine ascendante, filtrée
arrow((100.5, 29.5), (100.5, 98), color=ACCENT, lw=1.6, ls=(0, (4, 3)), ms=11)
ax.text(103.6, 62, "information humaine ascendante, filtrée à chaque niveau", rotation=90, ha="center",
        va="center", fontsize=6.2, color=ACCENT, weight="bold")
cross(100.5, 40.2, 1.2)
cross(100.5, 54, 1.2)
cross(100.5, 82.9, 1.2)
ax.text(96.5, 15.6, "à la base : silence et autocensure", ha="right", va="center", fontsize=5.8, color=ACCENT, style="italic")
# remontée chefs de salle : volumes seulement
ax.annotate("", xy=(18, 50), xytext=(18, 44.9),
            arrowprops=dict(arrowstyle="-|>", color=ACCENT, lw=0.9, ls="--"))
ax.text(19.5, 47.4, "volumes\nseulement", ha="left", va="center", fontsize=5.6, color=ACCENT, style="italic")
# RH à distance
ax.plot([88.6, 93.5, 93.5], [93.4, 93.4, 30.6], color=NAVY, lw=0.9, ls=":")
ax.annotate("", xy=(90, 29.6), xytext=(93.5, 30.6),
            arrowprops=dict(arrowstyle="-|>", color=NAVY, lw=0.9, ls=":"))
ax.text(92, 32.4, "RH : une matinée par semaine,\nperçues comme proches de la direction", fontsize=5.8,
        color=NAVY, style="italic", ha="right", va="center")

# ---------- Filtres cognitifs ----------
ax.add_patch(FancyBboxPatch((68.5, 64.5), 23, 9.8, boxstyle="round,pad=0.3", fc=ACCENT_L, ec=ACCENT, lw=0.8))
ax.text(80, 72.3, "Filtres cognitifs des décideurs", ha="center", fontsize=6.5, weight="bold", color=ACCENT)
ax.text(80, 68.1, "(hypothèses) attention, disponibilité,\nconfirmation et escalade, attribution,\n"
        "pensée de groupe, découplage", ha="center", va="center", fontsize=5.8, color="#333333")

# ---------- Légende ----------
ly = 7
ax.plot([3, 9], [ly, ly], color=NAVY, lw=1.2)
ax.text(10, ly, "délégation / désignation", va="center", fontsize=6.2)
ax.plot([36, 42], [ly, ly], color=ACCENT, lw=1.2, ls=(0, (4, 3)))
ax.text(43, ly, "flux d'information ascendant", va="center", fontsize=6.2)
cross(72, ly, 1.1)
ax.text(74.5, ly, "rupture ou filtrage", va="center", fontsize=6.2)
ax.text(50, 2, "Sources : Unisanté (s.d.-a, s.d.-b, s.d.-c) ; État de Vaud (s.d.-b, s.d.-c) ; organigramme de "
        "l'auteur (annexe A) ; entretiens E1 et E2.", ha="center", fontsize=5.8, color=GREY_D, style="italic")

fig.savefig("schema_orchestration_CTMG.png", dpi=300, bbox_inches="tight", facecolor="white")
print("ok")
