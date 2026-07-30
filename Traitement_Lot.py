"""
Application Streamlit — Suivi et traitement des lots de contrôle CEE (fichier "Synthèse")

Étapes (après import du fichier "Synthèse - <LOT>.xlsx") :
1. Tableau de copie rapide pour Odicée (blocs "Contrôle sur site" / "Contrôle par contact").
2. Calcul de la "Surface retenue dans la demande" pour les fiches BAR-EN-101/102/103/105, et
   correction proportionnelle des volumes (kWh cumac) quand la surface retenue diffère de la
   surface déclarée.
3. Mise à jour de la/les fiche(s) BAR (export "Import lots de travaux") si une surface a diminué.
4. Taux de contrôle du lot, comparaison aux seuils réglementaires (arrêté du 27/07/2026) et
   catégorisation du lot (Cas 1/2/3).
5. Export Excel par bailleur, avec commentaire généré selon le cas du lot.
6. Génération des mails clients et ouverture d'un brouillon Outlook par bailleur.

Tableau Odicée : REFERENCE interne de l'opération + Conclusion de l'audit + Conclusion du
contrôle par contact.
"""

import io
import os
import re
import urllib.parse
from decimal import ROUND_CEILING, ROUND_HALF_UP, Decimal
from pathlib import Path

import openpyxl
import pandas as pd
import streamlit as st
from openpyxl.styles import Alignment, Font, PatternFill

from cee_lots_data import get_seuils_fiche, parse_date_fr

st.set_page_config(page_title="Traitement des lots CEE", layout="wide")

FICHES_CONCERNEES = ("BAR-EN-101", "BAR-EN-102", "BAR-EN-103", "BAR-EN-105")


def sanitize_filename(name):
    return re.sub(r'[\\/*?:"<>|]', "", str(name)).strip()


# --------------------------------------------------------------------------------------
# Utilitaires génériques
# --------------------------------------------------------------------------------------

def normalize(s):
    """Nettoie un texte d'en-tête pour permettre des comparaisons robustes (retours à la
    ligne, astérisques de notes, espaces multiples...)."""
    if s is None:
        return ""
    s = str(s).replace("\n", " ")
    s = re.sub(r"\s+", " ", s).strip()
    return s.lstrip("*").strip()


def find_header_row(ws, marker, max_row=10):
    """Repère la ligne d'en-têtes en cherchant une cellule contenant `marker`."""
    marker_n = normalize(marker).lower()
    for r in range(1, max_row + 1):
        for c in range(1, ws.max_column + 1):
            v = ws.cell(row=r, column=c).value
            if v and marker_n in normalize(v).lower():
                return r
    return None


def build_header_map(ws, header_row):
    headers = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(row=header_row, column=c).value
        if v:
            headers[c] = normalize(v)
    return headers


def _loosen(s):
    """Variante encore plus tolérante utilisée uniquement pour le rapprochement des
    en-têtes de colonnes : les tirets et espaces sont traités de façon équivalente
    (ex : 'non-qualité' == 'non qualité')."""
    return re.sub(r"[\s\-]+", " ", s).strip().lower()


def find_col(headers, prefix):
    """Retourne l'indice de la 1ère colonne dont l'en-tête commence par `prefix`."""
    prefix_n = _loosen(normalize(prefix))
    for c, h in headers.items():
        if _loosen(h).startswith(prefix_n):
            return c
    return None


def to_number(v):
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(",", ".")
    if s == "" or s.lower() in ("non vérifiable", "nv", "sans objet"):
        return None
    try:
        return float(s)
    except ValueError:
        return None


def fmt_like(original_value, new_value):
    """Formate `new_value` en respectant le type de la cellule d'origine (texte vs nombre),
    sans conserver de zéros inutiles (ex : 104.3 et non 104.30)."""
    new_value = round(new_value, 2)
    if isinstance(original_value, str):
        return f"{new_value:g}"
    return new_value


def round_half_up(x):
    """Arrondi à l'entier le plus proche, avec .5 arrondi vers le haut (contrairement à
    round() de Python qui arrondit .5 vers l'entier pair le plus proche)."""
    return int(Decimal(str(x)).quantize(Decimal("1"), rounding=ROUND_HALF_UP))


def round_ceiling(x):
    """Arrondi systématiquement à l'entier supérieur."""
    return int(Decimal(str(x)).quantize(Decimal("1"), rounding=ROUND_CEILING))


# --------------------------------------------------------------------------------------
# UI
# --------------------------------------------------------------------------------------

st.title("Suivi des contrôles CEE — Synthèse QV580M")

synth_file = st.file_uploader(
    "1. Charger le fichier Synthèse (ex : Synthèse - QV580M.xlsx)", type=["xlsx"]
)

if not synth_file:
    st.info("Charge le fichier Synthèse pour commencer.")
    st.stop()

synth_bytes = synth_file.getvalue()

# Deux lectures : une pour les valeurs figées (calculs), une pour l'écriture (préserve les
# formules/format s'il y en a).
wb_write = openpyxl.load_workbook(io.BytesIO(synth_bytes), data_only=False)
wb_read = openpyxl.load_workbook(io.BytesIO(synth_bytes), data_only=True)
ws_write = wb_write.active
ws_read = wb_read.active

header_row = find_header_row(ws_read, "REFERENCE interne de l'opération")
if not header_row:
    st.error("Colonne \"REFERENCE interne de l'opération\" introuvable dans ce fichier.")
    st.stop()

headers = build_header_map(ws_read, header_row)

col_ref = find_col(headers, "REFERENCE interne de l'opération")
col_conclusion = find_col(headers, "Conclusion de l'audit")
col_conclusion_contact = find_col(headers, "Conclusion du contrôle par contact")
col_commentaires_generaux = find_col(headers, "Commentaires généraux")
col_non_qualite = find_col(headers, "Commentaire sur le type de non qualité relevée")
col_fiche = find_col(headers, "REFERENCE DE LA FICHE")
col_decl = find_col(headers, "Surface déclarée dans l'AH/facture")
col_mesuree = find_col(headers, "Surface mesurée par le bureau de contrôle")
col_estimee = find_col(headers, "Surface estimée par le bureau de contrôle")
col_retenue = find_col(headers, "Surface retenue dans la demande")
col_ecart = find_col(headers, "Écart entre surface mesurée")
col_vol_hp = find_col(headers, "VOLUME HORS PRECARITE")
col_vol_prec = find_col(headers, "VOLUME PRECARITE")


def cell_or_none(row, col):
    """Lit une cellule seulement si la colonne existe dans ce fichier ; None sinon.
    Utile pour les colonnes optionnelles (toutes les fiches n'ont pas de contrôle par
    contact, ni de colonnes de surface — certaines mesurent une longueur à la place)."""
    return ws_read.cell(row=row, column=col).value if col else None


# Certaines fiches n'ont pas de contrôle par contact (100% sur site), et certaines fiches
# (réseaux isolés, mesurées en longueur) n'ont pas du tout de colonnes de surface : ces
# groupes de colonnes sont donc optionnels, pas bloquants pour le reste de l'application.
has_contact = col_conclusion_contact is not None
has_surface_cols = all([col_decl, col_mesuree, col_estimee, col_retenue, col_ecart, col_vol_hp, col_vol_prec])

required = {
    "REFERENCE interne de l'opération": col_ref,
    "Conclusion de l'audit": col_conclusion,
    "REFERENCE DE LA FICHE": col_fiche,
}
missing = [name for name, c in required.items() if c is None]
if missing:
    st.error("Colonnes introuvables dans le fichier : " + ", ".join(missing))
    st.stop()

if not has_contact:
    st.info("Ce fichier n'a pas de colonne « Conclusion du contrôle par contact » — fiche à contrôle uniquement sur site.")
if not has_surface_cols:
    st.info(
        "Ce fichier n'a pas les colonnes de surface (déclarée/mesurée/estimée/retenue) — "
        "les étapes 2 et 3 (surface retenue, volumes, mise à jour fiche BAR) sont masquées "
        "(fiche non basée sur une surface, par ex. réseau isolé mesuré en longueur)."
    )

# Lignes dont la conclusion du contrôle sur site est renseignée (utilisées aussi pour le
# calcul de surface à l'étape 2, où seul le contrôle sur site mesure une surface).
rows = []
for r in range(header_row + 1, ws_read.max_row + 1):
    ref_val = ws_read.cell(row=r, column=col_ref).value
    concl_val = ws_read.cell(row=r, column=col_conclusion).value
    if ref_val is None or str(ref_val).strip() == "":
        continue
    if concl_val is None or str(concl_val).strip() == "":
        continue
    rows.append(r)

# Lignes dont la conclusion du contrôle par contact est renseignée (affichage uniquement).
rows_contact = []
if has_contact:
    for r in range(header_row + 1, ws_read.max_row + 1):
        ref_val = ws_read.cell(row=r, column=col_ref).value
        concl_val = ws_read.cell(row=r, column=col_conclusion_contact).value
        if ref_val is None or str(ref_val).strip() == "":
            continue
        if concl_val is None or str(concl_val).strip() == "":
            continue
        rows_contact.append(r)

st.success(
    f"{len(rows)} ligne(s) avec conclusion du contrôle sur site renseignée"
    + (f", {len(rows_contact)} ligne(s) avec conclusion du contrôle par contact renseignée." if has_contact else " (pas de contrôle par contact dans ce fichier).")
)

# --------------------------------------------------------------------------------------
# Étape 1 — Tableau à copier dans Odicée
# --------------------------------------------------------------------------------------

st.header(
    "1. Tableau à copier-coller dans Odicée",
    help=(
        "Deux blocs, un par type de contrôle, pour un copier-coller séparé :\n"
        "- Contrôle sur site : « REFERENCE interne » et « Conclusion du contrôle sur site » renseignées\n"
        "- Contrôle par contact : « REFERENCE interne » et « Conclusion du contrôle par contact » renseignées"
    ),
)

st.subheader("Bloc — Contrôle sur site")
odicee_site_data = [
    {
        "REFERENCE interne de l'opération": ws_read.cell(row=r, column=col_ref).value,
        "Conclusion du contrôle sur site": ws_read.cell(row=r, column=col_conclusion).value,
        "Commentaires généraux": cell_or_none(r, col_commentaires_generaux),
    }
    for r in rows
]
st.dataframe(
    pd.DataFrame(odicee_site_data),
    use_container_width=True,
    hide_index=True,
    column_config={
        "REFERENCE interne de l'opération": st.column_config.Column(
            help="Critère de tri : ligne affichée seulement si cette cellule ET « Conclusion du contrôle sur site » sont non vides."
        ),
        "Conclusion du contrôle sur site": st.column_config.Column(
            help="Colonne source dans le fichier : « Conclusion de l'audit ». C'est sa présence qui décide si la ligne apparaît dans ce bloc."
        ),
        "Commentaires généraux": st.column_config.Column(
            help="Affichée telle quelle (vide si non renseignée)."
        ),
    },
)

if has_contact:
    st.subheader("Bloc — Contrôle par contact")
    odicee_contact_data = [
        {
            "REFERENCE interne de l'opération": ws_read.cell(row=r, column=col_ref).value,
            "Conclusion du contrôle par contact": ws_read.cell(row=r, column=col_conclusion_contact).value,
            "Commentaire sur le type de non qualité relevée": cell_or_none(r, col_non_qualite),
        }
        for r in rows_contact
    ]
    st.dataframe(
        pd.DataFrame(odicee_contact_data),
        use_container_width=True,
        hide_index=True,
        column_config={
            "REFERENCE interne de l'opération": st.column_config.Column(
                help="Critère de tri : ligne affichée seulement si cette cellule ET « Conclusion du contrôle par contact » sont non vides."
            ),
            "Conclusion du contrôle par contact": st.column_config.Column(
                help="C'est la présence d'une valeur ici qui décide si la ligne apparaît dans ce bloc."
            ),
            "Commentaire sur le type de non qualité relevée": st.column_config.Column(
                help="Renseigné en général quand la conclusion est « Non satisfaisant » ; affiché tel quel (vide si non renseigné)."
            ),
        },
    )

if has_surface_cols:
    # --------------------------------------------------------------------------------------
    # Étape 2 — Surface retenue + volumes
    # --------------------------------------------------------------------------------------

    st.header(
        "2. Surface retenue dans la demande & correction des volumes",
        help=(
            "Parmi les lignes ci-dessus, seules celles dont « REFERENCE DE LA FICHE » commence par "
            "BAR-EN-101, BAR-EN-102, BAR-EN-103 ou BAR-EN-105 sont traitées.\n\n"
            "Règle de calcul de la surface retenue :\n"
            "1. Si Surface mesurée < Surface déclarée → retenue = Surface mesurée\n"
            "2. Sinon, si Écart > 10 % → retenue = Surface estimée\n"
            "3. Sinon → retenue = Surface déclarée\n\n"
            "Si la surface retenue diffère de la surface déclarée, les volumes sont recalculés "
            "au même prorata : VOLUME HORS PRECARITE est arrondi à l'entier supérieur, puis "
            "VOLUME PRECARITE = total recalculé (arrondi normalement) − VOLUME HORS PRECARITE."
        ),
    )
    st.caption(
        "Appliqué uniquement aux lignes BAR-EN-101 / BAR-EN-102 / BAR-EN-103 / BAR-EN-105 "
        "parmi les lignes ci-dessus."
    )

    changes = []
    skipped = []
    for r in rows:
        fiche_val = normalize(ws_read.cell(row=r, column=col_fiche).value or "")
        if not any(fiche_val.startswith(f) for f in FICHES_CONCERNEES):
            continue

        declaree = to_number(ws_read.cell(row=r, column=col_decl).value)
        mesuree = to_number(ws_read.cell(row=r, column=col_mesuree).value)
        estimee = to_number(ws_read.cell(row=r, column=col_estimee).value)
        ecart = to_number(ws_read.cell(row=r, column=col_ecart).value)
        ref_val = ws_read.cell(row=r, column=col_ref).value

        if declaree is None:
            skipped.append({"REFERENCE interne": ref_val, "Fiche": fiche_val, "Motif": "surface déclarée absente/non exploitable"})
            continue

        # Par défaut : surface déclarée. Uniquement remplacée si mesurée < déclarée, ou si
        # l'écart est strictement supérieur à 10 % (auquel cas on prend l'estimée).
        if mesuree is not None and mesuree < declaree:
            retenue = mesuree
        elif ecart is not None and ecart > 10 and estimee is not None:
            retenue = estimee
        else:
            retenue = declaree

        # Écriture de la surface retenue
        ws_write.cell(row=r, column=col_retenue).value = retenue

        decrease = retenue < declaree
        if retenue != declaree:
            ratio = retenue / declaree
            cell_hp = ws_write.cell(row=r, column=col_vol_hp)
            cell_prec = ws_write.cell(row=r, column=col_vol_prec)
            old_hp = to_number(ws_read.cell(row=r, column=col_vol_hp).value)
            old_prec = to_number(ws_read.cell(row=r, column=col_vol_prec).value)
            if old_hp is not None and old_prec is not None:
                # VOLUME HORS PRECARITE arrondi à l'entier supérieur, VOLUME PRECARITE déduit
                # du total (arrondi normalement) moins la valeur HORS PRECARITE déjà arrondie,
                # pour que la somme des deux reste cohérente avec le total recalculé.
                new_hp = round_ceiling(old_hp * ratio)
                new_total = round_half_up((old_hp + old_prec) * ratio)
                new_prec = new_total - new_hp
                cell_hp.value = new_hp
                cell_prec.value = new_prec
            elif old_hp is not None:
                cell_hp.value = round_half_up(old_hp * ratio)
            elif old_prec is not None:
                cell_prec.value = round_half_up(old_prec * ratio)

        id_lot = str(ref_val).rsplit("-", 1)[-1].strip()
        dossier = str(ref_val).split("-", 1)[0].strip()
        changes.append(
            {
                "REFERENCE interne": ref_val,
                "Dossier": dossier,
                "ID lot de travaux": id_lot,
                "Fiche": ws_read.cell(row=r, column=col_fiche).value,
                "Surface déclarée": declaree,
                "Surface mesurée": mesuree,
                "Surface estimée": estimee,
                "Écart (%)": ecart,
                "Surface retenue": retenue,
                "Diminution": "Oui" if decrease else "Non",
            }
        )

    any_decrease = any(c["Diminution"] == "Oui" for c in changes)

    if changes:
        with st.expander(f"Voir le détail des {len(changes)} ligne(s) traitées", expanded=False):
            st.dataframe(
                pd.DataFrame(changes),
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Dossier": st.column_config.Column(
                        help="Partie avant le « - » de REFERENCE interne de l'opération : sert de NUMERODOSSIER pour le lien Odicée à l'étape 3."
                    ),
                    "Fiche": st.column_config.Column(
                        help="Seules les fiches BAR-EN-101, BAR-EN-102, BAR-EN-103 et BAR-EN-105 sont retenues pour ce calcul."
                    ),
                    "Écart (%)": st.column_config.Column(
                        help="= (Surface déclarée − Surface mesurée) / Surface mesurée × 100. Détermine si on utilise la surface estimée quand mesurée ≥ déclarée (écart > 10 %)."
                    ),
                    "Surface retenue": st.column_config.Column(
                        help="Mesurée si mesurée < déclarée ; sinon estimée si écart > 10 % ; sinon déclarée."
                    ),
                    "Diminution": st.column_config.Column(
                        help="« Oui » si Surface retenue < Surface déclarée. Dès qu'au moins une ligne est à « Oui », l'étape 3 (import de fiche BAR) apparaît."
                    ),
                },
            )
    else:
        st.info("Aucune ligne BAR-EN-101/102/103/105 exploitable parmi les lignes filtrées.")

    if skipped:
        with st.expander(f"{len(skipped)} ligne(s) ignorée(s) (surface non exploitable)"):
            st.dataframe(pd.DataFrame(skipped), use_container_width=True, hide_index=True)

    # Fichier Synthèse mis à jour, prêt à télécharger
    buf_synth = io.BytesIO()
    wb_write.save(buf_synth)
    synth_path = Path(synth_file.name)
    synth_v2_name = f"{synth_path.stem} V2{synth_path.suffix}"
    st.download_button(
        "⬇️ Télécharger la Synthèse mise à jour",
        data=buf_synth.getvalue(),
        file_name=synth_v2_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    # --------------------------------------------------------------------------------------
    # Étape 3 — Mise à jour de la/les fiche(s) BAR
    # --------------------------------------------------------------------------------------

    if any_decrease:
        st.header(
            "3. Mise à jour de la/les fiche(s) BAR",
            help=(
                "N'apparaît que si au moins une ligne a « Diminution » = Oui à l'étape 2.\n\n"
                "La correspondance entre la Synthèse et la fiche BAR se fait sur :\n"
                "REFERENCE interne de l'opération (partie après le dernier « - ») "
                "= ID lot de travaux de la fiche BAR."
            ),
        )
        st.caption(
            "Au moins une surface retenue est inférieure à la surface déclarée : "
            "importe la ou les fiches BAR (export \"Import lots de travaux\") pour y reporter "
            "la surface retenue, sur la ligne dont l'\"ID lot de travaux\" correspond."
        )

        rows_to_apply = [c for c in changes if c["Surface retenue"] != c["Surface déclarée"]]

        dossier_fiches = {}
        for c in rows_to_apply:
            if c["Dossier"]:
                dossier_fiches.setdefault(c["Dossier"], set()).add(str(c["Fiche"]).strip())

        if dossier_fiches:
            st.markdown(
                "**Accès direct au(x) dossier(s) concerné(s) par une surface modifiée :**  \n"
                + "  \n".join(
                    f"- [{numero} ({', '.join(sorted(fiches))})](https://odicee.edf.fr/dossiers/{numero})"
                    for numero, fiches in sorted(dossier_fiches.items())
                )
            )

        bar_files = st.file_uploader(
            "Importer la ou les fiches BAR (ex : T155142_BAR-EN-102_A14_1.xlsx)",
            type=["xlsx"],
            accept_multiple_files=True,
            key="bar_upload",
        )

        if bar_files:
            for bf in bar_files:
                bar_bytes = bf.getvalue()
                bwb = openpyxl.load_workbook(io.BytesIO(bar_bytes), data_only=False)
                bws = bwb.active

                bar_header_row = find_header_row(bws, "ID lot de travaux")
                if not bar_header_row:
                    st.error(f"{bf.name} : colonne \"ID lot de travaux\" introuvable.")
                    continue

                bar_headers = build_header_map(bws, bar_header_row)
                col_id_lot = find_col(bar_headers, "ID lot de travaux")
                col_surface = find_col(bar_headers, "surface")

                if not col_id_lot or not col_surface:
                    st.error(f"{bf.name} : colonnes \"ID lot de travaux\" / \"surface\" introuvables.")
                    continue

                applied = []
                for r in range(bar_header_row + 1, bws.max_row + 1):
                    id_val = bws.cell(row=r, column=col_id_lot).value
                    id_num = to_number(id_val)
                    if id_num is None:
                        continue
                    id_lot_bar = str(int(id_num))

                    for chg in rows_to_apply:
                        if chg["ID lot de travaux"] == id_lot_bar:
                            cell = bws.cell(row=r, column=col_surface)
                            old_val = cell.value
                            new_val = fmt_like(old_val, chg["Surface retenue"])
                            cell.value = new_val
                            applied.append(
                                {
                                    "ID lot de travaux": id_lot_bar,
                                    "Ancienne surface": old_val,
                                    "Nouvelle surface": new_val,
                                }
                            )

                if applied:
                    st.write(f"**{bf.name}** — {len(applied)} ligne(s) mise(s) à jour :")
                    st.dataframe(pd.DataFrame(applied), use_container_width=True, hide_index=True)
                    out = io.BytesIO()
                    bwb.save(out)
                    st.download_button(
                        f"⬇️ Télécharger {bf.name} corrigé",
                        data=out.getvalue(),
                        file_name=f"MAJ_{bf.name}",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"dl_{bf.name}",
                    )
                else:
                    st.info(f"{bf.name} : aucune ligne correspondant à un \"ID lot de travaux\" à corriger n'a été trouvée.")
else:
    changes = []
    any_decrease = False

# ========================================================================================
# Étape 4 — Taux de contrôle, conformité réglementaire et catégorisation du lot
# ========================================================================================

st.header(
    "4. Taux de contrôle et conformité du lot",
    help=(
        "Taux de satisfaisant (site ou contact) = nombre de satisfaisant / nombre total "
        "d'opérations du lot (toutes les lignes avec une REFERENCE interne).\n\n"
        "Taux de non satisfaisant sur site = nombre de non satisfaisant sur site / nombre "
        "de lignes avec une conclusion de contrôle sur site renseignée.\n\n"
        "Les seuils réglementaires sont recherchés dans la table de l'arrêté du 27/07/2026 "
        "pour la fiche BAR du lot, à la date d'engagement la plus récente."
    ),
)

# Toutes les opérations du lot (une ligne = une opération dès que REFERENCE interne est renseignée)
all_op_rows = []
for r in range(header_row + 1, ws_read.max_row + 1):
    ref_val = ws_read.cell(row=r, column=col_ref).value
    if ref_val is not None and str(ref_val).strip() != "":
        all_op_rows.append(r)

total_ops = len(all_op_rows)


def classify_conclusion(value):
    v = normalize(value).lower() if value else ""
    if not v:
        return "non_visite"
    if v == "satisfaisant":
        return "satisfaisant"
    if v == "non satisfaisant":
        return "non_satisfaisant"
    if "vérifiable" in v or "inaccessible" in v:
        return "inaccessible"
    return "autre"


nb_satisfaisant_site = 0
nb_non_satisfaisant_site = 0
nb_controles_site = 0
nb_satisfaisant_contact = 0
fiche_lot = None
dates_engagement = []

col_date_engagement = find_col(headers, "DATE D'ENGAGEMENT")

for r in all_op_rows:
    cls_site = classify_conclusion(ws_read.cell(row=r, column=col_conclusion).value)
    if cls_site != "non_visite":
        nb_controles_site += 1
    if cls_site == "satisfaisant":
        nb_satisfaisant_site += 1
    elif cls_site == "non_satisfaisant":
        nb_non_satisfaisant_site += 1

    if classify_conclusion(cell_or_none(r, col_conclusion_contact)) == "satisfaisant":
        nb_satisfaisant_contact += 1

    if fiche_lot is None:
        f_val = ws_read.cell(row=r, column=col_fiche).value
        if f_val and str(f_val).strip():
            fiche_lot = str(f_val).strip()

    if col_date_engagement:
        d = parse_date_fr(ws_read.cell(row=r, column=col_date_engagement).value)
        if d:
            dates_engagement.append(d)

taux_s_site = (nb_satisfaisant_site / total_ops * 100) if total_ops else 0.0
taux_s_contact = (nb_satisfaisant_contact / total_ops * 100) if total_ops else 0.0
taux_ns_site = (nb_non_satisfaisant_site / nb_controles_site * 100) if nb_controles_site else 0.0
date_engagement_max = max(dates_engagement) if dates_engagement else None

seuil_site, seuil_contact = get_seuils_fiche(fiche_lot, date_engagement_max) if fiche_lot and date_engagement_max else (None, None)

col_info1, col_info2, col_info3 = st.columns(3)
with col_info1:
    st.metric("Fiche BAR du lot", fiche_lot or "—")
with col_info2:
    st.metric("Date d'engagement la plus récente", date_engagement_max.strftime("%d/%m/%Y") if date_engagement_max else "—")
with col_info3:
    seuil_txt = ""
    if seuil_site is not None:
        seuil_txt += f"Site ≥ {seuil_site:g}%"
    if seuil_contact is not None:
        seuil_txt += (" · " if seuil_txt else "") + f"Contact ≥ {seuil_contact:g}%"
    st.metric("Seuils réglementaires trouvés", seuil_txt or "non trouvés")

if fiche_lot and date_engagement_max and seuil_site is None and seuil_contact is None:
    st.warning(
        f"Aucun seuil trouvé dans la table de l'arrêté pour la fiche « {fiche_lot} » à la date "
        f"{date_engagement_max.strftime('%d/%m/%Y')}. Vérifie le code de fiche ou complète la "
        "table dans cee_lots_data.py."
    )

seuil_ns_max = st.number_input(
    "Seuil maximal de non satisfaisant sur site (%)",
    min_value=0.0, max_value=100.0, value=14.0, step=0.5,
    help="Modifiable ponctuellement. Le taux de non satisfaisant du lot ne doit pas dépasser ce seuil.",
)

# --- Conformité du taux de satisfaisant ---
site_ok = True if seuil_site is None else (taux_s_site >= seuil_site)
if seuil_contact is None:
    contact_ok = True
else:
    contact_ok_direct = taux_s_contact >= seuil_contact
    if seuil_site is not None:
        somme_ok = (taux_s_site + taux_s_contact) >= (seuil_site + seuil_contact)
        contact_ok = contact_ok_direct or somme_ok
    else:
        contact_ok = contact_ok_direct
taux_satisfaisant_conforme = site_ok and contact_ok

# --- Conformité du taux de non satisfaisant ---
ns_conforme = taux_ns_site <= seuil_ns_max

# --- Catégorisation du lot ---
if taux_satisfaisant_conforme and ns_conforme:
    cas_lot = 1
    conclusion_cas = "La totalité des opérations dans le lot est déposable."
elif taux_satisfaisant_conforme and not ns_conforme:
    cas_lot = 2
    conclusion_cas = "On peut déposer toutes les opérations visitées seulement."
else:
    cas_lot = 3
    conclusion_cas = "On ne peut déposer que les opérations contrôlées satisfaisantes et non satisfaisantes."

def carte_taux(titre, ok, valeur_txt, sous_texte):
    color = "#C6EFCE" if ok else "#FFCCCC"
    icone = "✅" if ok else "❌"
    return (
        f"<div style='background:{color};padding:12px;border-radius:8px;text-align:center'>"
        f"<b>{titre}</b> {icone}<br>"
        f"<span style='font-size:22px;font-weight:bold'>{valeur_txt}</span><br>"
        f"<small>{sous_texte}</small>"
        f"</div>"
    )


taux_cols = st.columns(3)
with taux_cols[0]:
    sous = (f"{nb_satisfaisant_site}/{total_ops} — seuil ≥ {seuil_site:g}%" if seuil_site is not None
            else f"{nb_satisfaisant_site}/{total_ops} — pas de seuil pour cette fiche")
    st.markdown(
        carte_taux("Taux satisfaisant sur site", site_ok, f"{taux_s_site:.1f} %", sous),
        unsafe_allow_html=True,
    )
with taux_cols[1]:
    sous = (f"{nb_satisfaisant_contact}/{total_ops} — seuil ≥ {seuil_contact:g}%" if seuil_contact is not None
            else f"{nb_satisfaisant_contact}/{total_ops} — pas de seuil pour cette fiche")
    st.markdown(
        carte_taux("Taux satisfaisant par contact", contact_ok, f"{taux_s_contact:.1f} %", sous),
        unsafe_allow_html=True,
    )
with taux_cols[2]:
    sous = f"{nb_non_satisfaisant_site}/{nb_controles_site} — seuil ≤ {seuil_ns_max:g}%"
    st.markdown(
        carte_taux("Taux non satisfaisant sur site", ns_conforme, f"{taux_ns_site:.1f} %", sous),
        unsafe_allow_html=True,
    )

cas_couleur = {1: "#e8f5e9", 2: "#fff3e0", 3: "#ffebee"}[cas_lot]
st.markdown(
    f"<div style='background-color:{cas_couleur}; padding:14px; border-radius:8px;'>"
    f"<b>Cas {cas_lot}</b> — {conclusion_cas}</div>",

    unsafe_allow_html=True,
)

# ========================================================================================
# Étape 5 — Export Excel et mail par bailleur
# ========================================================================================

st.header(
    "5. Export Excel et mail par bailleur",
    help=(
        "Un fichier par « RAISON SOCIALE du bénéficiaire de l'opération », avec les colonnes "
        "REFERENCE interne / Nom du site / Adresse / Code postal / Ville / Conclusion sur site / "
        "Conclusion par contact / Commentaire (généré selon le cas du lot).\n\n"
        "Le bouton mail ouvre le client mail par défaut (objet + corps pré-remplis) ; la pièce "
        "jointe n'est pas ajoutée automatiquement (limite du lien mailto) — télécharge-la et "
        "joins-la manuellement."
    ),
)

col_i = find_col(headers, "RAISON SOCIALE du bénéficiaire")
col_e = find_col(headers, "NOM DU SITE")
col_f = find_col(headers, "ADRESSE de l'opération")
col_g = find_col(headers, "CODE POSTAL")
col_h = find_col(headers, "VILLE")

if not all([col_i, col_e, col_f, col_g, col_h]):
    st.error("Colonnes bailleur/adresse introuvables (RAISON SOCIALE / NOM DU SITE / ADRESSE / CODE POSTAL / VILLE).")
else:
    def build_commentaire(cls_site, commentaires_generaux, cas):
        if cls_site == "satisfaisant":
            return "L'opération sera valorisée dans ce lot"
        if cls_site == "non_satisfaisant":
            txt = str(commentaires_generaux).strip() if commentaires_generaux else ""
            return txt or "Non satisfaisant sur site — voir Commentaires généraux"
        if cls_site == "inaccessible":
            if cas in (1, 2):
                return "L'opération sera valorisée dans ce lot"
            return "L'opération sera transférée dans un nouveau lot de contrôle si la date de fin de validité du dossier nous le permet"
        # non_visite
        if cas == 1:
            return "L'opération sera valorisée dans ce lot"
        return "L'opération sera transférée dans un nouveau lot de contrôle si l'opération nous le permet"

    bailleurs = {}
    for r in all_op_rows:
        bailleur = ws_read.cell(row=r, column=col_i).value
        bailleur = str(bailleur).strip() if bailleur else "(bailleur non renseigné)"

        concl_site_val = ws_read.cell(row=r, column=col_conclusion).value
        concl_contact_val = cell_or_none(r, col_conclusion_contact)
        cls_site = classify_conclusion(concl_site_val)
        commentaires_generaux_val = cell_or_none(r, col_commentaires_generaux)

        bailleurs.setdefault(bailleur, []).append(
            {
                "REFERENCE interne de l'opération": ws_read.cell(row=r, column=col_ref).value,
                "Nom du site": ws_read.cell(row=r, column=col_e).value,
                "Adresse": ws_read.cell(row=r, column=col_f).value,
                "Code postal": ws_read.cell(row=r, column=col_g).value,
                "Ville": ws_read.cell(row=r, column=col_h).value,
                "Conclusion du contrôle sur site": str(concl_site_val).strip() if concl_site_val and str(concl_site_val).strip() else "Non visité",
                "Conclusion du contrôle par contact": str(concl_contact_val).strip() if concl_contact_val and str(concl_contact_val).strip() else "Non visité",
                "Commentaire": build_commentaire(cls_site, commentaires_generaux_val, cas_lot),
                "Dossier": str(ws_read.cell(row=r, column=col_ref).value or "").split("-", 1)[0].strip(),
            }
        )

    COULEUR_COMMENTAIRE = {
        "satisfaisant": "C6EFCE",
        "non_satisfaisant": "FFC7CE",
        "inaccessible": "FFE0B2",
        "non_visite": "E0E0E0",
    }

    def build_excel_bailleur(lignes):
        wb_b = openpyxl.Workbook()
        ws_b = wb_b.active
        ws_b.title = "Résultats contrôle"
        headers_b = list(lignes[0].keys())
        headers_b = [h for h in headers_b if h != "Dossier"]  # colonne technique, pas affichée
        header_font = Font(bold=True, color="FFFFFF")
        header_fill = PatternFill("solid", fgColor="2F5496")
        for c, h in enumerate(headers_b, 1):
            cell = ws_b.cell(row=1, column=c, value=h)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        for r, ligne in enumerate(lignes, 2):
            for c, h in enumerate(headers_b, 1):
                cell = ws_b.cell(row=r, column=c, value=ligne[h])
                cell.alignment = Alignment(vertical="center", wrap_text=True)
            cls = classify_conclusion(ligne["Conclusion du contrôle sur site"] if ligne["Conclusion du contrôle sur site"] != "Non visité" else "")
            color = COULEUR_COMMENTAIRE.get(cls)
            if color:
                for c in range(1, len(headers_b) + 1):
                    ws_b.cell(row=r, column=c).fill = PatternFill("solid", fgColor=color)
        widths = [22, 25, 30, 12, 18, 22, 22, 45]
        for i, w in enumerate(widths[: len(headers_b)], 1):
            ws_b.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
        buf_b = io.BytesIO()
        wb_b.save(buf_b)
        return buf_b.getvalue()

    st.session_state["bailleurs_data"] = bailleurs
    st.session_state["cas_lot"] = cas_lot
    st.session_state["conclusion_cas"] = conclusion_cas

    guess_lot = re.sub(r"^Synth[eè]se\s*[-_]\s*", "", Path(synth_file.name).stem, flags=re.IGNORECASE).strip()
    num_lot = st.text_input("Numéro de lot (pour l'objet du mail)", value=guess_lot)

    for bailleur, lignes in sorted(bailleurs.items()):
        dossiers = sorted({l["Dossier"] for l in lignes if l["Dossier"]})
        liste_dossiers = "\n".join(f"- {d}" for d in dossiers)
        corps = (
            "Bonjour,\n\n"
            f"Nous avons reçu les résultats du lot {num_lot} ({fiche_lot}).\n\n"
            "La conclusion du lot est la suivante :\n\n"
            f"{conclusion_cas}\n\n"
            "Voici la liste des dossiers concernés par l'opération :\n\n"
            f"{liste_dossiers}\n\n"
            "Vous trouverez ci-joint les résultats des contrôles pour vos opérations.\n\n"
            "Votre interlocuteur EDF et nous-même restons disponibles.\n\n"
            "Bien à vous,\n\n"
            "L'équipe contrôle CEE,\n\n"
            "PROMOTELEC-SERVICES"
        )
        subject = f"Retour de contrôle {num_lot}"
        attach_name = f"{sanitize_filename(bailleur)}.xlsx"
        xlsx_bytes_b = build_excel_bailleur(lignes)
        mailto_url = f"mailto:?subject={urllib.parse.quote(subject)}&body={urllib.parse.quote(corps)}"

        with st.expander(f"🏢 {bailleur} — {len(lignes)} opération(s)"):
            bcol1, bcol2 = st.columns(2)
            with bcol1:
                st.link_button(f"📧 Envoyer le mail — {bailleur}", mailto_url)
            with bcol2:
                st.download_button(
                    f"⬇️ Télécharger l'Excel — {bailleur}",
                    data=xlsx_bytes_b,
                    file_name=attach_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"dl_bailleur_{bailleur}",
                )
