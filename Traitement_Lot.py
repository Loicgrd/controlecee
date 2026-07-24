"""
Application Streamlit — Suivi des contrôles CEE (fichier "Synthèse")

Étapes :
1. Import du fichier "Synthèse - <LOT>.xlsx"
2. Tableau de copie rapide pour Odicée (REFERENCE interne / Conclusion de l'audit / Commentaires généraux)
   sur les lignes dont la conclusion d'audit est renseignée.
3. Calcul de la "Surface retenue dans la demande" pour les fiches BAR-EN-101/102/103/105, et
   correction proportionnelle des volumes (kWh cumac) quand la surface retenue diffère de la
   surface déclarée.
4. Si au moins une surface a diminué : import de la/les fiche(s) BAR (export "Import lots de
   travaux") pour y répercuter la surface retenue sur la ligne "ID lot de travaux" correspondante.
5. Téléchargement des fichiers corrigés.

Tableau Odicée : REFERENCE interne de l'opération + Conclusion de l'audit + Conclusion du
contrôle par contact.
"""

import io
import re

import openpyxl
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Suivi contrôles CEE — QV580M", layout="wide")

FICHES_CONCERNEES = ("BAR-EN-101", "BAR-EN-102", "BAR-EN-103", "BAR-EN-105")


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


def find_col(headers, prefix):
    """Retourne l'indice de la 1ère colonne dont l'en-tête commence par `prefix`."""
    prefix_n = normalize(prefix).lower()
    for c, h in headers.items():
        if h.lower().startswith(prefix_n):
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
col_fiche = find_col(headers, "REFERENCE DE LA FICHE")
col_decl = find_col(headers, "Surface déclarée dans l'AH/facture")
col_mesuree = find_col(headers, "Surface mesurée par le bureau de contrôle")
col_estimee = find_col(headers, "Surface estimée par le bureau de contrôle")
col_retenue = find_col(headers, "Surface retenue dans la demande")
col_ecart = find_col(headers, "Écart entre surface mesurée")
col_vol_hp = find_col(headers, "VOLUME HORS PRECARITE")
col_vol_prec = find_col(headers, "VOLUME PRECARITE")

required = {
    "REFERENCE interne de l'opération": col_ref,
    "Conclusion de l'audit": col_conclusion,
    "Conclusion du contrôle par contact": col_conclusion_contact,
    "REFERENCE DE LA FICHE": col_fiche,
    "Surface déclarée dans l'AH/facture": col_decl,
    "Surface mesurée par le bureau de contrôle": col_mesuree,
    "Surface estimée par le bureau de contrôle": col_estimee,
    "Surface retenue dans la demande": col_retenue,
    "Écart entre surface mesurée...": col_ecart,
    "VOLUME HORS PRECARITE": col_vol_hp,
    "VOLUME PRECARITE": col_vol_prec,
}
missing = [name for name, c in required.items() if c is None]
if missing:
    st.error("Colonnes introuvables dans le fichier : " + ", ".join(missing))
    st.stop()

# Lignes dont la conclusion d'audit est renseignée
rows = []
for r in range(header_row + 1, ws_read.max_row + 1):
    ref_val = ws_read.cell(row=r, column=col_ref).value
    concl_val = ws_read.cell(row=r, column=col_conclusion).value
    if ref_val is None or str(ref_val).strip() == "":
        continue
    if concl_val is None or str(concl_val).strip() == "":
        continue
    rows.append(r)

st.success(f"{len(rows)} ligne(s) avec conclusion d'audit renseignée.")

# --------------------------------------------------------------------------------------
# Étape 1 — Tableau à copier dans Odicée
# --------------------------------------------------------------------------------------

st.header(
    "1. Tableau à copier-coller dans Odicée",
    help=(
        "Ligne conservée si les deux conditions sont vraies :\n"
        "- « REFERENCE interne de l'opération » est renseignée\n"
        "- « Conclusion de l'audit » est renseignée (Satisfaisant / Non satisfaisant / ...)"
    ),
)

odicee_data = []
for r in rows:
    odicee_data.append(
        {
            "REFERENCE interne de l'opération": ws_read.cell(row=r, column=col_ref).value,
            "Conclusion de l'audit": ws_read.cell(row=r, column=col_conclusion).value,
            "Conclusion du contrôle par contact": ws_read.cell(
                row=r, column=col_conclusion_contact
            ).value,
        }
    )
df_odicee = pd.DataFrame(odicee_data)
st.dataframe(
    df_odicee,
    use_container_width=True,
    hide_index=True,
    column_config={
        "REFERENCE interne de l'opération": st.column_config.Column(
            help="Critère de tri : ligne affichée seulement si cette cellule ET « Conclusion de l'audit » sont non vides."
        ),
        "Conclusion de l'audit": st.column_config.Column(
            help="C'est la présence d'une valeur ici (Satisfaisant / Non satisfaisant / ...) qui décide si la ligne apparaît dans ce tableau."
        ),
        "Conclusion du contrôle par contact": st.column_config.Column(
            help="Affichée à titre indicatif pour cette même ligne ; ne participe pas au filtre."
        ),
    },
)

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
        "Si la surface retenue diffère de la surface déclarée, les volumes HORS PRECARITE et "
        "PRECARITE sont recalculés au même prorata."
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

    if declaree is None or mesuree is None:
        skipped.append({"REFERENCE interne": ref_val, "Fiche": fiche_val, "Motif": "surface déclarée/mesurée non exploitable"})
        continue

    if mesuree < declaree:
        retenue = mesuree
    elif ecart is not None and ecart > 10:
        if estimee is None:
            skipped.append({"REFERENCE interne": ref_val, "Fiche": fiche_val, "Motif": "écart > 10% mais surface estimée absente"})
            continue
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
        if old_hp is not None:
            cell_hp.value = round(old_hp * ratio)
        if old_prec is not None:
            cell_prec.value = round(old_prec * ratio)

    id_lot = str(ref_val).rsplit("-", 1)[-1].strip()
    changes.append(
        {
            "REFERENCE interne": ref_val,
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
    st.dataframe(
        pd.DataFrame(changes),
        use_container_width=True,
        hide_index=True,
        column_config={
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
st.download_button(
    "⬇️ Télécharger la Synthèse mise à jour",
    data=buf_synth.getvalue(),
    file_name=f"MAJ_{synth_file.name}",
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
