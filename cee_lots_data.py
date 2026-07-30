"""
cee_lots_data.py — Données réglementaires pour le traitement des lots de contrôle CEE.

Contient la table des taux minimaux de contrôles satisfaisants par fiche BAR/BAT/IND/TRA/...
telle que définie par l'annexe de l'arrêté du 27 juillet 2026 (modifiant l'arrêté du
28 septembre 2021 relatif aux contrôles dans le cadre du dispositif des CEE).

Chaque ligne de la table couvre une période et donne, quand ils existent, un seuil "sur site"
et un seuil "par contact" (en %). Certaines fiches n'ont qu'un des deux types de contrôle.
"""

from datetime import date
import re

# --------------------------------------------------------------------------------------
# Table de l'annexe (arrêté du 27 juillet 2026)
# --------------------------------------------------------------------------------------
# Chaque entrée : (liste de fiches concernées, [(date_debut, date_fin_ou_None, seuil_site_ou_None, seuil_contact_ou_None), ...])
# date_fin = None signifie "à compter de date_debut, sans fin connue".

_RAW_TABLE = [
    (["AGRI-TH-104"], [
        (date(2022, 7, 1), date(2022, 12, 31), 7.5, 15.0),
        (date(2023, 1, 1), date(2023, 12, 31), 10.0, 20.0),
        (date(2024, 1, 1), date(2024, 12, 31), 12.5, 25.0),
        (date(2025, 1, 1), date(2026, 12, 31), 15.0, 30.0),
        (date(2027, 1, 1), None, 20.0, 30.0),
    ]),
    (["BAR-EN-101", "BAR-EN-102", "BAR-EN-103", "BAR-EN-106", "BAR-EN-107", "BAR-TH-145", "BAR-TH-164"], [
        (date(2022, 1, 1), date(2022, 12, 31), 7.5, 15.0),
        (date(2023, 1, 1), date(2023, 12, 31), 10.0, 20.0),
        (date(2024, 1, 1), date(2024, 12, 31), 12.5, 25.0),
        (date(2025, 1, 1), date(2026, 12, 31), 15.0, 30.0),
        (date(2027, 1, 1), None, 20.0, 30.0),
    ]),
    (["BAR-TH-174", "BAR-TH-175"], [
        (date(2024, 1, 1), None, 100.0, None),
    ]),
    (["BAR-TH-113", "BAR-TH-159"], [
        (date(2022, 4, 1), date(2022, 12, 31), 7.5, 15.0),
        (date(2023, 1, 1), date(2023, 12, 31), 10.0, 20.0),
        (date(2024, 1, 1), date(2024, 12, 31), 12.5, 25.0),
        (date(2025, 1, 1), date(2026, 12, 31), 15.0, 30.0),
        (date(2027, 1, 1), None, 20.0, 30.0),
    ]),
    (["BAR-TH-171", "BAR-TH-172"], [
        (date(2022, 4, 1), date(2022, 12, 31), 7.5, 15.0),
        (date(2023, 1, 1), date(2023, 12, 31), 10.0, 20.0),
        (date(2024, 1, 1), date(2024, 12, 31), 12.5, 25.0),
        (date(2025, 1, 1), date(2025, 12, 31), 15.0, 30.0),
        (date(2026, 1, 1), date(2026, 12, 31), 25.0, 30.0),
        (date(2027, 1, 1), date(2027, 12, 31), 50.0, None),
        (date(2028, 1, 1), None, 100.0, None),
    ]),
    (["BAR-EN-105"], [
        (date(2022, 7, 1), date(2022, 12, 31), 7.5, 15.0),
        (date(2023, 1, 1), date(2023, 12, 31), 10.0, 20.0),
        (date(2024, 1, 1), date(2024, 12, 31), 12.5, 25.0),
        (date(2025, 1, 1), date(2026, 12, 31), 15.0, 30.0),
        (date(2027, 1, 1), None, 20.0, 30.0),
    ]),
    (["BAR-TH-127"], [
        (date(2023, 4, 1), date(2023, 12, 31), 10.0, 20.0),
        (date(2024, 1, 1), date(2024, 12, 31), 12.5, 25.0),
        (date(2025, 1, 1), date(2026, 12, 31), 15.0, 30.0),
        (date(2027, 1, 1), None, 20.0, 30.0),
    ]),
    (["BAR-TH-125"], [
        (date(2024, 7, 1), date(2024, 12, 31), 12.5, 25.0),
        (date(2025, 1, 1), date(2026, 12, 31), 15.0, 30.0),
        (date(2027, 1, 1), None, 20.0, 30.0),
    ]),
    (["BAR-TH-106", "BAR-TH-107", "BAR-TH-107-SE", "BAR-TH-118", "BAR-TH-158"], [
        (date(2023, 4, 1), date(2023, 12, 31), None, 20.0),
        (date(2024, 1, 1), date(2024, 12, 31), None, 25.0),
        (date(2025, 1, 1), None, None, 30.0),
    ]),
    (["BAR-TH-112"], [
        (date(2023, 7, 1), date(2023, 12, 31), None, 20.0),
        (date(2024, 1, 1), date(2024, 12, 31), None, 25.0),
        (date(2025, 1, 1), None, None, 30.0),
    ]),
    (["BAR-TH-173"], [
        (date(2024, 11, 22), date(2025, 6, 30), None, 80.0),
        (date(2025, 7, 1), date(2026, 12, 31), 15.0, 50.0),
        (date(2027, 1, 1), None, 20.0, 50.0),
    ]),
    (["BAR-EN-104"], [
        (date(2024, 1, 1), date(2024, 12, 31), None, 25.0),
        (date(2025, 1, 1), None, None, 30.0),
    ]),
    (["BAT-EN-101", "BAT-EN-102", "BAT-EN-103", "BAT-EN-106", "BAT-EN-108"], [
        (date(2022, 1, 1), date(2022, 12, 31), 7.5, 15.0),
        (date(2023, 1, 1), date(2023, 12, 31), 10.0, 20.0),
        (date(2024, 1, 1), date(2024, 12, 31), 12.5, 25.0),
        (date(2025, 1, 1), date(2026, 12, 31), 15.0, 30.0),
        (date(2027, 1, 1), None, 20.0, 30.0),
    ]),
    (["BAT-TH-139"], [
        (date(2022, 7, 1), date(2022, 12, 31), 7.5, 15.0),
        (date(2023, 1, 1), date(2023, 12, 31), 10.0, 20.0),
        (date(2024, 1, 1), date(2024, 12, 31), 12.5, 25.0),
        (date(2025, 1, 1), date(2026, 12, 31), 15.0, 30.0),
        (date(2027, 1, 1), None, 20.0, 30.0),
    ]),
    (["BAT-TH-157"], [
        (date(2023, 4, 1), date(2023, 12, 31), 10.0, 20.0),
        (date(2024, 1, 1), date(2024, 12, 31), 12.5, 25.0),
        (date(2025, 1, 1), date(2026, 12, 31), 15.0, 30.0),
        (date(2027, 1, 1), None, 20.0, 30.0),
    ]),
    (["BAT-TH-113"], [
        (date(2024, 7, 1), date(2024, 12, 31), 12.5, 25.0),
        (date(2025, 1, 1), date(2026, 12, 31), 15.0, 30.0),
        (date(2027, 1, 1), None, 20.0, 30.0),
    ]),
    (["BAT-TH-102", "BAT-EQ-127", "BAT-EQ-133"], [
        (date(2023, 4, 1), date(2023, 12, 31), None, 20.0),
        (date(2024, 1, 1), date(2024, 12, 31), None, 25.0),
        (date(2025, 1, 1), None, None, 30.0),
    ]),
    (["IND-EN-101", "IND-EN-102", "IND-UT-131"], [
        (date(2022, 1, 1), date(2022, 12, 31), 7.5, 15.0),
        (date(2023, 1, 1), date(2023, 12, 31), 10.0, 20.0),
        (date(2024, 1, 1), date(2024, 12, 31), 12.5, 25.0),
        (date(2025, 1, 1), date(2026, 12, 31), 15.0, 30.0),
        (date(2027, 1, 1), None, 20.0, 30.0),
    ]),
    (["IND-UT-102", "IND-UT-116", "IND-UT-117", "IND-UT-129", "IND-BA-112"], [
        (date(2022, 7, 1), date(2022, 12, 31), 7.5, 15.0),
        (date(2023, 1, 1), date(2023, 12, 31), 10.0, 20.0),
        (date(2024, 1, 1), date(2024, 12, 31), 12.5, 25.0),
        (date(2025, 1, 1), date(2026, 12, 31), 15.0, 30.0),
        (date(2027, 1, 1), None, 20.0, 30.0),
    ]),
    (["IND-UT-134"], [
        (date(2023, 4, 1), date(2023, 12, 31), 10.0, 20.0),
        (date(2024, 1, 1), date(2024, 12, 31), 12.5, 25.0),
        (date(2025, 1, 1), date(2026, 12, 31), 15.0, 30.0),
        (date(2027, 1, 1), None, 20.0, 30.0),
    ]),
    (["TRA-SE-114", "TRA-SE-115"], [
        (date(2023, 1, 1), date(2023, 12, 31), None, 20.0),
        (date(2024, 1, 1), date(2024, 12, 31), None, 25.0),
        (date(2025, 1, 1), None, None, 30.0),
    ]),
    (["TRA-EQ-124"], [
        (date(2023, 4, 1), date(2023, 12, 31), 10.0, 20.0),
        (date(2024, 1, 1), date(2024, 12, 31), 12.5, 25.0),
        (date(2025, 1, 1), date(2026, 12, 31), 15.0, 30.0),
        (date(2027, 1, 1), None, 20.0, 30.0),
    ]),
    (["TRA-EQ-101", "TRA-EQ-107", "TRA-EQ-108"], [
        (date(2023, 4, 1), date(2023, 12, 31), None, 20.0),
        (date(2024, 1, 1), date(2024, 12, 31), None, 25.0),
        (date(2025, 1, 1), None, None, 30.0),
    ]),
    (["BAR-TH-160", "BAR-TH-161", "BAT-TH-146", "BAT-TH-155", "IND-UT-121", "RES-CH-108"], [
        (date(2023, 10, 1), None, 100.0, None),
    ]),
    (["BAT-TH-116"], [
        (date(2024, 1, 1), date(2024, 2, 29), None, 100.0),
        (date(2024, 3, 1), date(2024, 12, 31), 12.5, 25.0),
        (date(2025, 1, 1), date(2026, 12, 31), 15.0, 30.0),
        (date(2027, 1, 1), None, 20.0, 30.0),
    ]),
    (["RES-CH-106", "RES-CH-107", "RES-EC-104"], [
        (date(2024, 3, 1), None, 100.0, None),
    ]),
    (["IND-UT-137", "IND-UT-138", "IND-UT-139"], [
        (date(2025, 1, 1), None, 100.0, None),
    ]),
    (["BAR-TH-177"], [
        (date(2024, 11, 1), None, 100.0, None),
    ]),
    (["TRA-EQ-114", "TRA-EQ-128", "TRA-EQ-129"], [
        (date(2025, 6, 1), None, 100.0, None),
    ]),
    (["TRA-EQ-117"], [
        (date(2025, 6, 1), date(2026, 12, 31), 15.0, None),
        (date(2027, 1, 1), None, 20.0, None),
    ]),
    # TRA-EQ-130 : seuil différent selon le type de bénéficiaire (personne physique / collectivité,
    # État ou autre personne morale). Par défaut on retient le seuil "personne physique" (le plus
    # bas des deux) — à ajuster manuellement si le lot concerne des collectivités.
    (["TRA-EQ-130"], [
        (date(2025, 6, 1), date(2026, 12, 31), 15.0, None),
        (date(2027, 1, 1), None, 20.0, None),
    ]),
    (["IND-BA-110"], [
        (date(2025, 8, 1), None, 100.0, None),
    ]),
    (["BAT-TH-142"], [
        (date(2025, 8, 1), None, 100.0, None),
    ]),
    (["TRA-EQ-131"], [
        (date(2025, 9, 1), None, None, 75.0),
    ]),
    (["BAR-TH-143"], [
        (date(2026, 3, 1), date(2026, 12, 31), 15.0, 30.0),
        (date(2027, 1, 1), None, 20.0, 30.0),
    ]),
    (["BAR-TH-178", "BAR-TH-179", "BAR-TH-180", "BAT-TH-162", "BAT-TH-163", "BAT-TH-164"], [
        (date(2026, 5, 1), date(2026, 12, 31), 50.0, None),
        (date(2027, 1, 1), date(2027, 12, 31), 75.0, None),
        (date(2028, 1, 1), None, 100.0, None),
    ]),
]

# Index fiche -> liste de périodes, pour lookup direct.
SEUILS_PAR_FICHE = {}
for _fiches, _periodes in _RAW_TABLE:
    for _f in _fiches:
        SEUILS_PAR_FICHE[_f] = _periodes


def parse_date_fr(value):
    """Convertit une valeur de cellule Excel (str 'JJ/MM/AAAA', datetime, ou None) en date
    Python, ou None si non exploitable."""
    if value is None:
        return None
    if isinstance(value, date):
        return value if not hasattr(value, "date") else value.date()
    s = str(value).strip()
    if not s:
        return None
    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d"):
        try:
            from datetime import datetime
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def extract_fiche_code(raw):
    """Extrait le code fiche propre (ex : 'BAR-EN-101') d'une valeur brute de colonne
    'REFERENCE DE LA FICHE', qui peut porter un suffixe de variante interne
    (ex : 'BAR-EN-101_VA33_3' -> 'BAR-EN-101'). Retourne None si rien ne matche."""
    if not raw:
        return None
    m = re.match(r"^([A-Z]{2,5}-[A-Z]{2}-\d{3}(?:-SE)?)", str(raw).strip().upper())
    return m.group(1) if m else str(raw).strip().upper()


def get_seuils_fiche(fiche, date_engagement):
    """Retourne (seuil_site, seuil_contact) en % pour une fiche et une date d'engagement
    données, ou (None, None) si la fiche ou la période n'est pas trouvée dans la table."""
    if not fiche or not date_engagement:
        return None, None
    code = extract_fiche_code(fiche)
    periodes = SEUILS_PAR_FICHE.get(code)
    if not periodes:
        return None, None
    for debut, fin, seuil_site, seuil_contact in periodes:
        if date_engagement >= debut and (fin is None or date_engagement <= fin):
            return seuil_site, seuil_contact
    return None, None
