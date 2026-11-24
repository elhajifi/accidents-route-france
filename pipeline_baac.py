#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
DATA CHALLENGE - Accidents corporels de la route en France (BAAC 2023-2025)
===========================================================================
Pipeline complet : exploration -> diagnostic qualité -> anomalies -> nettoyage
-> transformation -> premiers indicateurs.

Utilisation
-----------
    python pipeline_baac.py --entree "base de données" --sortie sortie_baac

    --entree : dossier contenant les 12 CSV (caract-/lieux-/usagers-/vehicules-AAAA.csv)
    --sortie : dossier de résultats (créé s'il n'existe pas)
    --annees : (optionnel) ex. --annees 2023 2024 2025

Dépendances : pandas, numpy, matplotlib, openpyxl (python-dateutil est livré avec pandas)

Sorties
-------
    sortie_baac/
      donnees_propres/  accidents.csv  usagers.csv  vehicules.csv   (séparateur ; décimale ,)
      rapports/         01_exploration.xlsx  02_diagnostic_qualite.xlsx
                        03_journal_nettoyage.xlsx  04_indicateurs.xlsx
      figures/          *.png
"""
import argparse
import datetime as dt
import logging
import sys
from pathlib import Path

import numpy as np
import pandas as pd
import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
from dateutil.easter import easter  # noqa: E402

log = logging.getLogger("baac")

# =============================================================================
# 0. CONFIGURATION : référentiels (document de description des variables BAAC)
# =============================================================================
ANNEES = (2023, 2024, 2025)
TABLES = ("caract", "lieux", "usagers", "vehicules")

# Valeurs textuelles qui signifient "non renseigné" dans les fichiers bruts
MANQUANTS_TXT = {"", "-1", "N/A", "#VALEURMULTI"}

LABELS = {
    "lum": {1: "Plein jour", 2: "Crépuscule ou aube", 3: "Nuit sans éclairage public",
            4: "Nuit, éclairage public non allumé", 5: "Nuit, éclairage public allumé"},
    "agg": {1: "Hors agglomération", 2: "En agglomération"},
    "int": {1: "Hors intersection", 2: "Intersection en X", 3: "Intersection en T",
            4: "Intersection en Y", 5: "Intersection à plus de 4 branches", 6: "Giratoire",
            7: "Place", 8: "Passage à niveau", 9: "Autre intersection"},
    "atm": {1: "Normale", 2: "Pluie légère", 3: "Pluie forte", 4: "Neige - grêle",
            5: "Brouillard - fumée", 6: "Vent fort - tempête", 7: "Temps éblouissant",
            8: "Temps couvert", 9: "Autre"},
    "col": {1: "Deux véhicules - frontale", 2: "Deux véhicules - par l'arrière",
            3: "Deux véhicules - par le côté", 4: "Trois véhicules et + - en chaîne",
            5: "Trois véhicules et + - collisions multiples", 6: "Autre collision",
            7: "Sans collision"},
    "catr": {1: "Autoroute", 2: "Route nationale", 3: "Route départementale",
             4: "Voie communale", 5: "Hors réseau public", 6: "Parc de stationnement",
             7: "Route de métropole urbaine", 9: "Autre"},
    "circ": {1: "Sens unique", 2: "Bidirectionnelle", 3: "Chaussées séparées",
             4: "Voies d'affectation variable"},
    "surf": {1: "Normale", 2: "Mouillée", 3: "Flaques", 4: "Inondée", 5: "Enneigée",
             6: "Boue", 7: "Verglacée", 8: "Corps gras - huile", 9: "Autre"},
    "situ": {1: "Sur chaussée", 2: "Sur bande d'arrêt d'urgence", 3: "Sur accotement",
             4: "Sur trottoir", 5: "Sur piste cyclable", 6: "Sur autre voie spéciale",
             8: "Autres"},
    "grav": {1: "Indemne", 2: "Tué", 3: "Blessé hospitalisé", 4: "Blessé léger"},
    "catu": {1: "Conducteur", 2: "Passager", 3: "Piéton"},
    "sexe": {1: "Masculin", 2: "Féminin"},
    "trajet": {1: "Domicile-travail", 2: "Domicile-école", 3: "Courses-achats",
               4: "Utilisation professionnelle", 5: "Promenade-loisirs", 9: "Autre"},
}

# Domaines de valeurs valides (pour détecter les codes hors référentiel)
DOMAINES = {
    "caract": {c: set(LABELS[c]) for c in ("lum", "agg", "int", "atm", "col")},
    "lieux": {**{c: set(LABELS[c]) for c in ("catr", "circ", "surf", "situ")},
              "vosp": {0, 1, 2, 3}, "prof": {1, 2, 3, 4}, "plan": {1, 2, 3, 4},
              "infra": set(range(10)), "v1": {0, 1, 2, 3}},
    "usagers": {**{c: set(LABELS[c]) for c in ("grav", "catu", "sexe")},
                "trajet": set(LABELS["trajet"]) | {0}, "place": set(range(1, 11))},
    "vehicules": {"catv": set(range(0, 100)), "senc": {0, 1, 2, 3}, "motor": set(range(7)),
                  "choc": set(range(10)), "obsm": {0, 1, 2, 4, 5, 6, 9}, "obs": set(range(18))},
}

# Colonnes où "-1" signifie surtout "non applicable" (et non "oubli")
NA_STRUCTUREL = {
    "usagers": {"secu2", "secu3", "locp", "actp", "etatp"},   # piétons / équipements multiples
    "lieux": {"v1", "v2", "lartpc"},
}

# Regroupement des catégories de véhicules (catv)
GROUPE_CATV = {}
for _codes, _g in [((1, 80), "Vélo"), ((2, 30, 41), "Cyclomoteur / scooter <= 50 cm3"),
                   ((31, 32, 33, 34, 42, 43), "Moto / scooter > 50 cm3"),
                   ((3, 7), "Voiture"), ((10,), "Utilitaire"),
                   ((13, 14, 15, 16, 17), "Poids lourd"), ((37, 38), "Bus / autocar"),
                   ((39, 40), "Train / tramway"), ((50, 60), "EDP (trottinette...)"),
                   ((20, 21), "Engin agricole / spécial"), ((35, 36), "Quad"),
                   ((99,), "Autre")]:
    for _c in _codes:
        GROUPE_CATV[_c] = _g
CATV_NON_MOTORISES = {1, 50, 60, 80, 0, 99}

# Gravité : le codage n'est PAS ordinal -> on définit un rang explicite
RANG_GRAVITE = {2: 4, 3: 3, 4: 2, 1: 1}          # tué > hospitalisé > léger > indemne
LABEL_RANG = {4: "Mortel", 3: "Grave (hospitalisation)", 2: "Léger", 1: "Indemne"}

REGIONS = {
    "Auvergne-Rhône-Alpes": "01 03 07 15 26 38 42 43 63 69 73 74",
    "Bourgogne-Franche-Comté": "21 25 39 58 70 71 89 90",
    "Bretagne": "22 29 35 56",
    "Centre-Val de Loire": "18 28 36 37 41 45",
    "Corse": "2A 2B",
    "Grand Est": "08 10 51 52 54 55 57 67 68 88",
    "Hauts-de-France": "02 59 60 62 80",
    "Île-de-France": "75 77 78 91 92 93 94 95",
    "Normandie": "14 27 50 61 76",
    "Nouvelle-Aquitaine": "16 17 19 23 24 33 40 47 64 79 86 87",
    "Occitanie": "09 11 12 30 31 32 34 46 48 65 66 81 82",
    "Pays de la Loire": "44 49 53 72 85",
    "Provence-Alpes-Côte d'Azur": "04 05 06 13 83 84",
    "Guadeloupe": "971", "Martinique": "972", "Guyane": "973", "La Réunion": "974",
    "Mayotte": "976",
    "Autres collectivités d'outre-mer": "975 977 978 984 985 986 987 988",
}
DEP_VERS_REGION = {d: r for r, deps in REGIONS.items() for d in deps.split()}

# Emprises (lat_min, lat_max, lon_min, lon_max) servant à valider les coordonnées GPS
BBOX_DEFAUT_METRO = (41.0, 51.6, -5.6, 10.0)
BBOX_DEP = {
    "2A": (41.3, 43.1, 8.5, 9.7), "2B": (41.3, 43.1, 8.5, 9.7),
    "971": (15.8, 16.6, -61.9, -60.9), "972": (14.3, 14.95, -61.3, -60.7),
    "973": (2.0, 6.0, -55.0, -51.0), "974": (-21.5, -20.8, 55.1, 55.9),
    "976": (-13.1, -12.5, 44.9, 45.4), "975": (46.7, 47.2, -56.6, -56.1),
    "977": (17.8, 18.0, -63.2, -62.7), "978": (17.8, 18.2, -63.2, -62.9),
    "986": (-14.5, -13.0, -178.3, -176.0), "987": (-28.0, -7.0, -155.0, -134.0),
    "988": (-23.0, -19.5, 163.5, 168.5),
}

JOURS = ["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi", "Dimanche"]
MOIS = ["Janvier", "Février", "Mars", "Avril", "Mai", "Juin", "Juillet", "Août",
        "Septembre", "Octobre", "Novembre", "Décembre"]
TRANCHES_HORAIRES = ["Nuit (0h-5h)", "Matin (6h-9h)", "Journée (10h-15h)",
                     "Soir (16h-19h)", "Soirée (20h-23h)"]
TRANCHES_AGE = ["0-14", "15-17", "18-24", "25-34", "35-44", "45-54", "55-64", "65-74", "75+"]
VMA_VALIDES = set(range(5, 131, 5))   # vitesses maximales autorisées plausibles (km/h)


# =============================================================================
# Utilitaires
# =============================================================================
def num(s, moins1_manquant=True, virgule=False):
    """Texte -> numérique (float). '', 'N/A', '#VALEURMULTI' -> NaN ; '-1' -> NaN (option)."""
    if virgule:
        s = s.str.replace(",", ".", regex=False)
    x = pd.to_numeric(s, errors="coerce")
    return x.where(x != -1) if moins1_manquant else x


def entier(s, **kw):
    return num(s, **kw).astype("Int64")


def evaluer_coordonnees(lat, lon, dep):
    """Valide lat/long par rapport à l'emprise du département et tente une correction
    (inversion lat/long et/ou changement de signe). Retourne un DataFrame :
    lat_c, lon_c, hors_emprise (anomalie initiale), corrigee, invalide."""
    def bornes(i):
        return dep.map(lambda d: BBOX_DEP.get(d, BBOX_DEFAUT_METRO if len(d) == 2 else (-90, 90, -180, 180))[i]).astype(float)
    b0, b1, b2, b3 = (bornes(i) for i in range(4))

    def dans(la, lo):
        return la.between(b0, b1) & lo.between(b2, b3) & ~((la == 0) & (lo == 0))

    ok0 = dans(lat, lon)
    la_c, lo_c, resolu = lat.copy(), lon.copy(), ok0.copy()
    candidats = [(lat, -lon), (-lat, lon), (-lat, -lon), (lon, lat), (lon, -lat), (-lon, lat), (-lon, -lat)]
    for la, lo in candidats:
        m = ~resolu & dans(la, lo)
        la_c, lo_c = la_c.where(~m, la), lo_c.where(~m, lo)
        resolu = resolu | m
    return pd.DataFrame({"lat_c": la_c.where(resolu), "lon_c": lo_c.where(resolu), "hors_emprise": ~ok0,
                         "corrigee": resolu & ~ok0, "invalide": ~resolu})


def libelle(s, cle):
    """Code entier -> libellé texte (NaN si code absent du référentiel)."""
    return s.astype("float64").map({float(k): v for k, v in LABELS[cle].items()})


def pct(n, base):
    return round(100.0 * n / base, 2) if base else 0.0


def jours_feries(annee):
    p = easter(annee)
    fixes = [(1, 1), (5, 1), (5, 8), (7, 14), (8, 15), (11, 1), (11, 11), (12, 25)]
    mobiles = [p + dt.timedelta(days=1), p + dt.timedelta(days=39), p + dt.timedelta(days=50)]
    return {dt.date(annee, m, j) for m, j in fixes} | set(mobiles)


class Rapport:
    """Collecte de tableaux destinés à un classeur Excel."""

    def __init__(self):
        self.feuilles = {}

    def ajouter(self, nom, df):
        self.feuilles[nom[:31]] = df

    def ecrire(self, chemin):
        with pd.ExcelWriter(chemin, engine="openpyxl") as w:
            for nom, df in self.feuilles.items():
                df.to_excel(w, sheet_name=nom, index=not isinstance(df.index, pd.RangeIndex))
                ws = w.sheets[nom]
                for col in ws.columns:
                    larg = max(len(str(c.value)) if c.value is not None else 0 for c in col[:200])
                    ws.column_dimensions[col[0].column_letter].width = min(max(larg + 2, 8), 60)
                ws.freeze_panes = "A2"
        log.info("   -> %s", chemin)


# =============================================================================
# 1. CHARGEMENT
# =============================================================================
def trouver_fichier(dossier, table, annee):
    """Retrouve le CSV même si le nom varie (ex. caract-2024.csv, caracteristiques-2024.csv)."""
    prefixe = {"caract": "carac", "lieux": "lieux", "usagers": "usagers", "vehicules": "vehic"}[table]
    candidats = [p for p in dossier.glob("*.csv")
                 if p.name.lower().startswith(prefixe) and str(annee) in p.name]
    if not candidats:
        raise FileNotFoundError(f"Fichier introuvable : {table} {annee} dans {dossier}")
    return sorted(candidats)[0]


def lire_csv(chemin, annee):
    try:
        df = pd.read_csv(chemin, sep=";", dtype=str, keep_default_na=False, encoding="utf-8-sig")
    except UnicodeDecodeError:
        df = pd.read_csv(chemin, sep=";", dtype=str, keep_default_na=False, encoding="latin-1")
    # En-têtes : retire espaces / retours à la ligne parasites (ex. "id_vehicule\n" en 2025)
    df.columns = [c.replace("\n", "").replace("\r", "").strip() for c in df.columns]
    # Valeurs : espaces insécables -> espaces, trim
    for c in df.columns:
        df[c] = df[c].str.replace("\xa0", " ", regex=False).str.strip()
    # Identifiants : "155 680 557" -> "155680557"
    for c in ("id_usager", "id_vehicule"):
        if c in df.columns:
            df[c] = df[c].str.replace(" ", "", regex=False)
    df["annee_fichier"] = annee
    return df


def charger(dossier, annees):
    log.info("1. CHARGEMENT")
    brut = {}
    for t in TABLES:
        morceaux = []
        for a in annees:
            f = trouver_fichier(dossier, t, a)
            d = lire_csv(f, a)
            log.info("   %-10s %s : %7d lignes x %2d colonnes (%s)", t, a, len(d), d.shape[1] - 1, f.name)
            morceaux.append(d)
        cols = [set(m.columns) for m in morceaux]
        if any(c != cols[0] for c in cols):
            log.warning("   /!\\ colonnes différentes selon les années pour %s", t)
        brut[t] = pd.concat(morceaux, ignore_index=True)
    return brut


# =============================================================================
# 2. EXPLORATION APPROFONDIE
# =============================================================================
def explorer(brut, rap):
    log.info("2. EXPLORATION")
    volumes, profils = [], []
    for t, df in brut.items():
        for a, g in df.groupby("annee_fichier"):
            volumes.append({"table": t, "annee": a, "lignes": len(g), "colonnes": g.shape[1] - 1})
        for c in df.columns.drop("annee_fichier"):
            s = df[c]
            n = len(s)
            vides, m1 = (s == "").sum(), (s == "-1").sum()
            na, multi = (s == "N/A").sum(), (s == "#VALEURMULTI").sum()
            manq = vides + m1 + na + multi
            vc = s.value_counts()
            top = "; ".join(f"{k or '(vide)'}: {v}" for k, v in vc.head(5).items())
            if c in NA_STRUCTUREL.get(t, ()):
                interp = "'-1' = surtout non applicable"
            elif pct(manq, n) > 50:
                interp = "colonne peu exploitable (>50 % manquant)"
            else:
                interp = ""
            profils.append({"table": t, "colonne": c, "nb_distincts": s.nunique(),
                            "vides": vides, "moins_1": m1, "N/A": na, "#VALEURMULTI": multi,
                            "pct_manquant": pct(manq, n), "interpretation": interp,
                            "top5_valeurs": top})
    vol = pd.DataFrame(volumes)
    prof = pd.DataFrame(profils)
    log.info("   volumes par table et année :\n%s", vol.pivot(index="table", columns="annee", values="lignes"))
    rap.ajouter("volumes", vol)
    rap.ajouter("profil_colonnes", prof)

    # Rapport des cardinalités : lignes / accident dans chaque table
    n_acc = brut["caract"]["Num_Acc"].nunique()
    card = pd.DataFrame([{"table": t, "lignes": len(brut[t]),
                          "accidents_distincts": brut[t]["Num_Acc"].nunique(),
                          "lignes_par_accident": round(len(brut[t]) / max(brut[t]["Num_Acc"].nunique(), 1), 2)}
                         for t in TABLES])
    log.info("   %d accidents au total dans caract", n_acc)
    rap.ajouter("cardinalite", card)

    # Couverture temporelle par année (détecte une année incomplète, ex. 2025)
    c = brut["caract"]
    d = pd.to_datetime({"year": pd.to_numeric(c["an"]), "month": pd.to_numeric(c["mois"]),
                        "day": pd.to_numeric(c["jour"])}, errors="coerce")
    cov = pd.DataFrame({"annee": c["annee_fichier"], "date": d}).groupby("annee").agg(
        date_min=("date", "min"), date_max=("date", "max"),
        mois_distincts=("date", lambda s: s.dt.month.nunique()), accidents=("date", "size"))
    log.info("   couverture temporelle :\n%s", cov)
    rap.ajouter("couverture_temporelle", cov.reset_index())


# =============================================================================
# 3. DIAGNOSTIC QUALITÉ + 4. ANOMALIES
# =============================================================================
def diagnostiquer(brut, rap):
    log.info("3-4. DIAGNOSTIC QUALITÉ ET ANOMALIES")
    car, lie, usa, veh = (brut[t] for t in TABLES)
    res = []

    def ajouter(table, controle, n, base, traitement, gravite="Moyenne"):
        res.append({"table": table, "controle": controle, "nb_anomalies": int(n),
                    "pct_de_la_base": pct(n, base), "gravite": gravite, "traitement": traitement})

    # --- Doublons -----------------------------------------------------------
    ajouter("caract", "Num_Acc en double", car["Num_Acc"].duplicated().sum(), len(car), "Suppression", "Haute")
    ajouter("usagers", "id_usager en double", usa["id_usager"].duplicated().sum(), len(usa), "Suppression", "Haute")
    ajouter("vehicules", "id_vehicule en double", veh["id_vehicule"].duplicated().sum(), len(veh), "Suppression", "Haute")
    for t, d in brut.items():
        ajouter(t, "Lignes entièrement dupliquées", d.drop(columns="annee_fichier").duplicated().sum(),
                len(d), "Suppression", "Haute")
    ajouter("lieux", "Accidents avec plusieurs lignes 'lieux' (routes multiples)",
            (lie.groupby("Num_Acc").size() > 1).sum(), car["Num_Acc"].nunique(),
            "Conservation de la 1re ligne + indicateur nb_voies_lieux", "Info")

    # --- Intégrité référentielle -------------------------------------------
    acc = set(car["Num_Acc"])
    for t, d in (("lieux", lie), ("usagers", usa), ("vehicules", veh)):
        ajouter(t, "Num_Acc absent de 'caract' (orphelin)", (~d["Num_Acc"].isin(acc)).sum(), len(d),
                "Suppression", "Haute")
        ajouter("caract", f"Accident sans ligne dans '{t}'", (~car["Num_Acc"].isin(set(d["Num_Acc"]))).sum(),
                len(car), "Signalement", "Haute")
    cle_veh = set(zip(veh["Num_Acc"], veh["id_vehicule"]))
    ajouter("usagers", "Usager sans véhicule correspondant (Num_Acc, id_vehicule)",
            sum(k not in cle_veh for k in zip(usa["Num_Acc"], usa["id_vehicule"])), len(usa),
            "Signalement", "Haute")

    # --- Formats / dates ----------------------------------------------------
    ajouter("caract", "Num_Acc != 12 chiffres", (~car["Num_Acc"].str.fullmatch(r"\d{12}")).sum(), len(car), "Suppression")
    ajouter("caract", "Préfixe Num_Acc != année de l'accident", (car["Num_Acc"].str[:4] != car["an"]).sum(), len(car), "Signalement")
    ajouter("caract", "Année 'an' != année du fichier", (car["an"] != car["annee_fichier"].astype(str)).sum(), len(car), "Signalement")
    d = pd.to_datetime({"year": pd.to_numeric(car["an"], errors="coerce"),
                        "month": pd.to_numeric(car["mois"], errors="coerce"),
                        "day": pd.to_numeric(car["jour"], errors="coerce")}, errors="coerce")
    ajouter("caract", "Date invalide", d.isna().sum(), len(car), "Suppression", "Haute")
    h = pd.to_numeric(car["hrmn"].str[:2], errors="coerce")
    m = pd.to_numeric(car["hrmn"].str[3:5], errors="coerce")
    ajouter("caract", "Heure invalide (format HH:MM)", ((~car["hrmn"].str.fullmatch(r"\d{2}:\d{2}")) | (h > 23) | (m > 59)).sum(),
            len(car), "Mise à NaN")
    lum = pd.to_numeric(car["lum"], errors="coerce")
    ajouter("caract", "Incohérence luminosité/heure : 'plein jour' entre 0h et 4h59", ((lum == 1) & (h <= 4)).sum(),
            len(car), "Signalement (indicateur)", "Faible")
    ajouter("caract", "Incohérence luminosité/heure : 'nuit' entre 11h et 15h59", (lum.isin([3, 4, 5]) & h.between(11, 15)).sum(),
            len(car), "Signalement (indicateur)", "Faible")

    # --- Valeurs manquantes sur les variables clés -------------------------------
    cles = {"caract": ["lum", "atm", "col", "int", "agg", "com", "dep", "adr"],
            "lieux": ["catr", "surf", "circ", "nbv", "vma", "prof", "plan", "situ", "infra"],
            "usagers": ["grav", "sexe", "an_nais", "trajet", "place"],
            "vehicules": ["catv", "senc", "motor", "choc", "manv", "obs"]}
    for t, cols in cles.items():
        d_ = brut[t]
        for c in cols:
            manq = d_[c].isin(MANQUANTS_TXT).sum()
            ajouter(t, f"Valeur manquante codée : {c}", manq, len(d_), "Mise à NaN (non imputée)",
                    "Haute" if pct(manq, len(d_)) > 5 else "Faible")

    # --- Codes hors référentiel ----------------------------------------------------
    for t, doms in DOMAINES.items():
        for c, valides in doms.items():
            x = pd.to_numeric(brut[t][c], errors="coerce")
            inval = (~brut[t][c].isin(MANQUANTS_TXT)) & (~x.isin(valides))
            ajouter(t, f"Code hors référentiel : {c}", inval.sum(), len(brut[t]), "Mise à NaN")
    ajouter("lieux", "nbv = '#VALEURMULTI' (artefact Excel)", (lie["nbv"] == "#VALEURMULTI").sum(), len(lie), "Mise à NaN")
    vma = pd.to_numeric(lie["vma"], errors="coerce")
    ajouter("lieux", "VMA invalide (hors 5-130 km/h ou non multiple de 5)",
            (vma.notna() & (vma != -1) & ~vma.isin(VMA_VALIDES)).sum(), len(lie), "Mise à NaN")
    ajouter("lieux", "pr / pr1 avec séparateur de milliers ('1 561')",
            lie["pr"].str.contains(" ").sum() + lie["pr1"].str.contains(" ").sum(), len(lie) * 2, "Nettoyage du format", "Faible")

    # --- Géographie ------------------------------------------------------------------
    lat, lon = num(car["lat"], False, True), num(car["long"], False, True)
    metro = car["dep"].str.len() == 2
    ajouter("caract", "Coordonnées manquantes ou nulles", (lat.isna() | lon.isna() | ((lat == 0) & (lon == 0))).sum(), len(car), "Mise à NaN")
    ev = evaluer_coordonnees(lat, lon, car["dep"])
    ajouter("caract", "Coordonnées hors de l'emprise du département", ev["hors_emprise"].sum(), len(car), "Correction ou mise à NaN", "Haute")
    ajouter("caract", "  dont corrigeables (inversion lat/long ou signe)", ev["corrigee"].sum(), len(car), "Correction automatique", "Haute")
    ajouter("caract", "  dont non corrigeables", ev["invalide"].sum(), len(car), "Mise à NaN", "Haute")
    ajouter("caract", "  dont en outre-mer", (ev["hors_emprise"] & ~metro).sum(), len(car), "Voir ci-dessus", "Info")
    ajouter("caract", "Commune ne commençant pas par le code département (métropole)",
            (metro & ~car.apply(lambda r: r["com"].startswith(r["dep"]), axis=1)).sum(), len(car), "Signalement")
    ajouter("caract", "Département hors table de régions", (~car["dep"].isin(DEP_VERS_REGION)).sum(), len(car), "Région = NaN")

    # --- Usagers : âge, cohérence ---------------------------------------------------
    an_acc = car.drop_duplicates("Num_Acc").set_index("Num_Acc")["an"]
    annee = pd.to_numeric(usa["Num_Acc"].map(an_acc), errors="coerce")
    naiss = num(usa["an_nais"], False)
    age = annee - naiss
    ajouter("usagers", "Année de naissance manquante", naiss.isna().sum(), len(usa), "Mise à NaN", "Moyenne")
    ajouter("usagers", "Âge négatif (naissance postérieure à l'accident)", (age < 0).sum(), len(usa), "Mise à NaN")
    ajouter("usagers", "Âge > 100 ans", (age > 100).sum(), len(usa), "Mise à NaN")
    catv = pd.to_numeric(usa["id_vehicule"].map(veh.drop_duplicates("id_vehicule").set_index("id_vehicule")["catv"]), errors="coerce")
    inco = (usa["catu"] == "1") & (age < 14) & ~catv.isin(CATV_NON_MOTORISES)
    ajouter("usagers", "Conducteur de véhicule motorisé de moins de 14 ans", inco.sum(), len(usa), "Indicateur (conservé)", "Faible")
    ajouter("usagers", "Piéton (catu=3) avec place != 10", ((usa["catu"] == "3") & (usa["place"] != "10")).sum(), len(usa), "Signalement", "Faible")

    # --- Gravité par accident ---------------------------------------------------------
    gv = pd.to_numeric(usa["grav"], errors="coerce")
    ok = gv.isin([1, 2, 3, 4])
    ajouter("caract", "Accident sans aucune gravité usager exploitable",
            len(acc - set(usa.loc[ok, "Num_Acc"])), len(acc), "Gravité accident = NaN", "Haute")
    nb_us = usa.groupby("Num_Acc").size()
    ajouter("usagers", "Accidents atypiques (>= 20 usagers)", (nb_us >= 20).sum(), len(acc), "Conservés (valeurs extrêmes réelles)", "Info")

    df = pd.DataFrame(res)
    df["ordre"] = df["gravite"].map({"Haute": 0, "Moyenne": 1, "Faible": 2, "Info": 3})
    df = df.sort_values(["ordre", "nb_anomalies"], ascending=[True, False]).drop(columns="ordre")
    rap.ajouter("anomalies", df)
    top = df[df["nb_anomalies"] > 0].head(12)
    log.info("   %d contrôles, %d avec anomalies. Principaux :\n%s", len(df), (df["nb_anomalies"] > 0).sum(),
             top[["table", "controle", "nb_anomalies", "pct_de_la_base"]].to_string(index=False))


# =============================================================================
# 5. NETTOYAGE
# =============================================================================
class Journal:
    def __init__(self):
        self.lignes = []

    def __call__(self, table, action, n, detail=""):
        self.lignes.append({"table": table, "action": action, "nb_lignes_ou_valeurs": int(n), "detail": detail})
        log.info("   [%-9s] %-62s %8d", table, action, n)

    def df(self):
        return pd.DataFrame(self.lignes)


def nettoyer(brut, J):
    log.info("5. NETTOYAGE")
    # ---------- caract ----------
    c = brut["caract"].copy()
    n = c.duplicated(subset="Num_Acc").sum()
    c = c.drop_duplicates(subset="Num_Acc", keep="first")
    J("caract", "Suppression des Num_Acc en double", n)
    c["date"] = pd.to_datetime({"year": pd.to_numeric(c["an"], errors="coerce"),
                                "month": pd.to_numeric(c["mois"], errors="coerce"),
                                "day": pd.to_numeric(c["jour"], errors="coerce")}, errors="coerce")
    n = c["date"].isna().sum()
    c = c[c["date"].notna()].copy()
    J("caract", "Suppression des accidents à date invalide", n)
    c["annee"] = c["date"].dt.year.astype(int)
    c["heure"] = pd.to_numeric(c["hrmn"].str[:2], errors="coerce").astype("Int64")
    c["minute"] = pd.to_numeric(c["hrmn"].str[3:5], errors="coerce").astype("Int64")
    bad = (c["heure"] > 23) | (c["minute"] > 59)
    c.loc[bad.fillna(False), ["heure", "minute"]] = pd.NA
    J("caract", "Heures invalides mises à NaN", bad.fillna(False).sum())
    for col in ("lum", "agg", "int", "atm", "col"):
        manq = c[col].isin(MANQUANTS_TXT).sum()
        c[col] = entier(c[col])
        J("caract", f"{col} : '-1' converti en NaN", manq)
    c["dep"] = c["dep"].str.upper().str.zfill(2)
    c["com"] = c["com"].str.zfill(5)
    c["adr"] = c["adr"].where(~c["adr"].isin(["", "N/A"]))
    # coordonnées : validation par département + correction (signe / inversion)
    ev = evaluer_coordonnees(num(c["lat"], False, True), num(c["long"], False, True), c["dep"])
    c["lat"], c["long"] = ev["lat_c"], ev["lon_c"]
    c["coord_valide"] = ~ev["invalide"]
    c["coord_corrigee"] = ev["corrigee"]
    J("caract", "Coordonnées corrigées (inversion lat/long ou signe)", ev["corrigee"].sum(), "Corse, Polynésie, Mayotte...")
    J("caract", "Coordonnées non corrigeables mises à NaN", ev["invalide"].sum())
    c = c.drop(columns=["jour", "mois", "an", "hrmn", "annee_fichier"])

    # ---------- lieux ----------
    lx = brut["lieux"].copy()
    n = lx.drop(columns="annee_fichier").duplicated().sum()
    lx = lx.drop_duplicates(subset=lx.columns.drop("annee_fichier"))
    J("lieux", "Suppression des lignes entièrement dupliquées", n)
    n = (~lx["Num_Acc"].isin(c["Num_Acc"])).sum()
    lx = lx[lx["Num_Acc"].isin(c["Num_Acc"])]
    J("lieux", "Suppression des lignes orphelines", n)
    for col in ("catr", "v1", "circ", "vosp", "prof", "plan", "surf", "infra", "situ"):
        lx[col] = entier(lx[col])
    multi = (lx["nbv"] == "#VALEURMULTI").sum()
    lx["nbv"] = entier(lx["nbv"])
    J("lieux", "nbv '#VALEURMULTI' converti en NaN", multi)
    for col in ("pr", "pr1"):
        lx[col] = entier(lx[col].str.replace(" ", "", regex=False))
    lx["lartpc"] = num(lx["lartpc"], False, True)
    lx["larrout"] = num(lx["larrout"], True, True)
    vma = num(lx["vma"])
    inval = vma.notna() & ~vma.isin(VMA_VALIDES)
    lx["vma"] = vma.where(~inval).astype("Int64")
    J("lieux", "VMA invalides mises à NaN", inval.sum(), "valeurs hors 5-130 km/h ou non multiples de 5")
    lx["voie"] = lx["voie"].where(~lx["voie"].isin(["", "N/A"]))
    # une ligne par accident (1re ligne = route principale), avec indicateurs d'hétérogénéité
    g = lx.groupby("Num_Acc", sort=False)
    meta = pd.DataFrame({"nb_voies_lieux": g.size(),
                         "lieux_heterogenes": (g["surf"].nunique(dropna=False) > 1) | (g["vma"].nunique(dropna=False) > 1)})
    lx = lx.drop_duplicates("Num_Acc", keep="first").merge(meta, left_on="Num_Acc", right_index=True)
    J("lieux", "Agrégation à 1 ligne par accident", len(meta), f"{(meta['nb_voies_lieux'] > 1).sum()} accidents multi-voies")
    lx = lx[["Num_Acc", "catr", "voie", "circ", "nbv", "vosp", "prof", "plan", "surf", "infra", "situ",
             "vma", "larrout", "nb_voies_lieux", "lieux_heterogenes"]]

    # ---------- usagers ----------
    u = brut["usagers"].copy()
    n = u.duplicated(subset="id_usager").sum()
    u = u.drop_duplicates("id_usager")
    J("usagers", "Suppression des id_usager en double", n)
    n = (~u["Num_Acc"].isin(c["Num_Acc"])).sum()
    u = u[u["Num_Acc"].isin(c["Num_Acc"])]
    J("usagers", "Suppression des usagers orphelins", n)
    for col in ("place", "catu", "grav", "sexe", "secu1", "secu2", "secu3", "locp", "etatp"):
        u[col] = entier(u[col])
    t0 = u["trajet"].isin(["-1", "0", ""]).sum()
    u["trajet"] = entier(u["trajet"])
    u.loc[u["trajet"] == 0, "trajet"] = pd.NA
    J("usagers", "trajet : '-1' et '0' (non renseigné) -> NaN", t0)
    u["actp"] = u["actp"].where(~u["actp"].isin(["", "-1"]))
    u["an_nais"] = num(u["an_nais"], False)
    J("usagers", "grav '-1' -> NaN (usagers sans gravité)", u["grav"].isna().sum())
    u = u.drop(columns="annee_fichier")

    # ---------- vehicules ----------
    v = brut["vehicules"].copy()
    n = v.duplicated(subset="id_vehicule").sum()
    v = v.drop_duplicates("id_vehicule")
    J("vehicules", "Suppression des id_vehicule en double", n)
    n = (~v["Num_Acc"].isin(c["Num_Acc"])).sum()
    v = v[v["Num_Acc"].isin(c["Num_Acc"])]
    J("vehicules", "Suppression des véhicules orphelins", n)
    for col in ("senc", "catv", "obs", "obsm", "choc", "manv", "motor", "occutc"):
        v[col] = entier(v[col])
    for col in ("catv", "motor", "senc"):
        z = (v[col] == 0).sum()
        v.loc[v[col] == 0, col] = pd.NA
        J("vehicules", f"{col} : code 0 ('inconnu/indéterminable') -> NaN", z)
    v = v.drop(columns="annee_fichier")
    return c, lx, u, v


# =============================================================================
# 6. TRANSFORMATION DES VARIABLES
# =============================================================================
def transformer(c, lx, u, v, J):
    log.info("6. TRANSFORMATION")
    # ---------- véhicules ----------
    v["groupe_vehicule"] = v["catv"].astype("float64").map({float(k): g for k, g in GROUPE_CATV.items()})
    v["groupe_vehicule"] = v["groupe_vehicule"].fillna("Non déterminé")
    v["catv_libelle_groupe"] = v["groupe_vehicule"]
    v = v.drop(columns="catv_libelle_groupe")

    # ---------- usagers ----------
    u = u.merge(c[["Num_Acc", "annee"]], on="Num_Acc", how="left")
    age = u["annee"] - u["an_nais"]
    inco = (age < 0) | (age > 100)
    u["age"] = age.where(~inco).astype("Int64")
    J("usagers", "Âge calculé (annee - an_nais) ; âges <0 ou >100 -> NaN", inco.sum())
    u = u.merge(v[["id_vehicule", "catv", "groupe_vehicule"]], on="id_vehicule", how="left")
    u["age_incoherent_conducteur"] = ((u["catu"] == 1) & (u["age"] < 14)
                                     & ~u["catv"].isin(CATV_NON_MOTORISES)).fillna(False)
    u["tranche_age"] = pd.cut(u["age"].astype("float64"), [-1, 14, 17, 24, 34, 44, 54, 64, 74, 200],
                              labels=TRANCHES_AGE)
    u["categorie_usager"] = libelle(u["catu"], "catu")
    u["sexe_libelle"] = libelle(u["sexe"], "sexe")
    u["gravite"] = libelle(u["grav"], "grav")
    u["trajet_libelle"] = libelle(u["trajet"], "trajet")
    u["tue"] = u["grav"].isin([2])
    u["hospitalise"] = u["grav"].isin([3])
    u["blesse_leger"] = u["grav"].isin([4])
    u["indemne"] = u["grav"].isin([1])
    u["grave"] = u["tue"] | u["hospitalise"]                       # tué ou hospitalisé
    u["rang_gravite"] = u["grav"].astype("float64").map({float(k): r for k, r in RANG_GRAVITE.items()})
    u["mode_transport"] = np.where(u["catu"] == 3, "Piéton", u["groupe_vehicule"])
    u.loc[u["catu"].isna() & u["groupe_vehicule"].isna(), "mode_transport"] = "Non déterminé"

    # ---------- accidents (1 ligne = 1 accident) ----------
    a = c.merge(lx, on="Num_Acc", how="left")
    d = a["date"]
    a["mois_num"] = d.dt.month
    a["mois_nom"] = pd.Categorical(d.dt.month.map(dict(enumerate(MOIS, 1))), categories=MOIS, ordered=True)
    a["jour_semaine"] = pd.Categorical(d.dt.dayofweek.map(dict(enumerate(JOURS))), categories=JOURS, ordered=True)
    a["weekend"] = d.dt.dayofweek >= 5
    feries = set().union(*(jours_feries(y) for y in a["annee"].unique()))
    a["jour_ferie"] = d.dt.date.isin(feries)
    a["saison"] = d.dt.month.map({12: "Hiver", 1: "Hiver", 2: "Hiver", 3: "Printemps", 4: "Printemps",
                                  5: "Printemps", 6: "Été", 7: "Été", 8: "Été", 9: "Automne",
                                  10: "Automne", 11: "Automne"})
    h = a["heure"].astype("float64")
    a["tranche_horaire"] = pd.Categorical(
        pd.cut(h, [-1, 5, 9, 15, 19, 23], labels=TRANCHES_HORAIRES), categories=TRANCHES_HORAIRES, ordered=True)
    a["periode_lumiere"] = a["lum"].astype("float64").map({1: "Jour", 2: "Crépuscule/aube", 3: "Nuit", 4: "Nuit", 5: "Nuit"})
    a["incoherence_lum_heure"] = (((a["lum"] == 1) & (a["heure"] <= 4)) | (a["lum"].isin([3, 4, 5]) & a["heure"].between(11, 15))).fillna(False)

    for col in ("lum", "agg", "int", "atm", "col", "catr", "circ", "surf", "situ"):
        a[f"{col}_libelle"] = libelle(a[col], col)
    a["meteo_groupe"] = a["atm"].astype("float64").map({1: "Normale", 8: "Normale (couvert)", 2: "Pluie", 3: "Pluie",
                                                        4: "Neige / grêle", 5: "Brouillard / fumée", 6: "Vent fort",
                                                        7: "Éblouissant", 9: "Autre"})
    a["region"] = a["dep"].map(DEP_VERS_REGION)
    a["metropole"] = a["dep"].str.len() == 2
    a["vma_classe"] = pd.cut(a["vma"].astype("float64"), [0, 30, 50, 80, 90, 110, 130],
                             labels=["<=30", "31-50", "51-80", "81-90", "91-110", "111-130"])

    # agrégats usagers -> accident
    g = u.groupby("Num_Acc")
    ag = pd.DataFrame({"nb_usagers": g.size(), "nb_tues": g["tue"].sum(), "nb_hospitalises": g["hospitalise"].sum(),
                       "nb_blesses_legers": g["blesse_leger"].sum(), "nb_indemnes": g["indemne"].sum(),
                       "nb_pietons": (u["catu"] == 3).groupby(u["Num_Acc"]).sum(),
                       "rang_max": g["rang_gravite"].max()})
    gv = v.groupby("Num_Acc")
    ag["nb_vehicules"] = gv.size()
    for nom, grp in (("implique_velo", ["Vélo"]), ("implique_2rm", ["Cyclomoteur / scooter <= 50 cm3", "Moto / scooter > 50 cm3"]),
                     ("implique_pl", ["Poids lourd"]), ("implique_edp", ["EDP (trottinette...)"]),
                     ("implique_voiture", ["Voiture"])):
        ag[nom] = v["groupe_vehicule"].isin(grp).groupby(v["Num_Acc"]).any()
    a = a.merge(ag, left_on="Num_Acc", right_index=True, how="left")
    for col in ("nb_usagers", "nb_tues", "nb_hospitalises", "nb_blesses_legers", "nb_indemnes", "nb_pietons", "nb_vehicules"):
        a[col] = a[col].fillna(0).astype(int)
    for col in ("implique_velo", "implique_2rm", "implique_pl", "implique_edp", "implique_voiture"):
        a[col] = a[col].fillna(False).astype(bool)
    a["implique_pieton"] = a["nb_pietons"] > 0
    a["gravite_accident"] = a["rang_max"].map(LABEL_RANG)
    a["accident_mortel"] = a["nb_tues"] > 0
    a["accident_grave"] = (a["nb_tues"] + a["nb_hospitalises"]) > 0       # tué ou hospitalisé
    a["nb_blesses"] = a["nb_hospitalises"] + a["nb_blesses_legers"]
    a["gravite_inconnue"] = a["rang_max"].isna()
    a = a.drop(columns=["rang_max"])

    # variables accident -> usagers (analyses croisées)
    u = u.merge(a[["Num_Acc", "date", "dep", "region", "metropole", "lum_libelle", "meteo_groupe", "catr_libelle",
                   "surf_libelle", "agg_libelle", "jour_semaine", "tranche_horaire"]], on="Num_Acc", how="left")
    J("accidents", "Table accident construite", len(a), f"{a.shape[1]} colonnes")
    J("usagers", "Table usagers enrichie", len(u), f"{u.shape[1]} colonnes")
    J("vehicules", "Table véhicules enrichie", len(v), f"{v.shape[1]} colonnes")
    return a, u, v


# =============================================================================
# 7. PREMIERS INDICATEURS
# =============================================================================
def synthese(a, par, tri=None, min_n=0):
    """Indicateurs de fréquence et de gravité par modalité (niveau accident)."""
    b = a.copy()
    cols = [par] if isinstance(par, str) else list(par)
    for col in cols:
        if str(b[col].dtype) == "category":
            b[col] = b[col].cat.add_categories("Non renseigné").fillna("Non renseigné") if "Non renseigné" not in b[col].cat.categories else b[col].fillna("Non renseigné")
        else:
            b[col] = b[col].astype(object).where(b[col].notna(), "Non renseigné")
    g = b.groupby(par, observed=True)
    o = g.agg(nb_accidents=("Num_Acc", "size"), accidents_mortels=("accident_mortel", "sum"),
              accidents_graves=("accident_grave", "sum"), nb_tues=("nb_tues", "sum"),
              nb_hospitalises=("nb_hospitalises", "sum"), nb_usagers=("nb_usagers", "sum"))
    o["pct_accidents"] = (100 * o["nb_accidents"] / len(b)).round(2)
    o["pct_accidents_mortels"] = (100 * o["accidents_mortels"] / o["nb_accidents"]).round(2)
    o["pct_accidents_graves"] = (100 * o["accidents_graves"] / o["nb_accidents"]).round(2)
    o["tues_pour_100_accidents"] = (100 * o["nb_tues"] / o["nb_accidents"]).round(2)
    glob = 100 * b["accident_grave"].mean()
    o["indice_gravite"] = (o["pct_accidents_graves"] / glob).round(2)     # >1 = sur-représentation de la gravité
    if min_n:
        o = o[o["nb_accidents"] >= min_n]
    if tri == "freq":
        o = o.sort_values("nb_accidents", ascending=False)
    return o


def synthese_usagers(u, par):
    b = u.copy()
    for col in ([par] if isinstance(par, str) else par):
        if str(b[col].dtype) == "category":
            b[col] = b[col].cat.add_categories("Non renseigné").fillna("Non renseigné")
        else:
            b[col] = b[col].astype(object).where(b[col].notna(), "Non renseigné")
    o = b.groupby(par, observed=True).agg(nb_usagers=("id_usager", "size"), nb_tues=("tue", "sum"),
                                          nb_hospitalises=("hospitalise", "sum"), nb_graves=("grave", "sum"))
    o["pct_usagers"] = (100 * o["nb_usagers"] / len(b)).round(2)
    o["pct_tues"] = (100 * o["nb_tues"] / o["nb_usagers"]).round(2)
    o["pct_graves"] = (100 * o["nb_graves"] / o["nb_usagers"]).round(2)
    return o


def fig_barh(df, col, titre, fichier, dossier, min_n=100, top=None):
    d = df[df["nb_accidents"] >= min_n].sort_values("pct_accidents_graves")
    if top:
        d = d.tail(top)
    fig, ax = plt.subplots(figsize=(9, max(3, 0.42 * len(d) + 1.2)))
    ax.barh(d.index.astype(str), d["pct_accidents_graves"], color="#c0392b")
    glob_line = None
    for i, (idx, r) in enumerate(d.iterrows()):
        ax.text(r["pct_accidents_graves"] + 0.3, i, f"{r['pct_accidents_graves']:.1f} %  (n={int(r['nb_accidents']):,})".replace(",", " "),
                va="center", fontsize=8)
    ax.set_xlabel("% d'accidents graves (tué ou hospitalisé)")
    ax.set_title(titre, fontsize=11, weight="bold")
    ax.set_xlim(0, d["pct_accidents_graves"].max() * 1.35)
    plt.tight_layout()
    fig.savefig(dossier / fichier, dpi=130)
    plt.close(fig)


def indicateurs(a, u, v, rap, dossier_fig):
    log.info("7. INDICATEURS")
    # ---- KPI globaux ----
    def kpi(d, uu, vv):
        return {"accidents": len(d), "usagers": len(uu), "vehicules": len(vv),
                "tues": int(d["nb_tues"].sum()), "hospitalises": int(d["nb_hospitalises"].sum()),
                "blesses_legers": int(d["nb_blesses_legers"].sum()),
                "pct_accidents_graves": round(100 * d["accident_grave"].mean(), 2),
                "pct_accidents_mortels": round(100 * d["accident_mortel"].mean(), 2),
                "tues_pour_100_accidents": round(100 * d["nb_tues"].sum() / len(d), 2),
                "usagers_par_accident": round(len(uu) / len(d), 2),
                "vehicules_par_accident": round(len(vv) / len(d), 2)}
    rows = {str(y): kpi(a[a["annee"] == y], u[u["annee"] == y], v[v["Num_Acc"].isin(a.loc[a["annee"] == y, "Num_Acc"])])
            for y in sorted(a["annee"].unique())}
    rows["Total"] = kpi(a, u, v)
    k = pd.DataFrame(rows)
    log.info("   KPI :\n%s", k.to_string())
    rap.ajouter("kpi_globaux", k.reset_index().rename(columns={"index": "indicateur"}))

    # ---- Dimension temporelle ----
    t_annee = synthese(a, "annee")
    t_mois = synthese(a, "mois_nom")
    t_mois_annee = a.groupby(["annee", "mois_num"]).size().unstack("annee")
    t_jour = synthese(a, "jour_semaine")
    t_heure = synthese(a, "heure")
    t_tranche = synthese(a, "tranche_horaire")
    t_saison = synthese(a, "saison")
    t_we = synthese(a, "weekend")
    t_ferie = synthese(a, "jour_ferie")
    heat = a.pivot_table(index="jour_semaine", columns="heure", values="Num_Acc", aggfunc="size", observed=True)
    for nom, df in (("temps_annee", t_annee), ("temps_mois", t_mois), ("temps_jour_semaine", t_jour),
                    ("temps_heure", t_heure), ("temps_tranche_horaire", t_tranche), ("temps_saison", t_saison),
                    ("temps_weekend", t_we), ("temps_jour_ferie", t_ferie),
                    ("temps_mois_x_annee", t_mois_annee), ("temps_jour_x_heure", heat)):
        rap.ajouter(nom, df)

    # ---- Dimension géographique ----
    g_dep = synthese(a, "dep", tri="freq")
    g_reg = synthese(a, "region", tri="freq")
    g_metro = synthese(a, "metropole")
    g_catr = synthese(a, "catr_libelle", tri="freq")
    g_agg = synthese(a, "agg_libelle", tri="freq")
    g_com = synthese(a, "com", tri="freq").head(30)
    for nom, df in (("geo_departement", g_dep), ("geo_region", g_reg), ("geo_metropole", g_metro),
                    ("geo_type_route", g_catr), ("geo_agglomeration", g_agg), ("geo_top30_communes", g_com)):
        rap.ajouter(nom, df)

    # ---- Facteurs de gravité ----
    f = {"lum_libelle": "Luminosité", "meteo_groupe": "Météo", "surf_libelle": "État de la chaussée",
         "catr_libelle": "Type de route", "agg_libelle": "Agglomération", "int_libelle": "Intersection",
         "col_libelle": "Type de collision", "vma_classe": "Vitesse maximale autorisée (km/h)",
         "implique_pieton": "Piéton impliqué", "implique_2rm": "2RM impliqué", "implique_velo": "Vélo impliqué",
         "implique_pl": "Poids lourd impliqué", "implique_edp": "EDP impliquée"}
    gf = {col: synthese(a, col) for col in f}
    for col, df in gf.items():
        rap.ajouter(f"grav_{col}", df)

    # ---- Profil des usagers ----
    p_mode = synthese_usagers(u, "mode_transport").sort_values("nb_usagers", ascending=False)
    p_age = synthese_usagers(u, "tranche_age")
    p_sexe = synthese_usagers(u, "sexe_libelle")
    p_cat = synthese_usagers(u, "categorie_usager")
    p_age_sexe = synthese_usagers(u, ["tranche_age", "sexe_libelle"])
    for nom, df in (("usagers_mode", p_mode), ("usagers_age", p_age), ("usagers_sexe", p_sexe),
                    ("usagers_categorie", p_cat), ("usagers_age_x_sexe", p_age_sexe)):
        rap.ajouter(nom, df)
    log.info("   Gravité par mode de transport :\n%s", p_mode[["nb_usagers", "pct_tues", "pct_graves"]].to_string())

    # ---- Figures ----
    plt.rcParams.update({"axes.spines.top": False, "axes.spines.right": False, "font.size": 9})
    fig, ax = plt.subplots(figsize=(9, 4.5))
    for y in t_mois_annee.columns:
        ax.plot(range(1, 13), t_mois_annee[y].values, marker="o", label=str(y))
    ax.set_xticks(range(1, 13)); ax.set_xticklabels([m[:3] for m in MOIS])
    ax.set_title("Accidents corporels par mois", weight="bold"); ax.set_ylabel("Nombre d'accidents"); ax.legend()
    plt.tight_layout(); fig.savefig(dossier_fig / "01_accidents_par_mois.png", dpi=130); plt.close(fig)

    fig, ax = plt.subplots(figsize=(10, 4.5))
    ax.bar(t_heure.index.astype(int), t_heure["nb_accidents"], color="#2c3e50")
    ax.set_ylabel("Nombre d'accidents"); ax.set_xlabel("Heure")
    ax2 = ax.twinx(); ax2.plot(t_heure.index.astype(int), t_heure["pct_accidents_graves"], color="#c0392b", marker="o")
    ax2.set_ylabel("% accidents graves", color="#c0392b"); ax2.spines["right"].set_visible(True)
    ax.set_title("Fréquence (barres) et gravité (courbe) par heure", weight="bold")
    plt.tight_layout(); fig.savefig(dossier_fig / "02_accidents_par_heure.png", dpi=130); plt.close(fig)

    fig, ax = plt.subplots(figsize=(11, 4))
    im = ax.imshow(heat.values, aspect="auto", cmap="YlOrRd")
    ax.set_yticks(range(len(heat.index))); ax.set_yticklabels(heat.index)
    ax.set_xticks(range(0, 24, 2)); ax.set_xticklabels(range(0, 24, 2))
    ax.set_xlabel("Heure"); ax.set_title("Accidents : jour de la semaine x heure", weight="bold")
    fig.colorbar(im, ax=ax, label="Nombre d'accidents"); plt.tight_layout()
    fig.savefig(dossier_fig / "03_heatmap_jour_heure.png", dpi=130); plt.close(fig)

    for i, (col, titre) in enumerate([("lum_libelle", "Luminosité"), ("meteo_groupe", "Météo"),
                                      ("surf_libelle", "État de la chaussée"), ("catr_libelle", "Type de route"),
                                      ("vma_classe", "Vitesse maximale autorisée"), ("int_libelle", "Intersection")], 4):
        fig_barh(gf[col], col, f"Gravité des accidents selon : {titre}", f"{i:02d}_gravite_{col}.png", dossier_fig)

    d = p_mode[p_mode["nb_usagers"] >= 500].sort_values("pct_graves")
    fig, ax = plt.subplots(figsize=(9, 5))
    ax.barh(d.index.astype(str), d["pct_graves"], color="#8e44ad", label="% graves (tués + hospitalisés)")
    ax.barh(d.index.astype(str), d["pct_tues"], color="#c0392b", label="% tués")
    ax.set_title("Gravité pour l'usager selon le mode de transport", weight="bold"); ax.legend(loc="lower right")
    plt.tight_layout(); fig.savefig(dossier_fig / "10_gravite_mode_transport.png", dpi=130); plt.close(fig)

    fig, ax = plt.subplots(figsize=(9, 4.5))
    pos = list(range(len(p_age)))
    ax.bar(pos, p_age["nb_usagers"], color="#2c3e50")
    ax.set_xticks(pos); ax.set_xticklabels(p_age.index.astype(str))
    ax.set_ylabel("Nombre d'usagers")
    ax2 = ax.twinx(); ax2.plot(pos, p_age["pct_tues"], color="#c0392b", marker="o")
    ax2.set_ylabel("% tués", color="#c0392b"); ax2.spines["right"].set_visible(True)
    ax.set_title("Usagers impliqués (barres) et % de tués (courbe) par tranche d'âge", weight="bold")
    plt.tight_layout(); fig.savefig(dossier_fig / "11_usagers_par_age.png", dpi=130); plt.close(fig)

    top = g_dep.head(15).sort_values("nb_accidents")
    fig, ax = plt.subplots(figsize=(8, 5.5))
    ax.barh(top.index.astype(str), top["nb_accidents"], color="#16a085")
    ax.set_title("15 départements avec le plus d'accidents (nombre brut, non rapporté à la population)", weight="bold", fontsize=10)
    plt.tight_layout(); fig.savefig(dossier_fig / "12_top_departements.png", dpi=130); plt.close(fig)
    log.info("   %d figures enregistrées", len(list(dossier_fig.glob('*.png'))))


# =============================================================================
# MAIN
# =============================================================================
def main():
    ap = argparse.ArgumentParser(description="Pipeline BAAC 2023-2025")
    ap.add_argument("--entree", required=True, help="Dossier contenant les CSV BAAC")
    ap.add_argument("--sortie", default="sortie_baac", help="Dossier de sortie")
    ap.add_argument("--annees", nargs="+", type=int, default=list(ANNEES))
    args = ap.parse_args()
    logging.basicConfig(level=logging.INFO, format="%(message)s", stream=sys.stdout)

    entree, sortie = Path(args.entree), Path(args.sortie)
    d_don, d_rap, d_fig = sortie / "donnees_propres", sortie / "rapports", sortie / "figures"
    for d in (d_don, d_rap, d_fig):
        d.mkdir(parents=True, exist_ok=True)

    brut = charger(entree, args.annees)
    r_expl, r_diag, r_ind = Rapport(), Rapport(), Rapport()
    explorer(brut, r_expl)
    diagnostiquer(brut, r_diag)
    J = Journal()
    c, lx, u, v = nettoyer(brut, J)
    a, u, v = transformer(c, lx, u, v, J)
    indicateurs(a, u, v, r_ind, d_fig)

    log.info("8. EXPORT")
    r_expl.ecrire(d_rap / "01_exploration.xlsx")
    r_diag.ecrire(d_rap / "02_diagnostic_qualite.xlsx")
    jr = Rapport(); jr.ajouter("journal_nettoyage", J.df()); jr.ecrire(d_rap / "03_journal_nettoyage.xlsx")
    r_ind.ecrire(d_rap / "04_indicateurs.xlsx")
    for nom, df in (("accidents", a), ("usagers", u), ("vehicules", v)):
        df.to_csv(d_don / f"{nom}.csv", sep=";", decimal=",", index=False, encoding="utf-8-sig")
        log.info("   -> %s (%d lignes, %d colonnes)", d_don / f"{nom}.csv", len(df), df.shape[1])
    log.info("Terminé.")


if __name__ == "__main__":
    main()