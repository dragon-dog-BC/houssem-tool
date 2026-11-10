"""
app.py - Version web (navigateur) de l'outil de remplissage des rapports GTG.
Reutilise directement la logique testee de fill_gtg_report.py.

Chaque utilisateur entre SES cles API gratuites (Gemini : autant de cles que
voulu, une par ligne ; Mistral en option comme secours). Quand une cle est
epuisee, l'outil passe tout seul a la suivante. Les photos qui donnent 0 valeur
sont memorisees puis relues automatiquement a la fin.
Rien n'est stocke cote serveur entre les sessions.
"""

import io
import json
import os
import re
import tempfile
import time
import zipfile

import openpyxl
import streamlit as st

from fill_gtg_report import (
    HEURES,
    QuotaExhaustedError,
    build_extraction_prompt,
    build_sensor_map,
    extract_from_image_gemini,
    extract_from_image_mistral,
    norm_tag,
)

try:
    from google import genai
except ImportError:
    genai = None


st.set_page_config(page_title="Remplissage rapport GTG", page_icon="🔧", layout="centered")
st.title("🔧 Remplissage automatique des rapports GTG")
st.caption(
    "Uploadez votre rapport Excel et vos photos d'écran de tournée — "
    "l'outil lit les valeurs des capteurs et remplit le rapport automatiquement."
)

with st.expander("ℹ️ Comment ça marche / confidentialité"):
    st.markdown(
        "- Collez **toutes vos clés Gemini** (une par ligne) : quand une clé est épuisée, l'outil passe tout seul à la suivante.\n"
        "- Les photos qui donnent **0 valeur** sont mémorisées et **relues automatiquement** à la fin.\n"
        "- Rien n'est conservé sur le serveur après votre session : vos fichiers et clés "
        "restent dans votre navigateur/session uniquement.\n"
        "- Clé Gemini gratuite : https://aistudio.google.com/apikey\n"
        "- Clé Mistral gratuite (optionnelle, en secours si Gemini est saturé) : "
        "https://console.mistral.ai/api-keys"
    )

# ---------------------------------------------------------------------------
# Etat de session
# ---------------------------------------------------------------------------
defaults = {
    "sensor_map": None, "excel_by_turbine": {}, "default_template_bytes": None,
    "excel_names_key": None, "excel_display_names": [],
    "matched": [], "unmatched": [], "seen_tags": set(), "processed_idx": 0,
    "photo_names": [], "run_state": None, "final_bytes": None,
    "failed_idx": [],
}
for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v


# ---------------------------------------------------------------------------
# Gestion de plusieurs cles API (rotation automatique)
# ---------------------------------------------------------------------------
GEMINI_MODEL = "gemini-3.5-flash"
PACE_SECONDS = 5        # delai mini entre 2 requetes sur la MEME cle (limite ~10 req/min)
COOLDOWN_SECONDS = 65   # pause d'une cle qui a touche la limite "par minute"


class AllKeysExhausted(Exception):
    """Plus aucune cle (ni secours) disponible."""


class KeyPool:
    def __init__(self):
        self.keys, self.dead, self.cool, self.streak = [], {}, {}, {}

    def set_keys(self, keys):
        self.keys = list(keys)
        self.dead = {k: v for k, v in self.dead.items() if k in self.keys}
        self.cool = {k: v for k, v in self.cool.items() if k in self.keys}
        self.streak = {k: v for k, v in self.streak.items() if k in self.keys}

    def reset(self):
        self.dead, self.cool, self.streak = {}, {}, {}

    def number(self, k):
        return self.keys.index(k) + 1

    def label(self, k):
        return f"n°{self.number(k)} (…{k[-4:]})"

    def alive(self):
        return [k for k in self.keys if k not in self.dead]

    def acquire(self):
        """(cle, 0) si une cle est utilisable maintenant ; (None, secondes) si toutes
        sont en pause ; (None, None) si toutes sont definitivement epuisees."""
        alive = self.alive()
        if not alive:
            return None, None
        now = time.time()
        for k in alive:
            if self.cool.get(k, 0) <= now:
                return k, 0
        return None, max(1, int(min(self.cool[k] for k in alive) - now) + 1)

    def success(self, k):
        self.streak[k] = 0
        self.cool[k] = time.time() + PACE_SECONDS

    def failure(self, k, msg):
        """Enregistre un echec de quota. Retourne la raison si la cle est
        definitivement hors service, sinon None (simple pause)."""
        low = msg.lower()
        self.streak[k] = self.streak.get(k, 0) + 1
        if "cle_invalide" in low:
            self.dead[k] = "clé invalide ou expirée"
        elif any(x in low for x in ("perday", "per day", "daily", "per_day")):
            self.dead[k] = "quota journalier épuisé"
        elif self.streak[k] >= 3:
            self.dead[k] = "quota épuisé"
        else:
            self.cool[k] = time.time() + COOLDOWN_SECONDS
        return self.dead.get(k)


if "key_pool" not in st.session_state:
    st.session_state.key_pool = KeyPool()


# ---------------------------------------------------------------------------
# 1) Fichier(s) Excel
# ---------------------------------------------------------------------------
st.header("1. Rapport(s) Excel")
st.caption(
    "➡️ Première tournée (fichiers vierges) : uploadez UN SEUL modèle Excel, il sera "
    "dupliqué automatiquement pour chaque turbine détectée.\n\n"
    "➡️ Tournée suivante sur les MÊMES fichiers (ex: vous avez déjà rempli le 4h et voulez "
    "ajouter le 16h) : uploadez les fichiers déjà remplis renvoyés par l'outil la dernière "
    "fois (ex: rapport_GTG812_4h.xlsx, rapport_GTG813_4h.xlsx...). L'outil les reconnaît "
    "par leur nom et ajoute les nouvelles valeurs SANS toucher à ce qui est déjà rempli."
)
excel_files = st.file_uploader("Fichier(s) Excel (.xlsx)", type=["xlsx"], accept_multiple_files=True)

TURBINE_IN_NAME_RE = re.compile(r'GTG[\s_-]?(\d{2,4})', re.IGNORECASE)

if excel_files:
    names_key = tuple(sorted(f.name for f in excel_files))
    if names_key != st.session_state.get("excel_names_key"):
        with st.spinner("Analyse de la structure du/des classeur(s)..."):
            excel_by_turbine = {}   # turbine_id -> bytes (fichiers deja nommes par turbine)
            default_template_bytes = None  # fichier "generique" sans turbine dans le nom
            sensor_map = None

            for f in excel_files:
                data = f.read()
                with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
                    tmp.write(data)
                    tmp_path = tmp.name
                try:
                    if sensor_map is None:  # la structure est identique, on ne l'analyse qu'une fois
                        sensor_map = build_sensor_map(tmp_path)
                finally:
                    os.unlink(tmp_path)

                m = TURBINE_IN_NAME_RE.search(f.name)
                if m:
                    excel_by_turbine[m.group(1)] = data
                elif default_template_bytes is None:
                    default_template_bytes = data

            st.session_state.sensor_map = sensor_map
            st.session_state.excel_by_turbine = excel_by_turbine
            st.session_state.default_template_bytes = default_template_bytes or next(iter(excel_by_turbine.values()))
            st.session_state.excel_names_key = names_key
            st.session_state.excel_display_names = [f.name for f in excel_files]
            # reset une eventuelle tournee precedente
            st.session_state.matched, st.session_state.unmatched = [], []
            st.session_state.seen_tags, st.session_state.processed_idx = set(), 0
            st.session_state.run_state, st.session_state.final_bytes = None, None

if st.session_state.get("sensor_map"):
    n_named = len(st.session_state.get("excel_by_turbine", {}))
    st.success(
        f"{len(st.session_state.excel_display_names)} fichier(s) chargé(s) — "
        f"{len(st.session_state.sensor_map)} capteur(s)/champ(s) détecté(s)."
        + (f" {n_named} fichier(s) reconnu(s) comme turbine(s) existante(s) "
           f"({', '.join('GTG' + t for t in st.session_state.excel_by_turbine)})." if n_named else "")
    )

st.divider()

# ---------------------------------------------------------------------------
# 2) Parametres

# ---------------------------------------------------------------------------
HEURE_GROUPS = {
    "4h": ["4h"], "8h": ["8h"], "12h": ["12h"], "16h": ["16h"], "20h": ["20h"], "24h": ["24h"],
    "4h + 8h (même valeur dans les deux)": ["4h", "8h"],
    "16h + 20h (même valeur dans les deux)": ["16h", "20h"],
}

st.header("2. Paramètres de la tournée")
c1, c2 = st.columns(2)
with c1:
    heure_choice = st.selectbox("Heure du relevé", list(HEURE_GROUPS.keys()))
    target_heures = HEURE_GROUPS[heure_choice]
with c2:
    auto_turbine = st.checkbox(
        "🔎 Détecter automatiquement la turbine sur chaque photo",
        value=True,
        help="Coché : vous pouvez mélanger les photos de plusieurs turbines (GTG 811, 812, 813...) "
             "en un seul lot, l'outil lit le numéro affiché sur chaque écran et trie tout seul. "
             "Décoché : toutes les photos uploadées sont considérées comme appartenant à UNE seule "
             "turbine, que vous précisez ci-dessous.",
    )
turbine = None
if not auto_turbine:
    turbine = st.text_input("Numéro de turbine (ex: 812)", "")

st.subheader("Clés API (Gemini — gratuites)")
keys_text = st.text_area(
    "Vos clés Gemini — une par ligne",
    height=130,
    placeholder="AIzaSy...clé 1\nAIzaSy...clé 2\nAIzaSy...clé 3",
    help="Collez toutes vos clés d'un coup. Quand une clé est épuisée, l'outil passe "
         "automatiquement à la suivante. Astuce : créez chaque clé dans un PROJET Google "
         "différent (ou un compte différent), sinon les clés d'un même projet partagent le même quota. "
         "Obtenez des clés gratuites sur https://aistudio.google.com/apikey",
)
gemini_keys = list(dict.fromkeys(k for k in re.split(r"[\s,;]+", keys_text or "") if k))
pool = st.session_state.key_pool
pool.set_keys(gemini_keys)
if gemini_keys:
    st.caption(f"🔑 {len(gemini_keys)} clé(s) détectée(s), {len(pool.alive())} utilisable(s).")
    if pool.dead:
        with st.expander("Clés hors service"):
            for k, why in pool.dead.items():
                st.write(f"Clé {pool.label(k)} : {why}")

c3, c4 = st.columns(2)
with c3:
    max_reread = st.number_input(
        "Relectures auto des photos à 0 valeur", min_value=0, max_value=5, value=2, step=1,
        help="À la fin de la tournée, les photos qui ont donné 0 valeur sont relues "
             "automatiquement, jusqu'à ce nombre de passages.")

use_fallback = st.checkbox("Activer un secours automatique avec Mistral si toutes les clés Gemini sont épuisées")
mistral_key = None
if use_fallback:
    mistral_key = st.text_input("Votre clé Mistral", type="password",
                                 help="Obtenez une clé gratuite sur https://console.mistral.ai/api-keys")


st.divider()

# ---------------------------------------------------------------------------
# 3) Photos
# ---------------------------------------------------------------------------
st.header("3. Photos des écrans")
photos = st.file_uploader(
    "Toutes les photos de cette tournée (peu importe l'ordre)",
    type=["jpg", "jpeg", "png", "webp"], accept_multiple_files=True,
)
if photos:
    st.write(f"📷 {len(photos)} photo(s) sélectionnée(s).")

st.divider()

# ---------------------------------------------------------------------------
# 4) Lancer
# ---------------------------------------------------------------------------
st.header("4. Lancer le remplissage")


if genai is None:
    st.error("Le package google-genai n'est pas installé.")

can_run = bool(st.session_state.sensor_map and photos and gemini_keys and target_heures and genai is not None)
if not can_run:
    st.info("Complétez les étapes 1 à 3 ci-dessus pour activer le bouton.")

run_clicked = st.button("▶️ Remplir le rapport", disabled=not can_run, type="primary")
resume_clicked = False
if st.session_state.run_state == "stopped":
    st.warning(
        f"⏸️ Arrêté après {st.session_state.processed_idx}/{len(st.session_state.photo_names)} photo(s) "
        "— toutes les clés sont épuisées. Ajoutez de nouvelles clés ci-dessus (ou attendez le renouvellement "
        "du quota), puis :"
    )
    resume_clicked = st.button("🔁 Reprendre là où ça s'est arrêté")

log_box = st.container(height=250) if hasattr(st, "container") else st.container()
progress_bar = st.progress(0.0)




def run_batch(resume=False, retry=False):
    sensor_map = st.session_state.sensor_map
    prompt = build_extraction_prompt(turbine or None, auto_detect_turbine=auto_turbine)
    pool = st.session_state.key_pool
    pool.set_keys(gemini_keys)
    clients = {}
    backend = {"name": "gemini"}

    photo_list = photos if not (resume or retry) else st.session_state.get("_photo_cache")
    st.session_state["_photo_cache"] = photo_list
    st.session_state.photo_names = [p.name for p in photo_list]

    if not (resume or retry):
        pool.reset()
        st.session_state.matched, st.session_state.unmatched = [], []
        st.session_state.seen_tags, st.session_state.processed_idx = set(), 0
        st.session_state.failed_idx = []

    start_idx = st.session_state.processed_idx
    total = len(photo_list)

    # ---- lecture d'UNE image avec rotation automatique des cles ----
    def extract(tmp_path):
        while True:
            if backend["name"] == "gemini":
                key, wait = pool.acquire()
                if key is None and wait is None:      # toutes les cles Gemini sont mortes
                    if use_fallback and mistral_key:
                        log_box.warning("⚠️ Toutes les clés Gemini sont épuisées. Bascule sur Mistral...")
                        backend["name"] = "mistral"
                        continue
                    raise AllKeysExhausted()
                if key is None:                        # toutes en pause "par minute"
                    log_box.write(f"⏳ Toutes les clés sont en pause (limite par minute), attente {wait}s...")
                    time.sleep(wait)
                    continue
                if key not in clients:
                    clients[key] = genai.Client(api_key=key)
                try:
                    res = extract_from_image_gemini(
                        clients[key], tmp_path, GEMINI_MODEL, prompt, fast_fail_quota=True)
                except QuotaExhaustedError as qe:
                    dead_reason = pool.failure(key, str(qe))
                    if dead_reason:
                        log_box.warning(f"🔑 Clé {pool.label(key)} hors service ({dead_reason}). "
                                         f"Passage à la clé suivante... ({len(pool.alive())} restante(s))")
                    else:
                        log_box.write(f"⏸️ Clé {pool.label(key)} : limite atteinte, mise en pause {COOLDOWN_SECONDS}s. "
                                       "Essai avec la clé suivante...")
                    continue
                pool.success(key)
                return res, f"Gemini clé {pool.label(key)}"
            else:
                try:
                    return extract_from_image_mistral(
                        tmp_path, "pixtral-large-latest", prompt, mistral_key), "Mistral"
                except QuotaExhaustedError:
                    raise AllKeysExhausted()

    # ---- traitement complet d'UNE photo ----
    def process_photo(i):
        photo = photo_list[i]
        photo.seek(0)
        suffix = os.path.splitext(photo.name)[1] or ".jpg"
        with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as tmp:
            tmp.write(photo.read())
            tmp_path = tmp.name
        try:
            result, used = extract(tmp_path)
        finally:
            os.unlink(tmp_path)

        readings = result.get("readings", []) or []
        valid = [r for r in readings if norm_tag(r.get("tag")) is not None and r.get("value") is not None]

        # photo a "0 valeur" (echec, reponse vide, ou rien d'exploitable) -> memorisee pour relecture
        is_zero = bool(result.get("_failed")) or len(valid) == 0
        if is_zero:
            if i not in st.session_state.failed_idx:
                st.session_state.failed_idx.append(i)
        elif i in st.session_state.failed_idx:
            st.session_state.failed_idx.remove(i)

        title = result.get("screen_title") or "(titre non lu)"
        screen_turbine = result.get("turbine") if auto_turbine else turbine
        log_box.write(
            f"**[{i + 1}/{total}]** {photo.name} ({used}) → {len(valid)} valeur(s) lue(s) ({title}"
            + (f", turbine {screen_turbine}" if screen_turbine else "") + ")"
            + (" — ⚠️ 0 valeur : mémorisée pour relecture" if is_zero else ""))

        for reading in readings:
            tag = norm_tag(reading.get("tag"))
            value = reading.get("value")
            reading_turbine = (reading.get("turbine") if auto_turbine else None) or screen_turbine
            if tag is None or value is None:
                continue
            if auto_turbine and not reading_turbine:
                st.session_state.unmatched.append((tag, value, title, "turbine non identifiée", None))
                continue
            st.session_state.seen_tags.add(tag)
            if tag in sensor_map:
                sheet_name = sensor_map[tag]['sheet']
                for h in target_heures:
                    coord = sensor_map[tag]['cells'].get(h)
                    if coord is None:
                        st.session_state.unmatched.append((tag, value, title, f"heure '{h}' introuvable", reading_turbine))
                        continue
                    coord_display = "+".join(coord) if isinstance(coord, list) else coord
                    st.session_state.matched.append((tag, value, sheet_name, coord_display, title, reading_turbine))
            else:
                st.session_state.unmatched.append((tag, value, title, "tag absent du classeur", reading_turbine))

    # ---- 1) passage principal ----
    if not retry:
        for i in range(start_idx, total):
            try:
                process_photo(i)
            except AllKeysExhausted:
                st.session_state.processed_idx = i
                st.session_state.run_state = "stopped"
                st.error(
                    f"⚠️ Toutes les clés sont épuisées après {i}/{total} photo(s). "
                    "Ajoutez de nouvelles clés ci-dessus puis cliquez sur *Reprendre*. "
                    "Les photos à 0 valeur seront relues à la fin."
                )
                return
            st.session_state.processed_idx = i + 1
            progress_bar.progress((i + 1) / total)

    # ---- 2) relecture automatique des photos a 0 valeur ----
    exhausted = False
    passes = int(max_reread) if not retry else max(int(max_reread), 1)
    for p in range(1, passes + 1):
        todo = sorted(st.session_state.failed_idx)
        if not todo:
            break
        log_box.write(f"🔁 **Relecture {p}/{passes}** : {len(todo)} photo(s) à 0 valeur...")
        for pos, i in enumerate(todo, 1):
            try:
                process_photo(i)
            except AllKeysExhausted:
                exhausted = True
                break
            progress_bar.progress(pos / len(todo))
        if exhausted:
            break
    if exhausted:
        st.warning("⚠️ Clés épuisées pendant la relecture. Ajoutez des clés puis utilisez "
                   "le bouton de relecture plus bas.")

    # --- construit un classeur separe par turbine detectee, a partir du meme modele ---

    by_turbine = {}
    for tag, value, sheet_name, coord_display, title, tb in st.session_state.matched:
        key = tb or "turbine_inconnue"
        by_turbine.setdefault(key, []).append((sheet_name, coord_display, value))

    zip_buf = io.BytesIO()
    heure_label = "_".join(target_heures)  # ex "4h_8h" pour un groupe, "12h" pour un seul
    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
        base_name = "rapport"
        excel_by_turbine = st.session_state.excel_by_turbine
        for tb, entries in by_turbine.items():
            # reprend le fichier deja rempli de CETTE turbine s'il a ete fourni
            # (garde tout ce qui etait deja rempli, ex les autres heures),
            # sinon repart du modele vierge generique.
            base_bytes = excel_by_turbine.get(tb, st.session_state.default_template_bytes)
            wb = openpyxl.load_workbook(io.BytesIO(base_bytes))
            for sheet_name, coord_display, value in entries:
                ws = wb[sheet_name]
                for c in coord_display.split('+'):
                    ws[c] = value
            buf = io.BytesIO()
            wb.save(buf)
            fname = f"{base_name}_GTG{tb}_{heure_label}.xlsx" if tb != "turbine_inconnue" else f"{base_name}_turbine_INCONNUE_{heure_label}.xlsx"
            zf.writestr(fname, buf.getvalue())

    st.session_state.final_bytes = zip_buf.getvalue()
    st.session_state.turbines_found = sorted(by_turbine.keys())
    st.session_state.run_state = "done"


if run_clicked:
    st.session_state["_photo_cache"] = photos
    run_batch(resume=False)

if resume_clicked:
    run_batch(resume=True)


# ---------------------------------------------------------------------------
# 5) Resultats
# ---------------------------------------------------------------------------
if st.session_state.run_state == "done":
    st.success("✅ Tournée terminée !")
    failed_idx = sorted(st.session_state.failed_idx)
    if failed_idx:
        names = [st.session_state.photo_names[i] for i in failed_idx if i < len(st.session_state.photo_names)]
        st.warning(
            f"⚠️ {len(names)} photo(s) sont restées à 0 valeur, même après relecture : "
            f"**{', '.join(names)}**"
        )
        if st.button(f"🔁 Relire uniquement ces {len(names)} photo(s)", type="primary"):
            run_batch(retry=True)
            st.rerun()
    matched = st.session_state.matched
    unmatched = st.session_state.unmatched
    missing = sorted(set(st.session_state.sensor_map.keys()) - st.session_state.seen_tags)
    turbines_found = st.session_state.get("turbines_found", [])

    if auto_turbine:
        st.subheader(f"🏭 {len(turbines_found)} turbine(s) détectée(s) : {', '.join(turbines_found)}")

    st.subheader(f"✅ {len(matched)} valeur(s) écrite(s)")
    with st.expander("Voir le détail"):
        for tag, value, sheet, coord, title, tb in matched:
            tb_label = f"[GTG {tb}] " if tb else ""
            st.write(f"{tb_label}`{tag}` = {value} → {sheet}!{coord}  _(depuis: {title})_")

    if unmatched:
        st.subheader(f"⚠️ {len(unmatched)} valeur(s) à vérifier")
        with st.expander("Voir le détail"):
            for tag, value, title, reason, tb in unmatched:
                tb_label = f"[GTG {tb}] " if tb else ""
                st.write(f"{tb_label}`{tag}` = {value}  ({reason}, depuis: {title})")

    if missing:
        st.subheader(f"❓ {len(missing)} capteur(s)/champ(s) jamais vus (toutes turbines confondues)")
        with st.expander("Voir la liste"):
            st.write(", ".join(missing))

    zip_name = f"rapport_rempli_{'_'.join(target_heures)}.zip"
    st.download_button(
        f"⬇️ Télécharger les {len(turbines_found)} rapport(s) remplis (.zip)",
        data=st.session_state.final_bytes,
        file_name=zip_name,
        mime="application/zip",
        type="primary",
    )
