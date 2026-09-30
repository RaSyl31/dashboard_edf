import io
import os
import re
import time
import zipfile
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path

# ⚠️ AVANT d'importer pytesseract : limite les threads internes de Tesseract
os.environ["OMP_THREAD_LIMIT"] = "1"

import cv2
import numpy as np
import pytesseract
import streamlit as st
from PIL import Image

# pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

st.set_page_config(page_title="Renommage AG", page_icon="🖼️", layout="centered")

RESULTS_VERSION = 6
PATTERN_AG = re.compile(r"AG\s*[O0-9]{1,3}\s*-\s*[\w\s\-]{5,}", re.IGNORECASE)

# Nombre de threads parallèles (adapté à Streamlit Cloud : 2 vCPU)
MAX_WORKERS = 4


# --------------------------------------------------------------------
# PRÉTRAITEMENTS (réduits à l'essentiel)
# --------------------------------------------------------------------

def best_variants(img_pil: Image.Image) -> list[Image.Image]:
    """
    Génère UNIQUEMENT 2 variantes très efficaces :
    - Otsu inversé (texte blanc sur rouge → noir sur blanc)
    - Otsu normal
    Ordre : le plus prometteur en premier.
    """
    big = img_pil.resize((img_pil.width * 3, img_pil.height * 3), Image.LANCZOS).convert("L")
    arr = np.array(big)
    _, otsu = cv2.threshold(arr, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    return [
        Image.fromarray(cv2.bitwise_not(otsu)),   # inversé : priorité n°1
        Image.fromarray(otsu),
    ]


def priority_zones(img_pil: Image.Image) -> list[Image.Image]:
    """
    Zones classées par probabilité décroissante de contenir le code.
    Le bandeau AG est presque toujours dans les 25 % du haut.
    """
    w, h = img_pil.size
    return [
        img_pil.crop((0, 0, w, int(h * 0.18))),   # bandeau principal
        img_pil.crop((0, 0, w, int(h * 0.30))),   # marge
        img_pil,                                   # fallback
    ]


# --------------------------------------------------------------------
# OCR
# --------------------------------------------------------------------

def ocr_once(img_pil: Image.Image, psm: int, lang: str = "fra+eng") -> str:
    try:
        return pytesseract.image_to_string(
            img_pil, lang=lang, config=f"--psm {psm} --oem 1"
        )
    except Exception:
        return ""


def extract_code_ag(img_pil: Image.Image) -> tuple[str | None, str]:
    """
    Stratégie d'arrêt précoce :
    1) PSM 6 + PSM 7 sur la 1ère variante de la 1ère zone (2 appels)
    2) Si rien : PSM 11 + PSM 12 sur les autres variantes
    3) Si rien : élargir les zones
    Retourne (code, texte_ocr_brut).
    """
    all_text_parts: list[str] = []

    zones = priority_zones(img_pil)

    # --- Passe 1 : la zone la plus probable, 2 variantes, 2 PSM = 4 appels
    zone0 = zones[0]
    variants0 = best_variants(zone0)
    for v in variants0[:1]:           # uniquement l'inversé
        for psm in (6, 7):
            t = ocr_once(v, psm)
            all_text_parts.append(t)
            m = PATTERN_AG.search(t)
            if m:
                return clean_code(m.group(0)), "\n".join(all_text_parts)

    # --- Passe 2 : 2ème variante + PSM élargis
    for v in variants0[1:]:
        for psm in (6, 7, 11, 12):
            t = ocr_once(v, psm)
            all_text_parts.append(t)
            m = PATTERN_AG.search(t)
            if m:
                return clean_code(m.group(0)), "\n".join(all_text_parts)

    # --- Passe 3 : zones élargies
    for zone in zones[1:]:
        for v in best_variants(zone):
            for psm in (6, 11):
                t = ocr_once(v, psm)
                all_text_parts.append(t)
                m = PATTERN_AG.search(t)
                if m:
                    return clean_code(m.group(0)), "\n".join(all_text_parts)

    return None, "\n".join(all_text_parts)


def clean_code(code: str) -> str:
    code = code.strip()
    code = re.sub(r"\s*-\s*", "-", code)
    code = re.sub(r"\s+", " ", code)
    code = code.strip("- ").strip()
    return code


# --------------------------------------------------------------------
# UTILITAIRES
# --------------------------------------------------------------------

def sanitize(text: str, max_len: int = 120) -> str:
    text = re.sub(r"[^\w\-]", "_", text)
    text = re.sub(r"_+", "_", text).strip("_")
    return text[:max_len]


def unique_name(base: str, original: str, used: set) -> str:
    suffix = ".jpg"
    name = f"{base}{suffix}"
    i = 1
    while name in used:
        name = f"{base}_{i}{suffix}"
        i += 1
    used.add(name)
    return name


def make_zip(results: list[dict]) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for r in results:
            if r.get("new_name"):
                img = r["image"].convert("RGB")
                b = io.BytesIO()
                img.save(b, format="JPEG", quality=95, optimize=True)
                zf.writestr(r["new_name"], b.getvalue())
    buf.seek(0)
    return buf.getvalue()


# --------------------------------------------------------------------
# TRAITEMENT D'UNE IMAGE (appelé en parallèle)
# --------------------------------------------------------------------

def process_one(file_bytes: bytes, filename: str) -> dict:
    """Fonction pure, exécutée dans un thread. Ne touche pas à Streamlit."""
    img = Image.open(io.BytesIO(file_bytes)).convert("RGB")
    code, raw_text = extract_code_ag(img)
    new_name = None
    if code:
        clean = sanitize(code)
        if clean:
            new_name = clean  # nom sans suffixe, résolu après coup
    return {
        "original": filename,
        "code": code,
        "raw_text": raw_text,
        "new_name_base": new_name,
        "image": img,
    }


# --------------------------------------------------------------------
# INTERFACE
# --------------------------------------------------------------------

st.title("🖼️ Renommage automatique — Code AG")
st.write(
    "Téléversez vos images. L'outil détecte le code **AG…** et les renomme en `.jpg`. "
    f"Traitement parallèle ({MAX_WORKERS} threads)."
)

with st.sidebar:
    st.header("Options")
    show_debug = st.checkbox("Afficher le texte OCR brut (debug)", value=False)
    if st.button("🗑️ Vider le cache"):
        st.session_state.clear()
        st.rerun()

if st.session_state.get("results_version") != RESULTS_VERSION:
    st.session_state.pop("results", None)
    st.session_state["results_version"] = RESULTS_VERSION

files = st.file_uploader(
    "Sélectionnez vos images (10 à la fois recommandé)",
    type=["png", "jpg", "jpeg", "bmp", "tif", "tiff", "webp"],
    accept_multiple_files=True,
)

if files and st.button("🚀 Analyser et renommer", type="primary", use_container_width=True):
    # Lecture des octets en amont (obligatoire pour le threading)
    payload = [(f.getvalue(), f.name) for f in files]
    total = len(payload)

    progress = st.progress(0.0, text=f"0/{total} — démarrage...")
    status = st.empty()

    results: list[dict] = []
    t0 = time.time()
    done = 0

    # --- Traitement PARALLÈLE
    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        futures = {
            executor.submit(process_one, data, name): name
            for data, name in payload
        }
        for future in as_completed(futures):
            try:
                res = future.result()
            except Exception as e:
                res = {
                    "original": futures[future],
                    "code": None,
                    "raw_text": f"Erreur : {e}",
                    "new_name_base": None,
                    "image": None,
                }
            results.append(res)
            done += 1
            elapsed = time.time() - t0
            eta = (elapsed / done) * (total - done) if done else 0
            progress.progress(
                done / total,
                text=f"{done}/{total} — {futures[future]}  ({elapsed:.1f}s écoulées, ETA ~{eta:.1f}s)",
            )

    # --- Attribution des noms uniques (séquentiel, car dépend de l'état `used`)
    used = set()
    for r in results:
        if r.get("new_name_base"):
            r["new_name"] = unique_name(r["new_name_base"], r["original"], used)
        else:
            r["new_name"] = None

    # On remet les résultats dans l'ordre initial
    order = {f.name: i for i, f in enumerate(files)}
    results.sort(key=lambda r: order.get(r["original"], 999))

    progress.empty()
    status.empty()

    elapsed = time.time() - t0
    st.session_state["results"] = results
    st.session_state["last_duration"] = elapsed
    st.success(f"✅ Terminé en {elapsed:.1f}s pour {total} image(s).")


# --------------------------------------------------------------------
# AFFICHAGE
# --------------------------------------------------------------------
if st.session_state.get("results"):
    results = st.session_state["results"]
    ok = [r for r in results if r.get("new_name")]
    duration = st.session_state.get("last_duration", 0)

    st.info(
        f"✅ {len(ok)} / {len(results)} image(s) renommée(s) "
        f"— durée totale : **{duration:.1f}s** "
        f"({duration / max(len(results), 1):.2f}s / image)"
    )

    if ok:
        st.download_button(
            "📥 Télécharger les images renommées (ZIP)",
            data=make_zip(results),
            file_name="images_renommees.zip",
            mime="application/zip",
            type="primary",
            use_container_width=True,
        )

    st.divider()

    # Affichage compact : 3 colonnes
    cols = st.columns(3)
    for idx, r in enumerate(results):
        with cols[idx % 3]:
            if r.get("image") is not None:
                st.image(r["image"], use_container_width=True)
            st.markdown(f"**`{r['original']}`**")
            if r.get("new_name"):
                st.success(f"➡️ `{r['new_name']}`")
                st.caption(f"Code : `{r['code']}`")
            else:
                st.error("Aucun code AG détecté.")
                if show_debug:
                    with st.expander("OCR brut"):
                        st.text((r.get("raw_text") or "")[:3000])

    if st.button("🔄 Réinitialiser"):
        st.session_state.pop("results", None)
        st.session_state.pop("last_duration", None)
        st.rerun()
