import io
import re
import zipfile
from pathlib import Path

import cv2
import numpy as np
import pytesseract
import streamlit as st
from PIL import Image

# pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

st.set_page_config(page_title="Renommage AG", page_icon="🖼️", layout="centered")

# Version du schéma de résultats — à incrémenter si on change la structure
RESULTS_VERSION = 3

PATTERN_AG = re.compile(r"AG\s*[O0-9]{1,3}\s*-\s*[\w\s\-]{5,}", re.IGNORECASE)


# --------------------------------------------------------------------
# PRÉTRAITEMENTS
# --------------------------------------------------------------------

def variants_for_ocr(img_pil: Image.Image) -> list[Image.Image]:
    """Génère plusieurs versions prétraitées pour maximiser les chances."""
    out = []

    # 1. Agrandissement x2 puis x3 (netteté)
    for scale in (2, 3):
        big = img_pil.resize((img_pil.width * scale, img_pil.height * scale), Image.LANCZOS)
        out.append(big.convert("L"))

    # 2. Otsu (noir sur blanc) sur l'image agrandie x2
    big2 = img_pil.resize((img_pil.width * 2, img_pil.height * 2), Image.LANCZOS).convert("L")
    arr = np.array(big2)
    _, otsu = cv2.threshold(arr, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    out.append(Image.fromarray(otsu))
    # 3. Version inversée (essentiel pour texte blanc sur fond rouge)
    out.append(Image.fromarray(cv2.bitwise_not(otsu)))

    # 4. Seuillage spécifique pour faire ressortir le blanc (texte blanc sur rouge)
    #    On garde les pixels très clairs
    _, white_keep = cv2.threshold(arr, 180, 255, cv2.THRESH_BINARY)
    out.append(Image.fromarray(cv2.bitwise_not(white_keep)))

    return out


def tiles(img_pil: Image.Image) -> list[Image.Image]:
    """
    Découpe l'image en zones qui pourraient contenir le code :
    - bandeau haut global
    - chaque moitié verticale du haut
    - tiers supérieur
    - image entière
    """
    w, h = img_pil.size
    zones = [
        img_pil.crop((0, 0, w, int(h * 0.15))),
        img_pil.crop((0, 0, w, int(h * 0.25))),
        img_pil.crop((0, 0, w, int(h * 0.40))),
        img_pil.crop((0, 0, w // 2, int(h * 0.25))),          # moitié gauche du haut
        img_pil.crop((w // 2, 0, w, int(h * 0.25))),          # moitié droite du haut
        img_pil.crop((0, 0, w, int(h * 0.60))),
        img_pil,
    ]
    return zones


# --------------------------------------------------------------------
# EXTRACTION
# --------------------------------------------------------------------

def ocr_multi_psm(img_pil: Image.Image, lang: str = "fra+eng") -> str:
    """OCR avec plusieurs modes PSM, retourne tout concaténé."""
    text = ""
    for psm in (6, 7, 11, 12):
        try:
            text += "\n" + pytesseract.image_to_string(
                img_pil, lang=lang, config=f"--psm {psm}"
            )
        except Exception:
            pass
    return text


def extract_code_ag(img_pil: Image.Image) -> tuple[str | None, str]:
    """
    Parcourt toutes les zones × tous les prétraitements × tous les PSM.
    Retourne (code, texte_ocr_brut).
    """
    all_text = ""

    for zone in tiles(img_pil):
        for variant in variants_for_ocr(zone):
            text = ocr_multi_psm(variant)
            all_text += "\n" + text

            m = PATTERN_AG.search(text)
            if m:
                return clean_code(m.group(0)), all_text

    return None, all_text


def clean_code(code: str) -> str:
    code = code.strip()
    code = re.sub(r"\s*-\s*", "-", code)
    code = re.sub(r"\s+", " ", code)
    code = code.strip("- ").strip()
    return code


def sanitize(text: str, max_len: int = 120) -> str:
    text = re.sub(r"[^\w\-]", "_", text)
    text = re.sub(r"_+", "_", text).strip("_")
    return text[:max_len]


def unique_name(base: str, original: str, used: set) -> str:
    suffix = Path(original).suffix.lower() or ".png"
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
                ext = Path(r["new_name"]).suffix.lstrip(".").upper() or "PNG"
                if ext == "JPG":
                    ext = "JPEG"
                b = io.BytesIO()
                r["image"].save(b, format=ext)
                zf.writestr(r["new_name"], b.getvalue())
    buf.seek(0)
    return buf.getvalue()


# --------------------------------------------------------------------
# INTERFACE
# --------------------------------------------------------------------

st.title("🖼️ Renommage automatique — Code AG")
st.write("Téléversez vos photos. L'outil détecte le code **AG…** et renomme les fichiers.")

with st.sidebar:
    st.header("Options")
    show_debug = st.checkbox("Afficher le texte OCR brut (debug)", value=True)
    if st.button("🗑️ Vider le cache"):
        st.session_state.clear()
        st.rerun()

# --- Important : si on change la version du schéma, on invalide le cache
if st.session_state.get("results_version") != RESULTS_VERSION:
    st.session_state.pop("results", None)
    st.session_state["results_version"] = RESULTS_VERSION

files = st.file_uploader(
    "Sélectionnez vos images",
    type=["png", "jpg", "jpeg", "bmp", "tif", "tiff", "webp"],
    accept_multiple_files=True,
)

if files and st.button("🚀 Analyser et renommer", type="primary", use_container_width=True):
    results = []
    used = set()
    progress = st.progress(0.0, text="Traitement...")

    for i, f in enumerate(files):
        img = Image.open(f).convert("RGB")
        code, raw_text = extract_code_ag(img)

        new_name = None
        if code:
            clean = sanitize(code)
            if clean:
                new_name = unique_name(clean, f.name, used)

        results.append({
            "original": f.name,
            "code": code,
            "raw_text": raw_text,
            "new_name": new_name,
            "image": img,
        })

        progress.progress((i + 1) / len(files), text=f"{i + 1}/{len(files)} — {f.name}")

    progress.empty()
    st.session_state["results"] = results

# --- Affichage (avec .get() partout pour éviter le KeyError)
if st.session_state.get("results"):
    results = st.session_state["results"]
    ok = [r for r in results if r.get("new_name")]

    st.success(f"✅ {len(ok)} / {len(results)} image(s) renommée(s).")

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

    for r in results:
        col1, col2 = st.columns([1, 2])
        with col1:
            st.image(r["image"], use_container_width=True)
        with col2:
            st.markdown(f"**Original :** `{r['original']}`")
            if r.get("new_name"):
                st.markdown(f"**Nouveau :** `{r['new_name']}`")
                st.caption(f"Code détecté : `{r['code']}`")
            else:
                st.error("Aucun code AG détecté.")

            if show_debug:
                with st.expander("Voir le texte OCR brut"):
                    txt = r.get("raw_text") or "(vide)"
                    st.text(txt[:5000])

    if st.button("🔄 Réinitialiser"):
        st.session_state.pop("results", None)
        st.rerun()
