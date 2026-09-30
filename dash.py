import io
import re
import zipfile
from pathlib import Path

import cv2
import numpy as np
import pytesseract
import streamlit as st
from PIL import Image

# --- Si Tesseract n'est pas dans le PATH (Windows local uniquement) :
# pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

st.set_page_config(page_title="Renommage AG", page_icon="🖼️", layout="centered")

# Regex : AG + chiffres + tiret + suite (lettres/chiffres/espaces/tirets)
PATTERN_AG = re.compile(r"AG\s*\d+\s*-\s*[\w\s\-]+", re.IGNORECASE)


# --------------------------------------------------------------------
# PRÉTRAITEMENTS
# --------------------------------------------------------------------

def preprocess_variants(img_pil: Image.Image) -> list[Image.Image]:
    """Retourne plusieurs versions prétraitées de l'image pour maximiser les chances OCR."""
    variants = []

    # 1. Original agrandi x3 (Tesseract aime les grands caractères)
    big = img_pil.resize((img_pil.width * 3, img_pil.height * 3), Image.LANCZOS)
    variants.append(big.convert("L"))

    # 2. Niveaux de gris + seuillage Otsu (noir/blanc pur)
    gray = np.array(img_pil.convert("L"))
    _, otsu = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    variants.append(Image.fromarray(otsu))

    # 3. Version inversée (utile pour texte blanc sur fond sombre/rouge)
    variants.append(Image.fromarray(cv2.bitwise_not(otsu)))

    # 4. Version agrandie x3 + inversée
    big_arr = np.array(big.convert("L"))
    _, otsu_big = cv2.threshold(big_arr, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    variants.append(Image.fromarray(cv2.bitwise_not(otsu_big)))

    return variants


def crop_regions(img_pil: Image.Image) -> list[Image.Image]:
    """Retourne plusieurs zones de l'image (bandeau haut, moitié haute, image entière)."""
    w, h = img_pil.size
    return [
        img_pil.crop((0, 0, w, int(h * 0.10))),   # bandeau très haut
        img_pil.crop((0, 0, w, int(h * 0.18))),   # bandeau haut
        img_pil.crop((0, 0, w, int(h * 0.30))),   # tiers supérieur
        img_pil.crop((0, 0, w, int(h * 0.50))),   # moitié haute
        img_pil,                                   # image complète
    ]


# --------------------------------------------------------------------
# EXTRACTION
# --------------------------------------------------------------------

def try_ocr_on(image_pil: Image.Image, lang: str = "fra+eng") -> str:
    """OCR avec plusieurs modes PSM, retourne le texte concaténé."""
    text = ""
    for psm in (6, 7, 11):  # 6=bloc, 7=ligne unique, 11=texte épars
        try:
            cfg = f"--psm {psm}"
            text += "\n" + pytesseract.image_to_string(image_pil, lang=lang, config=cfg)
        except Exception:
            pass
    return text


def extract_code_ag(img_pil: Image.Image, debug: bool = False) -> tuple[str | None, str]:
    """
    Cherche le code AG en combinant :
    - plusieurs zones (bandeau, moitié haute, image entière)
    - plusieurs prétraitements (gris, Otsu, inversé, agrandi)
    - plusieurs modes Tesseract
    Retourne (code, texte_ocr_brut).
    """
    all_text = ""

    for region in crop_regions(img_pil):
        for variant in preprocess_variants(region):
            text = try_ocr_on(variant)
            all_text += "\n" + text

            match = PATTERN_AG.search(text)
            if match:
                code = clean_code(match.group(0))
                return code, all_text

    return None, all_text


def clean_code(code: str) -> str:
    """Normalise le code extrait."""
    code = code.strip()
    code = re.sub(r"\s*-\s*", "-", code)   # "A - B" -> "A-B"
    code = re.sub(r"\s+", " ", code)       # espaces multiples -> un seul
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
            if r["new_name"]:
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

# --- Affichage
if st.session_state.get("results"):
    results = st.session_state["results"]
    ok = [r for r in results if r["new_name"]]

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
            if r["new_name"]:
                st.markdown(f"**Nouveau :** `{r['new_name']}`")
                st.caption(f"Code détecté : `{r['code']}`")
            else:
                st.error("Aucun code AG détecté.")

            if show_debug:
                with st.expander("Voir le texte OCR brut"):
                    st.text(r["raw_text"][:5000] if r["raw_text"] else "(vide)")

    if st.button("🔄 Réinitialiser"):
        st.session_state.pop("results", None)
        st.rerun()
