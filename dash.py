import io
import re
import zipfile
from pathlib import Path

import numpy as np
import pytesseract
import streamlit as st
from PIL import Image

# --------------------------------------------------------------------
# Si Tesseract n'est pas dans le PATH (Windows), décommentez :
# pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
# --------------------------------------------------------------------

st.set_page_config(page_title="Renommage AG", page_icon="🖼️", layout="centered")

# Motif : AG + chiffres + tiret + suite de lettres/chiffres/espaces/tirets
PATTERN_AG = re.compile(r"AG\s*\d+\s*-\s*[\w\s\-]+", re.IGNORECASE)


# --------------------------------------------------------------------
# FONCTIONS
# --------------------------------------------------------------------

def ocr_image(img_pil: Image.Image, lang: str = "fra+eng") -> str:
    """OCR sur toute l'image."""
    try:
        return pytesseract.image_to_string(img_pil.convert("L"), lang=lang)
    except Exception as e:
        return ""


def extract_code_ag(img_pil: Image.Image) -> str | None:
    """
    Cherche un code AG en plusieurs passes :
    1) OCR sur le haut de l'image (bandeau) — cas le plus fréquent
    2) OCR sur l'image entière — fallback
    """
    w, h = img_pil.size

    # --- Passe 1 : rogner le haut (bandeau ~ 18% de la hauteur)
    top_crop = img_pil.crop((0, 0, w, int(h * 0.18)))
    # Agrandir pour améliorer l'OCR sur les petits caractères
    top_crop = top_crop.resize((top_crop.width * 2, top_crop.height * 2))

    text = ocr_image(top_crop)

    # --- Passe 2 : si rien trouvé, OCR complet
    if not PATTERN_AG.search(text):
        text = ocr_image(img_pil)

    # --- Extraction avec la regex
    match = PATTERN_AG.search(text)
    if not match:
        return None

    code = match.group(0)

    # Nettoyage : espaces multiples, tirets en fin, espaces autour des tirets
    code = code.strip()
    code = re.sub(r"\s*-\s*", "-", code)      # "A - B" -> "A-B"
    code = re.sub(r"\s+", " ", code)          # espaces multiples -> un seul
    code = code.strip("- ").strip()           # enlève tirets/espaces en fin
    return code


def sanitize(text: str, max_len: int = 120) -> str:
    """Nettoie le code pour en faire un nom de fichier valide."""
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
                img_bytes = io.BytesIO()
                r["image"].save(img_bytes, format=ext)
                zf.writestr(r["new_name"], img_bytes.getvalue())
    buf.seek(0)
    return buf.getvalue()


# --------------------------------------------------------------------
# INTERFACE
# --------------------------------------------------------------------

st.title("🖼️ Renommage automatique — Code AG")
st.write(
    "Téléversez vos photos. L'outil détecte le code **AG…** en haut de l'image "
    "et renomme les fichiers."
)

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
        code = extract_code_ag(img)

        if code:
            clean = sanitize(code)
            new_name = unique_name(clean, f.name, used) if clean else None
        else:
            clean, new_name = None, None

        results.append({
            "original": f.name,
            "code": code,
            "new_name": new_name,
            "image": img,
        })

        progress.progress((i + 1) / len(files), text=f"{i + 1}/{len(files)} — {f.name}")

    progress.empty()
    st.session_state["results"] = results

# --- Affichage des résultats
if st.session_state.get("results"):
    results = st.session_state["results"]
    ok = [r for r in results if r["new_name"]]

    st.success(f"✅ {len(ok)} / {len(results)} image(s) renommée(s) avec succès.")

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

    if st.button("🔄 Réinitialiser"):
        st.session_state.pop("results", None)
        st.rerun()
