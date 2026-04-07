import os
import pandas as pd
from PIL import Image, ImageDraw, ImageFont
from utils import eliminate_accents


def wrap_text(text, font, max_width, draw):
    words = text.split()
    lines = []
    line = ""
    for word in words:
        test_line = line + word + " "
        if draw.textlength(test_line, font=font) <= max_width:
            line = test_line
        else:
            lines.append(line.strip())
            line = word + " "
    lines.append(line.strip())
    return lines


def fit_text_to_box(text, font_path, start_size, max_width, max_height, draw):
    """
    Reduce el tamaño de fuente y ajusta el texto en líneas hasta que quepa en el área disponible.
    """
    size = start_size
    while size >= 10:
        font = ImageFont.truetype(font_path, size)
        lines = wrap_text(text, font, max_width, draw)
        total_height = len(lines) * (font.size + 6) - 6
        if total_height <= max_height:
            return font, lines
        size -= 1
    # fallback si no entra nunca
    font = ImageFont.truetype(font_path, 10)
    return font, wrap_text(text, font, max_width, draw)


def normalize_title_case(text):
    """
    Convierte textos completamente en mayúsculas a una capitalización natural.
    """
    import re
    ratio_upper = sum(1 for c in text if c.isupper()) / max(1, len(text))
    if ratio_upper > 0.6:
        text = text.lower()
        text = re.sub(r"(^|\.\s+)(\w)", lambda m: m.group(1) + m.group(2).upper(), text)
    return text


def create_exposition_certificate(
    certificate_template,
    expositor_name,
    title,
    authors,
    text_color,
    regular_font_path,
    bold_font_path,
    italic_font_path,
    expositor_size,
    title_size,
    authors_size,
    save_path='',
    width=None,
):
    expositor_font = ImageFont.truetype(bold_font_path, expositor_size)
    im = Image.open(certificate_template).convert("RGB")
    draw = ImageDraw.Draw(im)

    def center_text(text, y, font, anchor="ma"):
        draw.text((width / 2, y), text, fill=text_color, font=font, anchor=anchor)

    # --- Posiciones generales ---
    y_start = 810  # Centro exacto entre "Se certifica que" (680) y "Presentó el trabajo" (940)
    max_width = int(width * 0.85)

    # 1. Nombre
    center_text(expositor_name, y_start, expositor_font, anchor="mm")

    # 2. Título y Autores juntos (Rango de altura total libre: 980 a 1400)
    title = normalize_title_case(title)
    max_title_height = 200  # ajustado para que no se exceda
    title_font, title_lines = fit_text_to_box(
        title, bold_font_path, title_size, max_width, max_title_height, draw
    )
    
    max_authors_height = 130 # ajustado para los autores
    authors_font, authors_lines = fit_text_to_box(
        authors, italic_font_path, authors_size, max_width, max_authors_height, draw
    )
    
    # Calculamos el tamaño total del bloque para poder centrarlo matemáticamente
    title_spacing = 10
    authors_spacing = 10
    block_gap = 30
    
    # Usamos las estimaciones del tamaño vertical de toda la caja
    total_title_h = len(title_lines) * (title_font.size + title_spacing) - title_spacing
    total_authors_h = len(authors_lines) * (authors_font.size + authors_spacing) - authors_spacing
    total_block_h = total_title_h + block_gap + total_authors_h
    
    # El centro del espacio blanco libre está calculando entre Presentó (980) y la bajada inferior (1400) -> Mid = 1190
    y_start_block = 1190 - (total_block_h / 2)
    
    # Dibujar título
    y_current = y_start_block
    for i, line in enumerate(title_lines):
        center_text(line, y_current, title_font, anchor="ma")
        y_current += title_font.size + title_spacing
        
    # Salto hacia los autores
    y_current += block_gap - title_spacing
    
    # Dibujar autores
    for i, line in enumerate(authors_lines):
        center_text(line, y_current, authors_font, anchor="ma")
        y_current += authors_font.size + authors_spacing

    # Guardar
    filename = f'certificado_{eliminate_accents(expositor_name)}.pdf'
    im.save(os.path.join(save_path, filename))
    return filename


# === MAIN SCRIPT ===
if __name__ == "__main__":
    folder_path = "congreso_neurociencias_2026"
    csv_path = os.path.join(folder_path, "2026 - Presentación de póster.csv")
    certificate_template = os.path.join(folder_path, "certificate_poster.png")

    font_dir = os.path.join(folder_path, "nunito-sans")
    regular_font = os.path.join(font_dir, "NunitoSans-Regular.ttf")
    bold_font = os.path.join(font_dir, "NunitoSans-Bold.ttf")
    italic_font = os.path.join(font_dir, "NunitoSans-Italic.ttf")

    certificate_folder = os.path.join(folder_path, "certificados_expositores")
    os.makedirs(certificate_folder, exist_ok=True)

    text_color = "#000000"

    with Image.open(certificate_template) as img:
        width, _ = img.size

    # Tamaños máximos de partida
    expositor_size = 140
    title_size = 100
    authors_size = 70

    try:
        df = pd.read_csv(csv_path)

        for _, row in df.iterrows():
            title = row['TITULO DEL PÓSTER']
            expositor = row['PRESENTADOR']
            authors = row['AUTORES']

            if pd.isna(title) or pd.isna(expositor):
                continue

            try:
                filename = create_exposition_certificate(
                    certificate_template=certificate_template,
                    expositor_name=expositor,
                    title=title,
                    authors=authors,
                    text_color=text_color,
                    regular_font_path=regular_font,
                    bold_font_path=bold_font,
                    italic_font_path=italic_font,
                    expositor_size=expositor_size,
                    title_size=title_size,
                    authors_size=authors_size,
                    save_path=certificate_folder,
                    width=width
                )
                print(f"✅ Certificado creado: {filename}")
            except Exception as e:
                print(f"❌ Error creando certificado para {expositor}: {str(e)}")

    except Exception as e:
        print(f"❌ Error leyendo el CSV: {str(e)}")
