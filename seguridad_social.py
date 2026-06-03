"""
MÓDULO 3 — Planilla SS + Comprobante de pago
═══════════════════════════════════════════════
Recibe cualquier formato (PDF, PNG, JPG, JPEG) y los combina en un solo PDF.

USO:
  python seguridad_social.py \
    --planilla  "planilla.pdf" \
    --comprobante "comprobante.png" \
    --output "1128425027_SS_Mayo2026.pdf"
"""

import argparse
from pathlib import Path
from pypdf import PdfReader, PdfWriter
from PIL import Image
import io

def archivo_a_pdf_bytes(ruta: str) -> bytes:
    """Convierte cualquier archivo (PDF, PNG, JPG) a bytes de PDF."""
    ext = Path(ruta).suffix.lower()

    if ext == '.pdf':
        # Ya es PDF — devolver tal cual
        return open(ruta, 'rb').read()

    elif ext in ('.png', '.jpg', '.jpeg', '.webp', '.bmp', '.tiff'):
        # Imagen → convertir a PDF tamaño A4
        img = Image.open(ruta)

        # Convertir a RGB si tiene canal alpha (PNG)
        if img.mode in ('RGBA', 'LA', 'P'):
            fondo = Image.new('RGB', img.size, (255, 255, 255))
            if img.mode == 'P':
                img = img.convert('RGBA')
            fondo.paste(img, mask=img.split()[-1] if img.mode == 'RGBA' else None)
            img = fondo
        elif img.mode != 'RGB':
            img = img.convert('RGB')

        # Convertir imagen a PDF usando su tamaño natural
        # Sin escalar — calidad original, peso mínimo
        buf = io.BytesIO()
        img.save(buf, format='PDF', resolution=96)
        return buf.getvalue()

    else:
        raise ValueError(f'Formato no soportado: {ext}. Usa PDF, PNG o JPG.')


def combinar_ss(planilla: str, comprobante: str, output: str) -> str:
    """
    Combina planilla + comprobante en un solo PDF.
    Acepta PDF, PNG, JPG para cada archivo.
    """
    writer = PdfWriter()

    # Agregar planilla
    print(f'  Procesando planilla: {Path(planilla).name}')
    planilla_bytes = archivo_a_pdf_bytes(planilla)
    for page in PdfReader(io.BytesIO(planilla_bytes)).pages:
        writer.add_page(page)

    # Agregar comprobante
    print(f'  Procesando comprobante: {Path(comprobante).name}')
    comprobante_bytes = archivo_a_pdf_bytes(comprobante)
    for page in PdfReader(io.BytesIO(comprobante_bytes)).pages:
        writer.add_page(page)

    with open(output, 'wb') as f:
        writer.write(f)

    total = sum(1 for _ in PdfReader(output).pages)
    print(f'  ✓ PDF combinado: {total} página(s) → {output}')
    return output


if __name__ == '__main__':
    parser = argparse.ArgumentParser(
        description='Módulo 3 — Planilla SS + Comprobante de pago'
    )
    parser.add_argument('--planilla',     required=True, help='Planilla SS (PDF/PNG/JPG)')
    parser.add_argument('--comprobante',  required=True, help='Comprobante de pago (PDF/PNG/JPG)')
    parser.add_argument('--output',       default='SS_combinado.pdf', help='PDF de salida')
    args = parser.parse_args()

    combinar_ss(args.planilla, args.comprobante, args.output)
