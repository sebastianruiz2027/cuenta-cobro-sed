"""
MÓDULO 2 — Cuenta de cobro + Declaración juramentada
═══════════════════════════════════════════════════════
FLUJO:
  1. Recibe el PDF de Daniel (3 páginas)
  2. El usuario elige: declaración CON valores o CON ceros
  3. Estampa la firma en página 1 (cuenta) y en la página elegida (declaración)
  4. Entrega PDF de 2 páginas: cuenta + declaración elegida

COORDENADAS (medidas del PDF real 595x841pt):
  Página 1 — Cuenta de cobro:
    Firma sobre línea y=693.6 → espacio libre y=529→693 (~164pt)
    Campo firma: x=64→375, firma en x=66, y_base=694 (sobre la línea)

  Página 2 — Declaración CON valores:
    Firma sobre guiones y=685.8
    Campo firma: x=55→250, firma en x=57, y_base=686

  Página 3 — Declaración CON ceros:
    Mismas coordenadas que página 2

USO:
  python cuenta_cobro.py \
    --pdf    "1128425027_contratista_C79488_5.pdf" \
    --firma  "firma.png" \
    --version valores|ceros \
    --output "1128425027_CuentaCobro5_FIRMADA.pdf"
"""

import io, argparse
from pathlib import Path
from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas as rl_canvas

# ── COORDENADAS DE FIRMA ──
# Página 1 — Cuenta de cobro
# Espacio libre: entre y=529 (nombre) y y=693 (línea) = ~164pt
# Firma va en el centro de ese espacio
CUENTA_FIRMA_X    = 66
CUENTA_FIRMA_Y_TOP = 540   # desde arriba (debajo del nombre)
CUENTA_FIRMA_H    = 80     # alto imagen
CUENTA_FIRMA_W    = 200    # ancho imagen

# Páginas 2 y 3 — Declaraciones
# Guiones en y=685.8, firma va encima
DECL_FIRMA_X     = 57
DECL_FIRMA_Y_TOP = 640    # desde arriba (suficiente espacio encima de los guiones)
DECL_FIRMA_H     = 48
DECL_FIRMA_W     = 180


def estampar_firma_en_pagina(page, firma_path: str, x: float, y_top: float,
                              w: float, h: float) -> object:
    """
    Estampa la imagen de firma en una página.
    y_top: posición desde arriba de la página (pdfplumber coords)
    Convierte a coords ReportLab (desde abajo) internamente.
    """
    page_h = float(page.mediabox.height)
    page_w = float(page.mediabox.width)

    # Convertir y_top (desde arriba) a y_rl (desde abajo)
    y_rl = page_h - y_top - h

    buf = io.BytesIO()
    c = rl_canvas.Canvas(buf, pagesize=(page_w, page_h))
    c.drawImage(
        firma_path, x, y_rl,
        width=w, height=h,
        preserveAspectRatio=True,
        mask='auto'
    )
    c.save()
    buf.seek(0)

    from pypdf import PdfReader as PR
    overlay = PR(buf).pages[0]
    page.merge_page(overlay)
    return page


def procesar_cuenta_cobro(pdf_src: str, firma_path: str,
                           version: str, pdf_dst: str) -> str:
    """
    Procesa el PDF de Daniel:
    - Firma la página 1 (cuenta de cobro)
    - Elige página 2 (con valores) o página 3 (con ceros)
    - Firma la declaración elegida
    - Entrega PDF de 2 páginas

    Args:
        pdf_src:   PDF original de Daniel (3 páginas)
        firma_path: imagen PNG/JPG de la firma
        version:   'valores' o 'ceros'
        pdf_dst:   ruta del PDF de salida
    """
    if version not in ('valores', 'ceros'):
        raise ValueError("version debe ser 'valores' o 'ceros'")

    reader = PdfReader(pdf_src)
    writer = PdfWriter()

    # Página 0 = cuenta de cobro → firmar
    pag_cuenta = reader.pages[0]
    pag_cuenta = estampar_firma_en_pagina(
        pag_cuenta, firma_path,
        CUENTA_FIRMA_X, CUENTA_FIRMA_Y_TOP,
        CUENTA_FIRMA_W, CUENTA_FIRMA_H
    )
    writer.add_page(pag_cuenta)

    # Página 1 = declaración CON valores
    # Página 2 = declaración CON ceros
    idx_decl = 1 if version == 'valores' else 2
    pag_decl = reader.pages[idx_decl]
    pag_decl = estampar_firma_en_pagina(
        pag_decl, firma_path,
        DECL_FIRMA_X, DECL_FIRMA_Y_TOP,
        DECL_FIRMA_W, DECL_FIRMA_H
    )
    writer.add_page(pag_decl)

    with open(pdf_dst, 'wb') as f:
        writer.write(f)

    label = 'CON valores SS' if version == 'valores' else 'CON ceros'
    print(f'✓ Cuenta de cobro firmada ({label}): {pdf_dst}')
    return pdf_dst


# ── CLI ──
if __name__ == '__main__':
    parser = argparse.ArgumentParser(
        description='Módulo 2 — Cuenta de cobro + Declaración juramentada'
    )
    parser.add_argument('--pdf',     required=True, help='PDF de Daniel (3 páginas)')
    parser.add_argument('--firma',   required=True, help='Imagen PNG/JPG de la firma')
    parser.add_argument('--version', default='valores',
                        choices=['valores','ceros'],
                        help='Declaración: con valores SS (default) o con ceros')
    parser.add_argument('--output',  default='',    help='Ruta del PDF de salida')
    args = parser.parse_args()

    cedula = Path(args.pdf).stem.split('_')[0]
    output = args.output or f'{cedula}_CuentaCobro_FIRMADA.pdf'

    procesar_cuenta_cobro(args.pdf, args.firma, args.version, output)
