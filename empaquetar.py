"""
MÓDULO 4 — Empaquetador final
═══════════════════════════════════════════════════════════════
Combina todos los documentos en el orden correcto para SECOP/Drive.

ORDEN FINAL (igual al PDF 1128425027-CC5):
  1. Cuenta de cobro firmada (pág 1)       → del módulo 2
  2. Declaración juramentada firmada (pág 2) → del módulo 2
  3. Informe de actividades con VoBo supervisor (págs 3-4) → lo devuelve el supervisor
  4. Planilla SS + Comprobante (págs 5-6)  → del módulo 3

NOMBRE DEL ARCHIVO: {cedula}-CC{cobro}.pdf
  Ej: 1128425027-CC5.pdf

NOTAS:
  - La firma de la cuenta de cobro y del informe es la misma imagen
  - El informe que se empaqueta es el que devuelve el supervisor CON su VoBo
  - El PDF final debe ser < 100MB para subir a Drive/SECOP
  - Se comprime para reducir tamaño si es necesario

USO:
  python empaquetar.py \\
    --cuenta     "1128425027_CuentaCobro5_FIRMADA.pdf" \\
    --informe    "Informe_MAYO_con_VoBo.pdf" \\
    --ss         "1128425027_SS_Mayo2026.pdf" \\
    --cedula     1128425027 \\
    --cobro      5 \\
    --output     "1128425027-CC5.pdf"
"""

import argparse, os
from pathlib import Path
from pypdf import PdfReader, PdfWriter
from pypdf.generic import NameObject, NumberObject


def comprimir_pdf(writer: PdfWriter) -> PdfWriter:
    """Comprime el PDF para reducir tamaño."""
    for page in writer.pages:
        page.compress_content_streams()
    return writer


def empaquetar(
    cuenta_pdf: str,
    informe_pdf: str,
    ss_pdf: str,
    cedula: int,
    cobro: int,
    output: str = ''
) -> str:
    """
    Empaqueta todos los documentos en el orden correcto para SECOP.

    Args:
        cuenta_pdf:  PDF módulo 2 (cuenta de cobro + declaración firmadas)
        informe_pdf: PDF que devuelve el supervisor con VoBo (3 páginas)
        ss_pdf:      PDF módulo 3 (planilla SS + comprobante)
        cedula:      Cédula del contratista
        cobro:       Número de cobro
        output:      Ruta del PDF final (opcional)
    """
    # Nombre estándar para SECOP: {cedula}-CC{cobro}.pdf
    if not output:
        output = f'{cedula}-CC{cobro}.pdf'

    writer = PdfWriter()

    # 1. Cuenta de cobro + Declaración (módulo 2) → 2 páginas
    print(f'  [1] Cuenta de cobro + Declaración: {Path(cuenta_pdf).name}')
    r = PdfReader(cuenta_pdf)
    print(f'      {len(r.pages)} página(s)')
    for page in r.pages:
        writer.add_page(page)

    # 2. Informe con VoBo del supervisor → 3 páginas normalmente
    print(f'  [2] Informe con VoBo: {Path(informe_pdf).name}')
    r = PdfReader(informe_pdf)
    print(f'      {len(r.pages)} página(s)')
    for page in r.pages:
        writer.add_page(page)

    # 3. Planilla SS + Comprobante (módulo 3) → 2 páginas
    print(f'  [3] Planilla SS + Comprobante: {Path(ss_pdf).name}')
    r = PdfReader(ss_pdf)
    print(f'      {len(r.pages)} página(s)')
    for page in r.pages:
        writer.add_page(page)

    total_pages = len(writer.pages)

    # Comprimir para mantener tamaño bajo
    writer = comprimir_pdf(writer)

    with open(output, 'wb') as f:
        writer.write(f)

    size_mb = os.path.getsize(output) / (1024 * 1024)
    print(f'\n  ✓ PDF final: {output}')
    print(f'  ✓ Páginas: {total_pages}')
    print(f'  ✓ Tamaño: {size_mb:.2f} MB {"✓ OK" if size_mb < 100 else "⚠ > 100MB"}')

    return output


if __name__ == '__main__':
    parser = argparse.ArgumentParser(
        description='Módulo 4 — Empaquetador final para SECOP/Drive'
    )
    parser.add_argument('--cuenta',   required=True,
                        help='PDF cuenta de cobro + declaración firmadas (módulo 2)')
    parser.add_argument('--informe',  required=True,
                        help='PDF informe de actividades con VoBo del supervisor')
    parser.add_argument('--ss',       required=True,
                        help='PDF planilla SS + comprobante (módulo 3)')
    parser.add_argument('--cedula',   required=True, type=int,
                        help='Cédula del contratista')
    parser.add_argument('--cobro',    required=True, type=int,
                        help='Número de cobro')
    parser.add_argument('--output',   default='',
                        help='Ruta del PDF final (default: {cedula}-CC{cobro}.pdf)')
    args = parser.parse_args()

    print(f'\n{"═"*55}')
    print(f'  EMPAQUETANDO — Cédula {args.cedula} · Cobro #{args.cobro}')
    print(f'{"═"*55}')

    empaquetar(
        cuenta_pdf=args.cuenta,
        informe_pdf=args.informe,
        ss_pdf=args.ss,
        cedula=args.cedula,
        cobro=args.cobro,
        output=args.output
    )
