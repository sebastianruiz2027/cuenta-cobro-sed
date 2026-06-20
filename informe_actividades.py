"""
═══════════════════════════════════════════════════════════════════════
 INFORME DE ACTIVIDADES — COLEGIO MAYOR / SED MEDELLÍN
 Documento: "Informe de actividades terminado y firmado" (GJ-FR-005)
═══════════════════════════════════════════════════════════════════════

FLUJO:
  1. Lee la cédula y número de informe → los escribe en el Excel
  2. El Excel carga todo automáticamente desde el CONSOLIDADO (VLOOKUP)
  3. Lee evidencias: CSV de correos (opcional) + hasta 10 archivos sueltos
     (fotos, PDFs, notas escaneadas, capturas, actas escritas a mano...)
  4. Manda TODO a Claude con visión → clasifica y redacta metas por actividad
  5. LibreOffice convierte el Excel lleno a PDF con formato oficial intacto
  6. Estampa la imagen de firma del contratista en "Firma del Contratista"
  7. Entrega: PDF firmado listo para VoBo supervisor y empaque final

APLICA PARA CUALQUIER CONTRATISTA del formato 106407:
  - Funciona independiente del número de actividades (5, 8, 10, 12...)
  - Lee actividades reales desde el CONSOLIDADO según la cédula
  - El "corazón" del análisis es la evidencia real (fotos/notas/actas),
    el CSV de correos es solo un complemento opcional

DEPENDENCIAS:
  pip install openpyxl pypdf reportlab
  apt install libreoffice

USO DESDE OTRO SCRIPT:
  from informe_actividades import generar_informe
  generar_informe(
      excel_src="Formato Informe.xlsx",
      csv_evidencias="calendario.csv",       # opcional, puede ser ''
      archivos_evidencia=["foto1.jpg", "acta.pdf", ...],  # hasta 10
      firma_img="firma.png",
      cedula=1128425027,
      cobro=5,
      output_pdf="informe_firmado.pdf"
  )
"""

import csv
import io
import re
import shutil
import subprocess
import argparse
import base64
import mimetypes
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font
from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas as rl_canvas

MAX_ARCHIVOS_EVIDENCIA = 10

SKIP_REMITENTES = [
    's@mi', 'innovación digital', 'apropia tic', 'microsoft',
    'sharepoint', 'google calendar', 'cscsoporte', 'outlook',
    'miro team', 'comunicación organizacional', 'comunicaciones educacion',
    'the miro team', 'spinach ai', 'fireflies', 'supernormal',
]
SKIP_ASUNTOS = [
    'mantenimiento', 'intermitencia', 'inactivación', 'indisponibilidad',
    'actualización de seguridad', 'migración a la nube', 'boletin',
    'encuesta de satisfacción', 'recuerde diligenciar', 'black day',
    'descuento', 'oferta', 'cancelar suscripción', 'levi',
    'resumen diario de reacciones',
]


# ═════════════════════════════════════════════════════════════════════
# PASO 0 — LEER DATOS DEL CONTRATISTA DESDE EL CONSOLIDADO
# ═════════════════════════════════════════════════════════════════════
def leer_contratista(excel_src: str, cedula: int) -> dict:
    """Lee nombre, contrato y actividades reales del contratista."""
    wb = openpyxl.load_workbook(excel_src, data_only=True, read_only=True)
    if 'CONSOLIDADO' not in wb.sheetnames:
        raise ValueError("El Excel no tiene hoja 'CONSOLIDADO'")

    wsc = wb['CONSOLIDADO']
    cedula_str = str(cedula).strip()

    for row in wsc.iter_rows(values_only=True):
        if str(row[0] or '').strip() == cedula_str:
            nombre = str(row[2] or '').strip()
            contrato = str(row[1] or '').strip()
            actividades = []
            for j in range(30, 55):
                if j < len(row):
                    v = row[j]
                    if v and str(v).strip():
                        actividades.append(str(v).strip())
                    else:
                        break
            wb.close()
            return {
                'nombre': nombre, 'contrato': contrato,
                'actividades': actividades, 'num_actividades': len(actividades),
            }
    wb.close()
    raise ValueError(f"Cédula {cedula} no encontrada en el CONSOLIDADO")


# ═════════════════════════════════════════════════════════════════════
# PASO 1 — CSV DE CORREOS (OPCIONAL, COMPLEMENTO)
# ═════════════════════════════════════════════════════════════════════
def leer_evidencias_csv(csv_path: str, max_correos: int = 60) -> list[str]:
    """
    Lee el CSV de calendario/correos y devuelve solo los asuntos relevantes
    (limitados para no saturar el prompt). Es un COMPLEMENTO, no el centro.
    """
    if not csv_path:
        return []
    asuntos = []
    try:
        with open(csv_path, encoding='utf-8-sig') as f:
            reader = csv.DictReader(f)
            for row in reader:
                asunto = row.get('Asunto', '').strip()
                remitente = row.get('De: (nombre)', '').strip().lower()
                if any(s in remitente for s in SKIP_REMITENTES):
                    continue
                if any(s in asunto.lower() for s in SKIP_ASUNTOS):
                    continue
                if not asunto or len(asunto) < 5:
                    continue
                asuntos.append(asunto)
                if len(asuntos) >= max_correos:
                    break
    except Exception as e:
        print(f'  ⚠ No se pudo leer el CSV: {e}')
    return asuntos


# ═════════════════════════════════════════════════════════════════════
# PASO 2 — ARCHIVOS DE EVIDENCIA SUELTOS (EL CORAZÓN DEL ANÁLISIS)
# ═════════════════════════════════════════════════════════════════════
def preparar_archivos_evidencia(rutas: list[str]) -> list[dict]:
    """
    Prepara hasta MAX_ARCHIVOS_EVIDENCIA archivos (fotos, PDFs, notas
    escaneadas, capturas, actas escritas a mano) como bloques base64
    listos para mandar a la API de Anthropic con visión.
    Las imágenes se redimensionan para que la llamada sea más rápida.
    """
    bloques = []
    for ruta in (rutas or [])[:MAX_ARCHIVOS_EVIDENCIA]:
        p = Path(ruta)
        if not p.exists():
            continue
        mime, _ = mimetypes.guess_type(str(p))
        if not mime:
            continue

        try:
            if mime in ('image/jpeg', 'image/png', 'image/gif', 'image/webp'):
                # Redimensionar para que la llamada a la API sea rápida
                from PIL import Image
                img = Image.open(p)
                if img.mode in ('RGBA', 'P'):
                    img = img.convert('RGB')
                # Máximo 1024px en el lado más largo (suficiente para leer texto)
                max_side = 1024
                if max(img.size) > max_side:
                    ratio = max_side / max(img.size)
                    new_size = (int(img.size[0]*ratio), int(img.size[1]*ratio))
                    img = img.resize(new_size, Image.LANCZOS)
                buf_img = io.BytesIO()
                img.save(buf_img, format='JPEG', quality=80)
                b64 = base64.b64encode(buf_img.getvalue()).decode('utf-8')
                bloques.append({
                    'type': 'image',
                    'source': {'type': 'base64', 'media_type': 'image/jpeg', 'data': b64}
                })
            elif mime == 'application/pdf':
                data = p.read_bytes()
                # Límite de tamaño para PDFs (5MB) — evita llamadas eternas
                if len(data) > 5 * 1024 * 1024:
                    print(f'  ⚠ {p.name} muy pesado (>5MB), se omite')
                    continue
                b64 = base64.b64encode(data).decode('utf-8')
                bloques.append({
                    'type': 'document',
                    'source': {'type': 'base64', 'media_type': 'application/pdf', 'data': b64}
                })
        except Exception as e:
            print(f'  ⚠ No se pudo procesar {p.name}: {e}')
    return bloques


# ═════════════════════════════════════════════════════════════════════
# PASO 3 — REDACTAR METAS CON CLAUDE (VISIÓN + TEXTO)
# ═════════════════════════════════════════════════════════════════════
def redactar_metas(actividades: list[str],
                   asuntos_correo: list[str],
                   bloques_evidencia: list[dict]) -> list[str]:
    """
    Manda las actividades + evidencias (imágenes/PDFs + asuntos de correo)
    a Claude para que clasifique y redacte 3 bullets profesionales por
    actividad. Claude lee fotos, notas a mano, actas escaneadas, etc.

    Si la API no responde, cae a un fallback genérico (sin evidencias falsas).
    """
    import urllib.request
    import json as _json
    import os as _os

    n_acts = len(actividades)
    api_key = _os.environ.get('ANTHROPIC_API_KEY', '')

    acts_txt = '\n'.join(f'{i+1}. {a}' for i, a in enumerate(actividades))
    correos_txt = ('\n'.join(f'- {a}' for a in asuntos_correo)
                   if asuntos_correo else '(sin correos adjuntos)')

    instrucciones = (
        "Eres asistente de un contratista de la Secretaría de Educación de Medellín.\n"
        "Te voy a dar:\n"
        "1) La lista de actividades específicas de su contrato\n"
        "2) Evidencias del mes: pueden ser fotos, actas, notas escritas a mano, "
        "capturas de pantalla, PDFs escaneados, o asuntos de correo\n\n"
        "Tu tarea: lee TODAS las evidencias (incluida cualquier imagen adjunta), "
        "identifica qué trabajo reflejan, y redacta exactamente 3 bullets de meta "
        "cumplida POR CADA actividad, basados en lo que realmente encuentres.\n\n"
        "Reglas:\n"
        "- Cada bullet empieza con \"• Se \" (voz pasiva impersonal)\n"
        "- Máximo 85 caracteres por bullet\n"
        "- Lenguaje formal, específico, profesional — nunca copies texto literal\n"
        "- Usa verbos variados: asesoró, revisó, apoyó, consolidó, articuló, gestionó\n"
        "- Si una actividad no tiene evidencia clara, redacta un bullet genérico "
        "coherente con su descripción (sin inventar datos falsos)\n\n"
        f"ACTIVIDADES DEL CONTRATO:\n{acts_txt}\n\n"
        f"ASUNTOS DE CORREO DEL MES (complemento):\n{correos_txt}\n\n"
        "Responde SOLO JSON sin markdown, en este formato exacto:\n"
        '{"metas": ["• Se ...\\n• Se ...\\n• Se ...", ...]}\n'
        f"El array debe tener exactamente {n_acts} elementos, uno por actividad, en orden."
    )

    content_blocks = [{'type': 'text', 'text': instrucciones}]
    content_blocks.extend(bloques_evidencia)  # imágenes / PDFs de evidencia

    modelos = [
        'claude-sonnet-4-5-20250929',
        'claude-sonnet-4-6',
        'claude-3-5-sonnet-20241022',
    ]

    if not api_key:
        print('  ⚠ ANTHROPIC_API_KEY no configurada — usando redacción genérica')
    else:
        for modelo in modelos:
            try:
                data = _json.dumps({
                    'model': modelo,
                    'max_tokens': 200 * n_acts,
                    'messages': [{'role': 'user', 'content': content_blocks}]
                }).encode('utf-8')
                req = urllib.request.Request(
                    'https://api.anthropic.com/v1/messages',
                    data=data,
                    headers={
                        'Content-Type': 'application/json',
                        'x-api-key': api_key,
                        'anthropic-version': '2023-06-01'
                    },
                    method='POST'
                )
                with urllib.request.urlopen(req, timeout=90) as resp:
                    result = _json.loads(resp.read())
                txt = (result['content'][0]['text']
                       .strip().replace('```json', '').replace('```', '').strip())
                metas = _json.loads(txt)['metas']
                if len(metas) == n_acts:
                    print(f'  ✓ Metas redactadas por IA ({modelo}) '
                          f'con {len(bloques_evidencia)} evidencia(s) visual(es)')
                    return metas
            except Exception as e:
                print(f'  ⚠ Modelo {modelo} falló: {e}')
                continue
        print('  ⚠ Ningún modelo respondió — usando redacción genérica')

    # Fallback genérico (sin datos inventados)
    metas = []
    for act in actividades:
        verbo = act.split()[0].lower() if act else 'apoyar'
        bullets = (
            f'• Se {verbo}aron las actividades establecidas en el contrato.\n'
            f'• Se dio seguimiento al cumplimiento de los objetivos del periodo.\n'
            f'• Se reportaron los avances correspondientes a la actividad.'
        )
        metas.append(bullets)
    return metas


# ═════════════════════════════════════════════════════════════════════
# PASO 4 — LLENAR EL EXCEL
# ═════════════════════════════════════════════════════════════════════
def llenar_excel(excel_src: str, cedula: int, cobro: int,
                 metas: list[str], excel_dst: str) -> None:
    shutil.copy(excel_src, excel_dst)
    wb = openpyxl.load_workbook(excel_dst)
    ws = wb['INFORME']
    ws['G14'] = cedula
    ws['E12'] = cobro

    COL_CHARS = 88
    FONT_SIZE = 10

    for i, meta in enumerate(metas):
        fila = 16 + i
        cell = ws[f'E{fila}']
        cell.value = meta
        cell.alignment = Alignment(wrap_text=True, vertical='top', horizontal='left')
        cell.font = Font(name='Calibri', size=FONT_SIZE)
        lineas_total = sum(max(1, -(-len(b) // COL_CHARS)) for b in meta.splitlines())
        altura = max(144, lineas_total * 15 + 10)
        ws.row_dimensions[fila].height = altura

    wb.save(excel_dst)
    print(f'  ✓ Excel llenado: G14={cedula}, E12={cobro}, '
          f'metas en E16:E{15+len(metas)}')


# ═════════════════════════════════════════════════════════════════════
# PASO 5 — CONVERTIR EXCEL A PDF CON LIBREOFFICE
# ═════════════════════════════════════════════════════════════════════
def excel_a_pdf(excel_path: str, output_dir: str) -> str:
    result = subprocess.run(
        ['libreoffice', '--headless', '--convert-to', 'pdf',
         excel_path, '--outdir', output_dir],
        capture_output=True, text=True, timeout=150
    )
    if result.returncode != 0:
        raise RuntimeError(f'LibreOffice falló: {result.stderr}')
    pdf_path = str(Path(output_dir) / (Path(excel_path).stem + '.pdf'))
    print(f'  ✓ Convertido a PDF: {pdf_path}')
    return pdf_path


# ═════════════════════════════════════════════════════════════════════
# PASO 6 — ESTAMPAR FIRMA EN "FIRMA DEL CONTRATISTA" (última página)
# ═════════════════════════════════════════════════════════════════════
def estampar_firma(pdf_src: str, firma_img: str,
                   pdf_dst: str, cedula: int, nombre: str) -> None:
    """
    Coordenadas medidas del PDF real generado (595x841pt):
      Línea de firma en y=814.5 desde arriba
      Texto anterior termina en y=778.0 → espacio libre = 36.5pt
    """
    reader = PdfReader(pdf_src)
    writer = PdfWriter()
    total = len(reader.pages)

    for i in range(total - 1):
        writer.add_page(reader.pages[i])

    last = reader.pages[-1]
    page_w = float(last.mediabox.width)
    page_h = float(last.mediabox.height)

    buf = io.BytesIO()
    c = rl_canvas.Canvas(buf, pagesize=(page_w, page_h))

    firma_path = Path(firma_img) if firma_img else None

    if firma_path and firma_path.exists():
        firma_w, firma_h = 165, 32
        firma_x = 42
        firma_y = page_h - 814.5
        c.drawImage(
            str(firma_path), firma_x, firma_y,
            width=firma_w, height=firma_h,
            preserveAspectRatio=True, mask='auto'
        )
        print(f'  ✓ Firma imagen estampada desde: {firma_path}')
    else:
        c.setFont('Helvetica-Oblique', 11)
        c.setFillColorRGB(0.05, 0.05, 0.4)
        c.drawString(60, page_h - 803, nombre)
        c.setFont('Helvetica', 7)
        c.setFillColorRGB(0.3, 0.3, 0.3)
        c.drawString(60, page_h - 814, f'C.C. {cedula}')
        print(f'  ⚠ Sin imagen de firma — usando texto como fallback')

    c.save()
    buf.seek(0)
    overlay = PdfReader(buf).pages[0]
    last.merge_page(overlay)
    writer.add_page(last)

    with open(pdf_dst, 'wb') as f:
        writer.write(f)
    print(f'  ✓ PDF firmado guardado: {pdf_dst}')


# ═════════════════════════════════════════════════════════════════════
# FUNCIÓN PRINCIPAL — ORQUESTA TODO EL PROCESO
# ═════════════════════════════════════════════════════════════════════
def generar_informe(
    excel_src: str,
    cedula: int,
    cobro: int,
    csv_evidencias: str = '',
    archivos_evidencia: list[str] = None,
    firma_img: str = '',
    output_pdf: str = '',
    workdir: str = '/tmp',
    progress_cb=None,
) -> str:
    """progress_cb: función opcional(str) para reportar progreso en vivo"""
    def report(msg):
        print(msg)
        if progress_cb:
            try:
                progress_cb(msg)
            except Exception:
                pass

    """
    Genera el 'Informe de actividades terminado y firmado'.

    Args:
        excel_src:          Ruta al Excel del formato (con hoja CONSOLIDADO)
        cedula:              Cédula del contratista
        cobro:               Número de informe/cobro
        csv_evidencias:      Ruta al CSV de correos (OPCIONAL, complemento)
        archivos_evidencia:  Lista de hasta 10 rutas: fotos, PDFs, notas
                             escaneadas, actas, capturas — el corazón del análisis
        firma_img:           Ruta a PNG/JPG de la firma dibujada
        output_pdf:          Ruta del PDF de salida
        workdir:             Directorio temporal
    """
    archivos_evidencia = archivos_evidencia or []

    mes_label = f'Informe{cobro}'
    if not output_pdf:
        output_pdf = str(Path(workdir) / f'{cedula}_{mes_label}_FIRMADO.pdf')

    print(f'\n{"═"*60}')
    print(f'  INFORME DE ACTIVIDADES — Cédula {cedula} · Cobro #{cobro}')
    print(f'{"═"*60}')

    report('Leyendo datos del contratista...')
    datos = leer_contratista(excel_src, cedula)
    print(f'  Contratista: {datos["nombre"]}')
    print(f'  Contrato:    {datos["contrato"]}')
    print(f'  Actividades: {datos["num_actividades"]}')

    report('Preparando evidencias...')
    asuntos = leer_evidencias_csv(csv_evidencias)
    print(f'  Correos (complemento): {len(asuntos)}')
    bloques = preparar_archivos_evidencia(archivos_evidencia)
    print(f'  Archivos de evidencia: {len(bloques)} de {len(archivos_evidencia)} subidos '
          f'(máx {MAX_ARCHIVOS_EVIDENCIA})')

    report('Analizando evidencias y redactando metas con IA...')
    metas = redactar_metas(datos['actividades'], asuntos, bloques)

    report('Llenando formato Excel...')
    excel_tmp = str(Path(workdir) / f'{cedula}_Informe{cobro}_tmp.xlsx')
    llenar_excel(excel_src, cedula, cobro, metas, excel_tmp)

    pdf_tmp = excel_a_pdf(excel_tmp, workdir)

    report('Convirtiendo a PDF y estampando firma...')
    estampar_firma(pdf_tmp, firma_img, output_pdf, cedula, datos['nombre'])

    print(f'\n{"═"*60}')
    print(f'  ✅ LISTO: {output_pdf}')
    print(f'{"═"*60}\n')

    return output_pdf


# ═════════════════════════════════════════════════════════════════════
# CLI
# ═════════════════════════════════════════════════════════════════════
if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Genera el informe de actividades firmado (GJ-FR-005).')
    parser.add_argument('--excel', required=True)
    parser.add_argument('--cedula', required=True, type=int)
    parser.add_argument('--cobro', required=True, type=int)
    parser.add_argument('--csv', default='')
    parser.add_argument('--evidencias', nargs='*', default=[],
                        help='Hasta 10 rutas de fotos/PDFs de evidencia')
    parser.add_argument('--firma', default='')
    parser.add_argument('--output', default='')
    parser.add_argument('--workdir', default='/tmp')
    args = parser.parse_args()

    generar_informe(
        excel_src=args.excel, cedula=args.cedula, cobro=args.cobro,
        csv_evidencias=args.csv, archivos_evidencia=args.evidencias,
        firma_img=args.firma, output_pdf=args.output, workdir=args.workdir,
    )
