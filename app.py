"""
Backend Flask — App cuenta de cobro SED Medellín
Arquitectura asíncrona: las tareas largas corren en background,
el navegador consulta el progreso. Esto elimina los timeouts
sin importar cuántas evidencias se suban.
"""
import os, tempfile, threading, uuid, time, shutil
from flask import Flask, request, send_file, jsonify
from flask_cors import CORS

from informe_actividades import generar_informe
from cuenta_cobro import procesar_cuenta_cobro
from seguridad_social import combinar_ss
from empaquetar import empaquetar

app = Flask(__name__, static_folder='.', static_url_path='')
CORS(app)

# ── Almacén de tareas en memoria ──
# { task_id: {'status': 'processing'|'done'|'error', 'progress': str,
#             'file': ruta_pdf_o_None, 'error': str_o_None, 'dirname': str} }
TASKS = {}
TASKS_LOCK = threading.Lock()


def nueva_tarea() -> str:
    task_id = uuid.uuid4().hex
    with TASKS_LOCK:
        TASKS[task_id] = {'status': 'processing', 'progress': 'Iniciando...',
                          'file': None, 'error': None, 'filename': None}
    return task_id


def set_progreso(task_id, texto):
    with TASKS_LOCK:
        if task_id in TASKS:
            TASKS[task_id]['progress'] = texto


def set_listo(task_id, filepath, filename):
    with TASKS_LOCK:
        if task_id in TASKS:
            TASKS[task_id]['status'] = 'done'
            TASKS[task_id]['file'] = filepath
            TASKS[task_id]['filename'] = filename


def set_error(task_id, mensaje):
    with TASKS_LOCK:
        if task_id in TASKS:
            TASKS[task_id]['status'] = 'error'
            TASKS[task_id]['error'] = mensaje


@app.route('/')
def index():
    return app.send_static_file('index.html')

@app.route('/health')
def health():
    return jsonify({'status': 'ok'})


@app.route('/api/status/<task_id>')
def api_status(task_id):
    with TASKS_LOCK:
        t = TASKS.get(task_id)
    if not t:
        return jsonify({'error': 'Tarea no encontrada'}), 404
    return jsonify({
        'status': t['status'],
        'progress': t['progress'],
        'error': t['error'],
    })


@app.route('/api/download/<task_id>')
def api_download(task_id):
    with TASKS_LOCK:
        t = TASKS.get(task_id)
    if not t or t['status'] != 'done' or not t['file']:
        return jsonify({'error': 'Archivo no disponible'}), 404
    return send_file(t['file'], mimetype='application/pdf',
                     as_attachment=True, download_name=t['filename'])


# ── MÓDULO 1: Informe de actividades (asíncrono) ───────────────────
def _job_informe(task_id, tmp, excel_path, firma_path, csv_path,
                 archivos_evidencia, cedula, cobro):
    try:
        set_progreso(task_id, 'Leyendo datos del contratista...')
        pdf_out = os.path.join(tmp, f'{cedula}-Informe{cobro}.pdf')

        def progress_cb(msg):
            set_progreso(task_id, msg)

        generar_informe(
            excel_src=excel_path, cedula=cedula, cobro=cobro,
            csv_evidencias=csv_path, archivos_evidencia=archivos_evidencia,
            firma_img=firma_path, output_pdf=pdf_out, workdir=tmp,
            progress_cb=progress_cb,
        )
        set_listo(task_id, pdf_out, f'{cedula}-Informe{cobro}.pdf')
    except Exception as e:
        set_error(task_id, str(e))


@app.route('/api/informe', methods=['POST'])
def api_informe():
    try:
        cedula = int(request.form.get('cedula') or 0)
        cobro  = int(request.form.get('cobro') or 1)
        if not cedula:
            return jsonify({'error': 'La cédula es requerida'}), 400

        # Carpeta persistente para esta tarea (no se borra hasta completar)
        tmp = tempfile.mkdtemp(prefix='informe_')

        excel_path = os.path.join(tmp, 'formato.xlsx')
        firma_path = os.path.join(tmp, 'firma.png')
        request.files['excel'].save(excel_path)
        request.files['firma'].save(firma_path)

        csv_path = ''
        if 'csv' in request.files and request.files['csv'].filename:
            csv_path = os.path.join(tmp, 'evidencias.csv')
            request.files['csv'].save(csv_path)

        archivos_evidencia = []
        for key in request.files:
            if key.startswith('evidencia_'):
                f = request.files[key]
                if f.filename:
                    ext = os.path.splitext(f.filename)[1] or '.bin'
                    path = os.path.join(tmp, f'{key}{ext}')
                    f.save(path)
                    archivos_evidencia.append(path)

        task_id = nueva_tarea()
        thread = threading.Thread(
            target=_job_informe,
            args=(task_id, tmp, excel_path, firma_path, csv_path,
                  archivos_evidencia, cedula, cobro),
            daemon=True
        )
        thread.start()

        return jsonify({'task_id': task_id})
    except Exception as e:
        return jsonify({'error': str(e)}), 500


# ── MÓDULO 2: Cuenta de cobro (síncrono — siempre es rápido) ───────
@app.route('/api/cuenta', methods=['POST'])
def api_cuenta():
    try:
        cedula  = request.form.get('cedula', '').strip()
        cobro   = request.form.get('cobro', '1')
        if not cedula:
            return jsonify({'error': 'La cédula es requerida'}), 400
        version = request.form.get('version', 'valores')

        with tempfile.TemporaryDirectory() as tmp:
            pdf_path   = os.path.join(tmp, 'daniel.pdf')
            firma_path = os.path.join(tmp, 'firma.png')
            request.files['pdf'].save(pdf_path)
            request.files['firma'].save(firma_path)

            pdf_out = os.path.join(tmp, f'{cedula}-CuentaCobro{cobro}.pdf')
            procesar_cuenta_cobro(pdf_path, firma_path, version, pdf_out)

            return send_file(
                pdf_out, mimetype='application/pdf',
                as_attachment=True, download_name=f'{cedula}-CuentaCobro{cobro}.pdf'
            )
    except Exception as e:
        return jsonify({'error': str(e)}), 500


# ── MÓDULO 3: Seguridad social (síncrono — siempre es rápido) ──────
@app.route('/api/ss', methods=['POST'])
def api_ss():
    try:
        cedula = request.form.get('cedula', '').strip()
        cobro  = request.form.get('cobro', '1')
        if not cedula:
            return jsonify({'error': 'La cédula es requerida'}), 400

        with tempfile.TemporaryDirectory() as tmp:
            planilla_path    = os.path.join(tmp, 'planilla.pdf')
            comprobante_path = os.path.join(tmp, 'comprobante')

            request.files['planilla'].save(planilla_path)
            comp_file = request.files['comprobante']
            ext = os.path.splitext(comp_file.filename)[1].lower() or '.pdf'
            comprobante_path += ext
            comp_file.save(comprobante_path)

            pdf_out = os.path.join(tmp, f'{cedula}-SS{cobro}.pdf')
            combinar_ss(planilla_path, comprobante_path, pdf_out)

            return send_file(
                pdf_out, mimetype='application/pdf',
                as_attachment=True, download_name=f'{cedula}-SS{cobro}.pdf'
            )
    except Exception as e:
        return jsonify({'error': str(e)}), 500


# ── MÓDULO 4: Empaquetar (síncrono — siempre es rápido) ────────────
@app.route('/api/empaquetar', methods=['POST'])
def api_empaquetar():
    try:
        cedula = int(request.form.get('cedula') or 0)
        cobro  = int(request.form.get('cobro') or 1)
        if not cedula:
            return jsonify({'error': 'La cédula es requerida'}), 400

        with tempfile.TemporaryDirectory() as tmp:
            cuenta_path  = os.path.join(tmp, 'cuenta.pdf')
            informe_path = os.path.join(tmp, 'informe_vobo.pdf')
            ss_path      = os.path.join(tmp, 'ss.pdf')

            request.files['cuenta'].save(cuenta_path)
            request.files['informe'].save(informe_path)
            request.files['ss'].save(ss_path)

            pdf_out = os.path.join(tmp, f'{cedula}-CC{cobro}.pdf')
            empaquetar(
                cuenta_pdf=cuenta_path, informe_pdf=informe_path, ss_pdf=ss_path,
                cedula=cedula, cobro=cobro, output=pdf_out
            )

            return send_file(
                pdf_out, mimetype='application/pdf',
                as_attachment=True, download_name=f'{cedula}-CC{cobro}.pdf'
            )
    except Exception as e:
        return jsonify({'error': str(e)}), 500


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
