"""
Backend Flask — App cuenta de cobro SED Medellín
Integra los 4 módulos: informe, cuenta cobro, SS, empaquetador
"""
import os, tempfile
from flask import Flask, request, send_file, jsonify
from flask_cors import CORS

from informe_actividades import generar_informe
from cuenta_cobro import procesar_cuenta_cobro
from seguridad_social import combinar_ss
from empaquetar import empaquetar

app = Flask(__name__, static_folder='.', static_url_path='')
CORS(app)

@app.route('/')
def index():
    return app.send_static_file('index.html')

@app.route('/health')
def health():
    return jsonify({'status': 'ok'})


# ── MÓDULO 1: Informe de actividades ──────────────────────────────
@app.route('/api/informe', methods=['POST'])
def api_informe():
    try:
        cedula = int(request.form.get('cedula') or 0)
        cobro  = int(request.form.get('cobro', 1))

        with tempfile.TemporaryDirectory() as tmp:
            # Guardar archivos
            excel_path = os.path.join(tmp, 'formato.xlsx')
            firma_path = os.path.join(tmp, 'firma.png')
            request.files['excel'].save(excel_path)
            request.files['firma'].save(firma_path)

            csv_path = None
            if 'csv' in request.files:
                csv_path = os.path.join(tmp, 'evidencias.csv')
                request.files['csv'].save(csv_path)

            # Generar informe
            pdf_out = os.path.join(tmp, f'{cedula}-Informe{cobro}.pdf')
            generar_informe(
                excel_src=excel_path,
                csv_evidencias=csv_path or '',
                firma_img=firma_path,
                cedula=cedula,
                cobro=cobro,
                output_pdf=pdf_out,
                workdir=tmp,
            )

            return send_file(
                pdf_out,
                mimetype='application/pdf',
                as_attachment=True,
                download_name=f'{cedula}-Informe{cobro}.pdf'
            )
    except Exception as e:
        return jsonify({'error': str(e)}), 500


# ── MÓDULO 2: Cuenta de cobro ──────────────────────────────────────
@app.route('/api/cuenta', methods=['POST'])
def api_cuenta():
    try:
        cedula  = request.form.get('cedula', '').strip()
        cobro   = request.form.get('cobro', '1')
        if not cedula:
            return jsonify({'error': 'La cédula es requerida'}), 400
        version = request.form.get('version', 'valores')  # 'valores' o 'ceros'

        with tempfile.TemporaryDirectory() as tmp:
            pdf_path   = os.path.join(tmp, 'daniel.pdf')
            firma_path = os.path.join(tmp, 'firma.png')
            request.files['pdf'].save(pdf_path)
            request.files['firma'].save(firma_path)

            pdf_out = os.path.join(tmp, f'{cedula}-CuentaCobro{cobro}.pdf')
            procesar_cuenta_cobro(pdf_path, firma_path, version, pdf_out)

            return send_file(
                pdf_out,
                mimetype='application/pdf',
                as_attachment=True,
                download_name=f'{cedula}-CuentaCobro{cobro}.pdf'
            )
    except Exception as e:
        return jsonify({'error': str(e)}), 500


# ── MÓDULO 3: Seguridad social ─────────────────────────────────────
@app.route('/api/ss', methods=['POST'])
def api_ss():
    try:
        cedula = request.form.get('cedula', '').strip()
        cobro  = request.form.get('cobro', '1')
        if not cedula:
            return jsonify({'error': 'La cédula es requerida'}), 400

        with tempfile.TemporaryDirectory() as tmp:
            planilla_path     = os.path.join(tmp, 'planilla.pdf')
            comprobante_path  = os.path.join(tmp, 'comprobante')

            request.files['planilla'].save(planilla_path)

            # Comprobante puede ser PDF, JPG, PNG
            comp_file = request.files['comprobante']
            ext = os.path.splitext(comp_file.filename)[1].lower() or '.pdf'
            comprobante_path += ext
            comp_file.save(comprobante_path)

            pdf_out = os.path.join(tmp, f'{cedula}-SS{cobro}.pdf')
            combinar_ss(planilla_path, comprobante_path, pdf_out)

            return send_file(
                pdf_out,
                mimetype='application/pdf',
                as_attachment=True,
                download_name=f'{cedula}-SS{cobro}.pdf'
            )
    except Exception as e:
        return jsonify({'error': str(e)}), 500


# ── MÓDULO 4: Empaquetar ───────────────────────────────────────────
@app.route('/api/empaquetar', methods=['POST'])
def api_empaquetar():
    try:
        cedula = int(request.form.get('cedula') or 0)
        cobro  = int(request.form.get('cobro', 1))

        with tempfile.TemporaryDirectory() as tmp:
            cuenta_path  = os.path.join(tmp, 'cuenta.pdf')
            informe_path = os.path.join(tmp, 'informe_vobo.pdf')
            ss_path      = os.path.join(tmp, 'ss.pdf')

            request.files['cuenta'].save(cuenta_path)
            request.files['informe'].save(informe_path)
            request.files['ss'].save(ss_path)

            pdf_out = os.path.join(tmp, f'{cedula}-CC{cobro}.pdf')
            empaquetar(
                cuenta_pdf=cuenta_path,
                informe_pdf=informe_path,
                ss_pdf=ss_path,
                cedula=cedula,
                cobro=cobro,
                output=pdf_out
            )

            return send_file(
                pdf_out,
                mimetype='application/pdf',
                as_attachment=True,
                download_name=f'{cedula}-CC{cobro}.pdf'
            )
    except Exception as e:
        return jsonify({'error': str(e)}), 500


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
