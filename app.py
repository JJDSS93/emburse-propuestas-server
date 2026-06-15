from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
from pptx import Presentation
from pptx.util import Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
import io, json

app = Flask(__name__)
CORS(app)

def hex_color(h):
    return RGBColor(int(h[0:2],16), int(h[2:4],16), int(h[4:6],16))

AZUL   = hex_color('0D2240')
AZUL2  = hex_color('1E6DB5')
BLANCO = hex_color('FFFFFF')
GRIS   = hex_color('666666')
PAR    = hex_color('EDF2F7')
IMPAR  = hex_color('FFFFFF')

def clear_slide(slide):
    for shape in list(slide.shapes):
        slide.shapes._spTree.remove(shape._element)

def add_textbox(slide, text, x, y, w, h, bold=False, size=32, color=None, align=PP_ALIGN.LEFT):
    if color is None: color = AZUL
    tb = slide.shapes.add_textbox(Emu(x), Emu(y), Emu(w), Emu(h))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.bold = bold
    run.font.size = Pt(size)
    run.font.color.rgb = color
    run.font.name = 'Calibri'
    return tb

def add_table(slide, headers, rows, x, y, w, h, col_widths):
    n_cols = len(headers)
    n_rows = len(rows) + 1
    tbl = slide.shapes.add_table(n_rows, n_cols, Emu(x), Emu(y), Emu(w), Emu(h)).table
    for i, cw in enumerate(col_widths):
        tbl.columns[i].width = Emu(cw)
    for j, ht in enumerate(headers):
        cell = tbl.cell(0, j)
        cell.text = ht
        cell.fill.solid()
        cell.fill.fore_color.rgb = AZUL
        p = cell.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        if p.runs:
            run = p.runs[0]
        else:
            run = p.add_run()
        run.font.bold = True
        run.font.color.rgb = BLANCO
        run.font.size = Pt(11)
        run.font.name = 'Calibri'
    for i, row in enumerate(rows):
        bg = PAR if i % 2 == 0 else IMPAR
        for j, val in enumerate(row):
            cell = tbl.cell(i+1, j)
            cell.text = str(val)
            cell.fill.solid()
            cell.fill.fore_color.rgb = bg
            p = cell.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.LEFT
            if p.runs:
                run = p.runs[0]
            else:
                run = p.add_run()
            run.font.color.rgb = AZUL
            run.font.size = Pt(11)
            run.font.name = 'Calibri'

@app.route('/health')
def health():
    return jsonify({'status': 'ok'})

@app.route('/generar', methods=['POST', 'OPTIONS'])
def generar():
    if request.method == 'OPTIONS':
        resp = app.make_default_options_response()
        resp.headers['Access-Control-Allow-Origin'] = '*'
        resp.headers['Access-Control-Allow-Headers'] = 'Content-Type'
        resp.headers['Access-Control-Allow-Methods'] = 'POST, OPTIONS'
        return resp

    try:
        pptx_file = request.files.get('plantilla')
        if not pptx_file:
            return jsonify({'error': 'No se recibió la plantilla'}), 400

        datos = json.loads(request.form.get('datos', '{}'))
        empresa       = datos.get('empresa', 'Cliente')
        moneda        = datos.get('moneda', 'USD')
        licencias     = datos.get('licencias', [])
        setup         = datos.get('setup', [])
        opcionales    = datos.get('opcionales', [])
        notas         = datos.get('notas', 'Precios en USD. IVA no incluido. Vigencia 30 días.')
        fecha         = datos.get('fecha', '')
        leyenda_lic   = (datos.get('leyenda_lic') or '').strip()
        leyenda_setup = (datos.get('leyenda_setup') or '').strip()
        leyenda_opc   = (datos.get('leyenda_opc') or '').strip()

        def fmt(n):
            sym = '$' if moneda in ('USD','MXN') else '€'
            return f"{sym}{float(n):,.2f} {moneda}"

        prs = Presentation(io.BytesIO(pptx_file.read()))

        lic_rows = [[r.get('nombre',''), str(r.get('vol','')),
            fmt(r['precio']) if r.get('precio') else '',
            fmt(float(r.get('vol') or 0)*float(r.get('precio') or 0)) if r.get('nombre') else '']
            for r in licencias]

        setup_rows = [[r.get('servicio',''), r.get('descripcion',''), r.get('tipo','Único'),
            fmt(r['precio']) if r.get('precio') else ''] for r in setup]

        opc_rows = [[r.get('servicio',''), r.get('descripcion',''), r.get('tipo','Anual'),
            fmt(r['precio']) if r.get('precio') else ''] for r in opcionales]

        lic_total   = sum(float(r.get('vol') or 0)*float(r.get('precio') or 0) for r in licencias)
        setup_total = sum(float(r.get('precio') or 0) for r in setup)
        opc_total   = sum(float(r.get('precio') or 0) for r in opcionales)
        res_rows = [
            ["Escenario 1 — Base", "Licencias únicamente", fmt(lic_total)],
            ["Escenario 2 — Completo", "Licencias + Setup & Onboarding", fmt(lic_total+setup_total)],
            ["Escenario 3 — Premium", "Licencias + Setup + Servicios Opcionales", fmt(lic_total+setup_total+opc_total)],
        ]

        s19 = prs.slides[18]
        clear_slide(s19)
        add_textbox(s19, "Licencias", 457200, 304800, 8229600, 533400, bold=True, size=32)
        add_table(s19, ["Producto / Licencia","Volumen",f"Precio unit. ({moneda})",f"Total ({moneda})"],
                  lic_rows, 457200, 990600, 8229600, 2200000, [2500000,1900000,1900000,1700000])
        add_textbox(s19, leyenda_lic, 457200, 3300000, 8229600, 304800, size=10, color=GRIS)

        s20 = prs.slides[19]
        clear_slide(s20)
        add_textbox(s20, "Setup & Onboarding", 457200, 304800, 8229600, 533400, bold=True, size=32)
        add_table(s20, ["Servicio","Descripción","Tipo de pago",f"Precio ({moneda})"],
                  setup_rows, 457200, 990600, 8229600, 2200000, [2000000,3100000,1500000,1629600])
        add_textbox(s20, leyenda_setup, 457200, 3300000, 8229600, 304800, size=10, color=GRIS)

        s21 = prs.slides[20]
        clear_slide(s21)
        add_textbox(s21, "Servicios Opcionales", 457200, 304800, 8229600, 533400, bold=True, size=32)
        add_table(s21, ["Servicio","Descripción","Tipo de pago",f"Precio ({moneda})"],
                  opc_rows, 457200, 990600, 8229600, 2200000, [2000000,3100000,1500000,1629600])
        add_textbox(s21, leyenda_opc, 457200, 3300000, 8229600, 304800, size=10, color=GRIS)

        s22 = prs.slides[21]
        clear_slide(s22)
        add_textbox(s22, "Resumen", 457200, 304800, 8229600, 533400, bold=True, size=32)
        add_table(s22, ["Escenario","Descripción",f"Inversión Año 1 ({moneda})"],
                  res_rows, 457200, 990600, 8229600, 1800000, [2300000,3800000,2129600])
        add_textbox(s22, "Condiciones y notas adicionales:", 457200, 2895600, 8229600, 304800, bold=True, size=14, color=AZUL2)
        notas_text = "\n".join(f"• {l.strip()}" for l in notas.split('\n') if l.strip())
        add_textbox(s22, notas_text, 457200, 3200400, 8229600, 700000, size=10, color=GRIS)

        out = io.BytesIO()
        prs.save(out)
        out.seek(0)

        fecha_str = (fecha or '').replace('-','')
        filename = f"Propuesta_Emburse_{empresa.replace(' ','_')}_{fecha_str}.pptx"

        resp = send_file(out,
            mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation',
            as_attachment=True, download_name=filename)
        resp.headers['Access-Control-Allow-Origin'] = '*'
        return resp

    except Exception as e:
        import traceback
        print(traceback.format_exc())
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    import os
    port = int(os.environ.get('PORT', 8080))
    app.run(host='0.0.0.0', port=port)
