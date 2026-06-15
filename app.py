from flask import Flask, request, jsonify, send_file
import zipfile, io, json

app = Flask(__name__)

def esc(s):
    return str(s or '').replace('&','&amp;').replace('<','&lt;').replace('>','&gt;').replace('"','&quot;')

NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main"'

RELS = '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>'

def hcell(t):
    return f'<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="es-MX" sz="1100" b="1" dirty="0"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:latin typeface="Calibri"/></a:rPr><a:t>{esc(t)}</a:t></a:r></a:p></a:txBody><a:tcPr marL="91440" marR="91440" marT="45720" marB="45720"><a:lnL w="12700"><a:solidFill><a:srgbClr val="0D2240"/></a:solidFill></a:lnL><a:lnR w="12700"><a:solidFill><a:srgbClr val="0D2240"/></a:solidFill></a:lnR><a:lnT w="12700"><a:solidFill><a:srgbClr val="0D2240"/></a:solidFill></a:lnT><a:lnB w="12700"><a:solidFill><a:srgbClr val="0D2240"/></a:solidFill></a:lnB><a:solidFill><a:srgbClr val="0D2240"/></a:solidFill></a:tcPr></a:tc>'

def dcell(t, bg):
    return f'<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:pPr algn="l"/><a:r><a:rPr lang="es-MX" sz="1100" dirty="0"><a:solidFill><a:srgbClr val="1A2740"/></a:solidFill><a:latin typeface="Calibri"/></a:rPr><a:t>{esc(t)}</a:t></a:r></a:p></a:txBody><a:tcPr marL="91440" marR="91440" marT="45720" marB="45720"><a:lnL w="9525"><a:solidFill><a:srgbClr val="CCCCCC"/></a:solidFill></a:lnL><a:lnR w="9525"><a:solidFill><a:srgbClr val="CCCCCC"/></a:solidFill></a:lnR><a:lnT w="9525"><a:solidFill><a:srgbClr val="CCCCCC"/></a:solidFill></a:lnT><a:lnB w="9525"><a:solidFill><a:srgbClr val="CCCCCC"/></a:solidFill></a:lnB><a:solidFill><a:srgbClr val="{bg}"/></a:solidFill></a:tcPr></a:tc>'

def make_table(x, y, cx, cy, widths, headers, rows):
    grid = "".join(f'<a:gridCol w="{w}"/>' for w in widths)
    hrow = f'<a:tr h="450000">{"".join(hcell(h) for h in headers)}</a:tr>'
    drows = "".join(f'<a:tr h="380000">{"".join(dcell(c, "EDF2F7" if i%2==0 else "FFFFFF") for c in row)}</a:tr>' for i,row in enumerate(rows))
    return f'<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="300" name="tabla"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr firstRow="1" bandRow="1"/><a:tblGrid>{grid}</a:tblGrid>{hrow}{drows}</a:tbl></a:graphicData></a:graphic></p:graphicFrame>'

def title_sp(t):
    return f'<p:sp><p:nvSpPr><p:cNvPr id="200" name="tit"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="457200" y="304800"/><a:ext cx="8229600" cy="533400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr><p:txBody><a:bodyPr anchor="ctr" bIns="0" lIns="0" rIns="0" tIns="0" wrap="square"><a:noAutofit/></a:bodyPr><a:lstStyle/><a:p><a:pPr algn="l"><a:buNone/></a:pPr><a:r><a:rPr b="1" lang="es-MX" sz="3200" dirty="0"><a:solidFill><a:srgbClr val="0D2240"/></a:solidFill><a:latin typeface="Calibri"/></a:rPr><a:t>{esc(t)}</a:t></a:r></a:p></a:txBody></p:sp>'

def div_sp():
    return '<p:sp><p:nvSpPr><p:cNvPr id="201" name="div"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="457200" y="876300"/><a:ext cx="8229600" cy="22860"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:linearGradFill rot="5400000" scaled="0"><a:gsLst><a:gs pos="0"><a:srgbClr val="C8D400"/></a:gs><a:gs pos="50000"><a:srgbClr val="00B4D4"/></a:gs><a:gs pos="100000"><a:srgbClr val="1E4FD8"/></a:gs></a:gsLst></a:linearGradFill></a:solidFill><a:ln><a:noFill/></a:ln></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>'

def note_sp(id, t, y):
    return f'<p:sp><p:nvSpPr><p:cNvPr id="{id}" name="n{id}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="457200" y="{y}"/><a:ext cx="8229600" cy="304800"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr><p:txBody><a:bodyPr bIns="0" lIns="91440" rIns="91440" tIns="0" wrap="square"/><a:lstStyle/><a:p><a:pPr algn="l"><a:buNone/></a:pPr><a:r><a:rPr lang="es-MX" sz="1000" i="1" dirty="0"><a:solidFill><a:srgbClr val="666666"/></a:solidFill><a:latin typeface="Calibri"/></a:rPr><a:t>{esc(t)}</a:t></a:r></a:p></a:txBody></p:sp>'

def label_sp(id, t, y):
    return f'<p:sp><p:nvSpPr><p:cNvPr id="{id}" name="l{id}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="457200" y="{y}"/><a:ext cx="8229600" cy="304800"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr><p:txBody><a:bodyPr bIns="0" lIns="91440" rIns="91440" tIns="0" wrap="square"/><a:lstStyle/><a:p><a:pPr algn="l"><a:buNone/></a:pPr><a:r><a:rPr lang="es-MX" sz="1400" b="1" dirty="0"><a:solidFill><a:srgbClr val="1E6DB5"/></a:solidFill><a:latin typeface="Calibri"/></a:rPr><a:t>{esc(t)}</a:t></a:r></a:p></a:txBody></p:sp>'

def make_slide(shapes):
    return f'<?xml version="1.0" encoding="UTF-8"?>\n<p:sld {NS} showMasterSp="1" showMasterPhAnim="1"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>{shapes}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr><p:transition spd="med" advClick="1"/></p:sld>'

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
        # Leer PPTX subido
        pptx_file = request.files.get('plantilla')
        if not pptx_file:
            return jsonify({'error': 'No se recibió la plantilla'}), 400

        datos = json.loads(request.form.get('datos', '{}'))
        empresa   = datos.get('empresa', 'Cliente')
        moneda    = datos.get('moneda', 'USD')
        licencias = datos.get('licencias', [])
        setup     = datos.get('setup', [])
        opcionales= datos.get('opcionales', [])
        notas     = datos.get('notas', 'Precios en USD. IVA no incluido. Vigencia 30 días.')
        fecha     = datos.get('fecha', '')

        def fmt(n):
            sym = '$' if moneda in ('USD','MXN') else '€'
            return f"{sym}{float(n):,.2f} {moneda}"

        # Licencias
        lic_h = ["Producto / Licencia", "Volumen", f"Precio unit. ({moneda})", f"Total ({moneda})"]
        lic_w = [2500000, 1900000, 1900000, 1700000]
        lic_rows = []
        for r in licencias:
            total = (float(r.get('vol') or 0)) * (float(r.get('precio') or 0))
            lic_rows.append([r.get('nombre',''), str(r.get('vol','')),
                fmt(r['precio']) if r.get('precio') else '',
                fmt(total) if r.get('nombre') else ''])

        # Setup
        setup_h = ["Servicio", "Descripción", "Tipo de pago", f"Precio ({moneda})"]
        setup_w = [2000000, 3100000, 1500000, 1629600]
        setup_rows = [[r.get('servicio',''), r.get('descripcion',''), r.get('tipo','Único'),
            fmt(r['precio']) if r.get('precio') else ''] for r in setup]

        # Opcionales
        opc_h = ["Servicio", "Descripción", "Tipo de pago", f"Precio ({moneda})"]
        opc_w = [2000000, 3100000, 1500000, 1629600]
        opc_rows = [[r.get('servicio',''), r.get('descripcion',''), r.get('tipo','Anual'),
            fmt(r['precio']) if r.get('precio') else ''] for r in opcionales]

        # Resumen
        lic_total   = sum((float(r.get('vol') or 0))*(float(r.get('precio') or 0)) for r in licencias)
        setup_total = sum(float(r.get('precio') or 0) for r in setup)
        opc_total   = sum(float(r.get('precio') or 0) for r in opcionales)
        res_h = ["Escenario", "Descripción", f"Inversión Año 1 ({moneda})"]
        res_w = [2300000, 3800000, 2129600]
        res_rows = [
            ["Escenario 1 — Base", "Licencias únicamente", fmt(lic_total)],
            ["Escenario 2 — Completo", "Licencias + Setup & Onboarding", fmt(lic_total+setup_total)],
            ["Escenario 3 — Premium", "Licencias + Setup + Servicios Opcionales", fmt(lic_total+setup_total+opc_total)],
        ]

        notas_lines = [l.strip() for l in notas.split('\n') if l.strip()]
        notas_shapes = label_sp(400, "Condiciones y notas adicionales:", 2895600)
        for i, l in enumerate(notas_lines):
            notas_shapes += note_sp(401+i, "• "+l, 3200400+i*304800)

        patches = {
            "ppt/slides/slide19.xml": make_slide(title_sp("Licencias")+div_sp()+make_table(457200,990600,8229600,2743200,lic_w,lic_h,lic_rows)+note_sp(400,"* Tarifa anual por usuario activo. Sujeto a contrato.",3810000)),
            "ppt/slides/slide20.xml": make_slide(title_sp("Setup & Onboarding")+div_sp()+make_table(457200,990600,8229600,2743200,setup_w,setup_h,setup_rows)+note_sp(400,"* Pago único al inicio. No incluye desarrollos adicionales.",3810000)),
            "ppt/slides/slide21.xml": make_slide(title_sp("Servicios Opcionales")+div_sp()+make_table(457200,990600,8229600,2743200,opc_w,opc_h,opc_rows)+note_sp(400,"* No incluidos en propuesta base. Sujetos a disponibilidad.",3810000)),
            "ppt/slides/slide22.xml": make_slide(title_sp("Resumen")+div_sp()+make_table(457200,990600,8229600,1800000,res_w,res_h,res_rows)+notas_shapes),
            "ppt/slides/_rels/slide19.xml.rels": RELS,
            "ppt/slides/_rels/slide20.xml.rels": RELS,
            "ppt/slides/_rels/slide21.xml.rels": RELS,
            "ppt/slides/_rels/slide22.xml.rels": RELS,
        }

        pptx_bytes = pptx_file.read()
        out = io.BytesIO()
        with zipfile.ZipFile(io.BytesIO(pptx_bytes), 'r') as zin:
            with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    if item.filename in patches:
                        zout.writestr(item, patches[item.filename])
                    else:
                        zout.writestr(item, zin.read(item.filename))

        out.seek(0)
        fecha_str = (fecha or '').replace('-','')
        filename = f"Propuesta_Emburse_{empresa.replace(' ','_')}_{fecha_str}.pptx"

        resp = send_file(out, mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation',
                        as_attachment=True, download_name=filename)
        resp.headers['Access-Control-Allow-Origin'] = '*'
        return resp

    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    import os
    port = int(os.environ.get('PORT', 8080))
    app.run(host='0.0.0.0', port=port)
