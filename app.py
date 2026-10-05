from fastapi import FastAPI, Form, UploadFile, File
from fastapi.responses import HTMLResponse, FileResponse
from fastapi.staticfiles import StaticFiles
from pathlib import Path
from typing import List
from uuid import uuid4
from docx import Document
from docx.shared import Inches
import shutil, os

app = FastAPI(title='APP DE TASACIONES RD')
BASE = Path(__file__).resolve().parent
DATA = BASE / 'data'
UPLOADS = DATA / 'uploads'
REPORTS = DATA / 'reports'
for p in (UPLOADS, REPORTS): p.mkdir(parents=True, exist_ok=True)

STYLE='''<style>
body{font-family:Arial,sans-serif;background:#f4f6f8;margin:0;color:#172033}.top{background:#172033;color:white;padding:18px 6%;}.top b{font-size:22px}.wrap{max-width:1050px;margin:28px auto;padding:0 18px}.card{background:white;border-radius:14px;padding:24px;box-shadow:0 2px 12px #0001;margin-bottom:20px}h1{margin:0 0 6px}h2{font-size:18px;margin-top:0}.grid{display:grid;grid-template-columns:repeat(2,1fr);gap:14px}.full{grid-column:1/-1}label{font-size:13px;font-weight:bold;display:block;margin-bottom:5px}input,select,textarea{box-sizing:border-box;width:100%;padding:11px;border:1px solid #cbd2dc;border-radius:8px}textarea{min-height:80px}.btn{background:#f2b705;border:0;padding:12px 20px;border-radius:8px;font-weight:bold;cursor:pointer}.note{background:#fff8df;padding:12px;border-radius:8px;font-size:13px}.ok{background:#eaf8ef;border-left:4px solid #2b9b55;padding:14px;border-radius:8px}.steps{display:flex;gap:8px;flex-wrap:wrap;margin:16px 0}.step{background:#e8edf3;padding:7px 11px;border-radius:18px;font-size:12px}@media(max-width:700px){.grid{grid-template-columns:1fr}}
</style>'''

def page(body):
    return HTMLResponse(f'<!doctype html><html lang="es"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>APP DE TASACIONES</title>{STYLE}</head><body><div class="top"><b>APP DE TASACIONES</b><div>República Dominicana · Prototipo de prueba</div></div>{body}</body></html>')

@app.get('/', response_class=HTMLResponse)
def home():
    return page('''<div class="wrap"><div class="card"><h1>Nueva tasación</h1><p>Complete los datos básicos y cargue los documentos para probar el flujo del expediente.</p><div class="steps"><span class="step">1 Datos</span><span class="step">2 Títulos</span><span class="step">3 Fotos</span><span class="step">4 Mejoras</span><span class="step">5 QA</span><span class="step">6 Word</span></div></div>
<form class="card" action="/generar" method="post" enctype="multipart/form-data"><h2>Datos del expediente</h2><div class="grid">
<div><label>Expediente</label><input name="expediente" placeholder="EXP-2026-001" required></div><div><label>Banco</label><select name="banco"><option>Banco Popular Dominicano</option><option>Popular Bank, LTD</option><option>Otro / Sin seleccionar</option></select></div>
<div><label>Propietario</label><input name="propietario" required></div><div><label>Fecha de tasación</label><input type="date" name="fecha"></div>
<div class="full"><label>Dirección</label><input name="direccion"></div><div><label>Sector</label><input name="sector"></div><div><label>Municipio / Provincia</label><input name="municipio"></div>
<div><label>Matrícula</label><input name="matricula"></div><div><label>Área total (m²)</label><input name="area"></div>
<div class="full"><label>Designación posicional (12 dígitos)</label><input name="designacion" maxlength="12"></div>
<div class="full"><label>Descripción de mejoras</label><textarea name="mejoras" placeholder="Terminaciones, distribución, estado físico..."></textarea></div>
<div><label>Plantilla Word maestra (.docx)</label><input type="file" name="plantilla" accept=".docx"></div><div><label>Foto de fachada (selección manual)</label><input type="file" name="fachada" accept="image/*"></div>
<div class="full"><label>Títulos / documentos legales</label><input type="file" name="titulos" multiple></div><div class="full"><label>Fotos adicionales</label><input type="file" name="fotos" accept="image/*" multiple></div>
</div><br><div class="note">Esta versión es para probar el flujo. La plantilla Word cargada se conserva como base; el motor definitivo de fidelidad visual seguirá ajustándose con pruebas reales.</div><br><button class="btn" type="submit">Generar tasación de prueba</button></form></div>''')

def save_upload(up: UploadFile|None, folder: Path):
    if not up or not up.filename: return None
    folder.mkdir(parents=True, exist_ok=True)
    path=folder/(str(uuid4())+'_'+Path(up.filename).name)
    with path.open('wb') as f: shutil.copyfileobj(up.file,f)
    return path

@app.post('/generar', response_class=HTMLResponse)
def generar(expediente:str=Form(...), banco:str=Form(...), propietario:str=Form(...), fecha:str=Form(''), direccion:str=Form(''), sector:str=Form(''), municipio:str=Form(''), matricula:str=Form(''), area:str=Form(''), designacion:str=Form(''), mejoras:str=Form(''), plantilla:UploadFile=File(None), fachada:UploadFile=File(None), titulos:List[UploadFile]=File([]), fotos:List[UploadFile]=File([])):
    eid=''.join(c for c in expediente if c.isalnum() or c in '-_') or str(uuid4())[:8]
    folder=UPLOADS/eid; folder.mkdir(parents=True,exist_ok=True)
    template_path=save_upload(plantilla,folder)
    fachada_path=save_upload(fachada,folder)
    title_paths=[p for u in titulos if (p:=save_upload(u,folder))]
    photo_paths=[p for u in fotos if (p:=save_upload(u,folder))]
    try:
        doc=Document(str(template_path)) if template_path else Document()
    except Exception:
        doc=Document()
    doc.add_heading('INFORME DE TASACIÓN - DATOS DE PRUEBA', level=1)
    fields=[('Expediente',expediente),('Banco',banco),('Propietario',propietario),('Fecha',fecha),('Dirección',direccion),('Sector',sector),('Municipio / Provincia',municipio),('Matrícula',matricula),('Área total',area+' m²' if area else ''),('Designación posicional',designacion)]
    for k,v in fields:
        if v: doc.add_paragraph(f'{k}: {v}')
    if banco=='Banco Popular Dominicano': doc.add_paragraph('Tasación solo es válida para el Banco Popular Dominicano')
    elif banco=='Popular Bank, LTD': doc.add_paragraph('Tasación válida solo para el Popular Bank, LTD')
    if mejoras:
        doc.add_heading('Mejoras',level=2); doc.add_paragraph(mejoras)
    if fachada_path:
        doc.add_heading('Fachada seleccionada',level=2)
        try: doc.add_picture(str(fachada_path),width=Inches(5.8))
        except Exception: doc.add_paragraph('[No se pudo insertar la imagen de fachada]')
    if photo_paths:
        doc.add_heading('Fotografías',level=2)
        for p in photo_paths:
            try: doc.add_picture(str(p),width=Inches(3.0))
            except Exception: pass
    if title_paths:
        doc.add_page_break(); doc.add_heading('Documentos legales cargados',level=1)
        for p in title_paths: doc.add_paragraph(p.name)
    out=REPORTS/f'{eid}_tasacion_prueba.docx'; doc.save(out)
    checks=[('Propietario',bool(propietario)),('Dirección',bool(direccion)),('Matrícula',bool(matricula)),('Área',bool(area)),('Fachada',bool(fachada_path)),('Plantilla maestra',bool(template_path)),('Títulos',bool(title_paths))]
    qa=sum(1 for _,x in checks if x); total=len(checks)
    rows=''.join(f'<li>{"✅" if ok else "⚠️"} {name}</li>' for name,ok in checks)
    return page(f'''<div class="wrap"><div class="card"><div class="ok"><b>Expediente generado.</b><br>QA preliminar: {qa}/{total} controles completos.</div><h2 style="margin-top:20px">Revisión rápida</h2><ul>{rows}</ul><p><a class="btn" href="/descargar/{out.name}" style="display:inline-block;text-decoration:none;color:#172033">Descargar Word de prueba</a></p><p><a href="/">← Crear otra tasación</a></p></div></div>''')

@app.get('/descargar/{filename}')
def descargar(filename:str):
    safe=Path(filename).name; p=REPORTS/safe
    return FileResponse(p, filename=safe, media_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document')

@app.get('/health')
def health(): return {'status':'ok','app':'tasaciones-rd'}
