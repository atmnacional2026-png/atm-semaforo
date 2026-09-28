"""Convierte el Excel del padrón de ATMs en padron_data.json para padron.html.

Uso: python3 scripts/padron_to_json.py <Padron.xlsx> [salida.json]
Requiere openpyxl (pip install openpyxl).
"""
import collections
import datetime
import json
import re
import sys
import unicodedata

import openpyxl

# Encabezado del Excel -> clave en el JSON
COLUMNAS = {
    'Jurisdicción': 'jurisdiccion',
    'ATM N°': 'atm',
    'CUE N°': 'cue',
    'Dominio': 'dominio',
    'Chasis N°': 'chasis',
    'Patrimonio': 'patrimonio',
    'Código estado funcional': 'codigo_ef',
    'Etiqueta estado funcional': 'etiqueta_ef',
    'Especialidad formativa declarada': 'especialidad_decl',
    'Especialidad real vigente': 'especialidad_real',
    '¿Coincide especialidad declarada/real?': 'coincide_esp',
    'Estado de uso': 'estado_uso',
    'Apta para circular': 'apta_circular',
    'Apta para dictar cursos': 'apta_cursos',
    'Ploteo MCH': 'ploteo',
    'Estado del convenio': 'convenio',
    'Fecha de vencimiento del convenio': 'convenio_venc',
    'Localidad actual': 'localidad',
    'Provincia de ubicación actual': 'provincia_ubic',
    'Institución/predio anfitrión': 'institucion',
    'Tipo de emplazamiento': 'emplazamiento',
    'Modalidad de tenencia en ubicación': 'tenencia',
    'Fecha de llegada a ubicación actual': 'fecha_llegada',
    'Localidad próxima planificada': 'localidad_prox',
    'Provincia próxima planificada': 'provincia_prox',
    'Fecha estimada de traslado': 'fecha_traslado',
    'Fecha inicio vigencia póliza': 'poliza_inicio',
    'Fecha fin vigencia póliza': 'poliza_fin',
    'Estado de póliza': 'poliza_estado',
    'póliza': 'poliza',
    'Cantidad de cursos activos': 'cursos_activos',
    'Matrícula total registrada': 'matricula',
    'Cantidad de cursos finalizados': 'cursos_finalizados',
    'Matrícula cursos finaliz': 'matricula_finalizados',
    'Cursos planificados': 'cursos_planificados',
    'Preinscriptos cursos planificados': 'preinscriptos',
    'Detalle de cursos': 'detalle_cursos',
    'Fecha de finalización último curso': 'fecha_ultimo_curso',
    'Modalidad de traslado': 'modalidad_traslado',
    'Acta Complementaria vinculada (N°)': 'acta',
    'Fecha de Acta Complementaria': 'acta_fecha',
    'Vínculo con Plan de Mejora': 'plan_mejora',
    'N° expediente Plan de Mejora': 'plan_mejora_exp',
    'Fecha de último relevamiento de la ficha': 'fecha_relevamiento',
    'Responsable jurisdiccional / referente a cargo': 'responsable',
    'Observaciones Estado Funcional': 'obs_ef',
    'Observación de especialidad': 'obs_esp',
    'Observaciones Generales': 'obs_gen',
}
NUMERICAS = {'cursos_activos', 'matricula', 'cursos_finalizados',
             'matricula_finalizados', 'cursos_planificados', 'preinscriptos'}
# Unifica variantes de escritura en campos categóricos
SINONIMOS = {
    'SI': 'Sí', 'Si': 'Sí', 'si': 'Sí', 'NO': 'No', 'no': 'No',
    'activa': 'Activa', 'inactiva': 'Inactiva',
    'Sin datos': 'Sin datos disponibles', 'Temporaria': 'Temporario',
    'No tiene': 'No tiene plan de mejora', '-': '',
    'RIO NEGRO': 'RÍO NEGRO',
}
CATEGORICAS = {'patrimonio', 'etiqueta_ef', 'especialidad_decl', 'especialidad_real',
               'coincide_esp', 'estado_uso', 'apta_circular', 'apta_cursos', 'ploteo',
               'convenio', 'tenencia', 'poliza_estado', 'modalidad_traslado',
               'plan_mejora', 'provincia_ubic', 'jurisdiccion'}


def sin_acentos(s):
    return unicodedata.normalize('NFKD', s).encode('ascii', 'ignore').decode().lower()


def unificar_especialidades(atms):
    # Misma especialidad escrita con distinta capitalización/acentos -> la forma más usada
    for clave in ('especialidad_decl', 'especialidad_real'):
        formas = collections.Counter(a[clave] for a in atms if a[clave])
        canon = {}
        for forma, _ in formas.most_common():
            canon.setdefault(sin_acentos(forma), forma)
        for a in atms:
            if a[clave]:
                a[clave] = canon[sin_acentos(a[clave])]


def limpiar(clave, v):
    if isinstance(v, (datetime.datetime, datetime.date)):
        return v.strftime('%Y-%m-%d')
    if v is None:
        v = ''
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    if clave in NUMERICAS:
        return v if isinstance(v, int) else (int(v) if str(v).strip().isdigit() else None)
    s = re.sub(r'[ \t​]+', ' ', str(v)).strip()
    if clave in CATEGORICAS:
        if clave in ('provincia_ubic', 'jurisdiccion'):
            s = s.upper()
        s = SINONIMOS.get(s, s)
    return s


def main(entrada, salida):
    wb = openpyxl.load_workbook(entrada, data_only=True, read_only=True)
    ws = wb.worksheets[0]
    filas = list(ws.iter_rows(values_only=True))
    fila_enc = next(i for i, f in enumerate(filas) if f and f[0] == 'Jurisdicción')
    enc = [str(h).strip() if h else None for h in filas[fila_enc]]
    atms = []
    for f in filas[fila_enc + 1:]:
        if not f or all(v in (None, '') for v in f):
            continue
        reg = {}
        for h, v in zip(enc, f):
            if h in COLUMNAS:
                reg[COLUMNAS[h]] = limpiar(COLUMNAS[h], v)
        if reg.get('atm') in ('', None):
            continue
        atms.append(reg)
    unificar_especialidades(atms)
    meta = {'hoja': ws.title, 'total': len(atms),
            'generado': datetime.date.today().isoformat()}
    with open(salida, 'w', encoding='utf-8') as fh:
        json.dump({'meta': meta, 'atms': atms}, fh, ensure_ascii=False)
    print(f'{len(atms)} ATMs -> {salida}')


if __name__ == '__main__':
    if len(sys.argv) < 2:
        sys.exit(__doc__)
    main(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else 'padron_data.json')
