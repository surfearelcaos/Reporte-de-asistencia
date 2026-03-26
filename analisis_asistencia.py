"""
analisis_asistencia.py
======================
Análisis de asistencia mensual — Austral Pack S.A.

Uso:
    python analisis_asistencia.py <informe_general.xlsx> <informe_semanal.xlsx>

Ejemplo:
    python analisis_asistencia.py informe-asistencia-general.xlsx informe-asistencia-semanal.xlsx

Requisitos:
    pip install pandas openpyxl

Salida:
    analisis_asistencia_<MES>.xlsx  (en la misma carpeta donde se ejecuta)
"""

import sys
import os
import numpy as np
import pandas as pd
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference

# ===========================================================
# PARÁMETROS (ajustables)
# ===========================================================
PESO_FRECUENCIA   = 0.5          # peso frecuencia en índice criticidad
PESO_DURACION     = 0.5          # peso duración en índice criticidad
UMBRAL_ATRASO_MAX = 480          # minutos: atrasos mayores se consideran marcación anómala

# Colores corporativos Austral Pack
COLOR_OSC  = "9e050d"
COLOR_MED  = "e0151e"
COLOR_CLAR = "fde8e8"
GRIS_CLAR  = "F5F5F5"
VERDE      = "E2EFDA"
ROJO_ALERT = "FCE4D6"
AMARILLO   = "FFF2CC"

# ===========================================================
# HELPERS DE ESTILO
# ===========================================================
def _border():
    s = Side(style='thin', color='CCCCCC')
    return Border(left=s, right=s, top=s, bottom=s)

def hdr(ws, row, col, val, bg=None, fg="FFFFFF", bold=True, sz=10, wrap=False, halign='center'):
    bg = bg or COLOR_OSC
    c = ws.cell(row=row, column=col, value=val)
    c.font = Font(name='Arial', bold=bold, color=fg, size=sz)
    c.fill = PatternFill('solid', start_color=bg)
    c.alignment = Alignment(horizontal=halign, vertical='center', wrap_text=wrap)
    return c

def dat(ws, row, col, val, bg=None, bold=False, fmt=None, halign='center'):
    c = ws.cell(row=row, column=col, value=val)
    c.font = Font(name='Arial', bold=bold, size=9)
    if bg:
        c.fill = PatternFill('solid', start_color=bg)
    if fmt:
        c.number_format = fmt
    c.alignment = Alignment(horizontal=halign, vertical='center')
    return c

def borders(ws, r1, r2, c1, c2):
    for row in ws.iter_rows(min_row=r1, max_row=r2, min_col=c1, max_col=c2):
        for cell in row:
            cell.border = _border()

def widths(ws, d):
    for col, w in d.items():
        ws.column_dimensions[col].width = w

# ===========================================================
# 1. CARGA DE DATOS
# ===========================================================
def _rut(val):
    """Normaliza RUT a string sin puntos ni guión, mayúsculas."""
    return str(val).strip().replace('.', '').replace('-', '').upper()

def cargar_general(path):
    df = pd.read_excel(path, header=5, engine='openpyxl')
    df = df.dropna(subset=['Rut'])
    df['rut'] = df['Rut'].apply(_rut)
    return df

def cargar_semanal(path):
    wb = openpyxl.load_workbook(path, read_only=True)
    ws = wb.active
    all_rows = list(ws.iter_rows(values_only=True))
    wb.close()

    DIAS = {'Lu', 'Ma', 'Mi', 'Ju', 'Vi', 'Sá', 'Sa', 'Do'}
    emp, rows = {}, []

    for i, row in enumerate(all_rows):
        if row[0] == 'Rut empleado:':
            emp = {
                'rut':    _rut(row[2]) if row[2] else None,
                'nombre': all_rows[i][9] if len(all_rows[i]) > 9 else None,
            }
        if row[0] and isinstance(row[0], str) and row[0][:2] in DIAS:
            rows.append({
                'rut':           emp.get('rut'),
                'nombre':        emp.get('nombre'),
                'fecha_str':     row[0],
                'dia_abrev':     row[0][:2],
                'entrada_turno': row[1],
                'salida_turno':  row[2],
                'entrada_real':  row[3],
                'salida_real':   row[6],
                'horas_asign':   row[7],
                'horas_asist':   row[8],
                'colacion_h':    row[9],
                'justificados':  row[16],
                'turno_nombre':  row[18],
            })

    return pd.DataFrame(rows)

def cargar_nomina(path):
    df = pd.read_excel(path, engine='openpyxl')
    df.columns = df.columns.str.strip()
    df = df.dropna(subset=['Rut'])
    df['rut']    = df['Rut'].apply(_rut)
    df['nombre'] = df['Nombre'].astype(str).str.strip()
    df['cargo']  = df['Cargo'].astype(str).str.strip()
    df['area']   = df['Área'].astype(str).str.strip()
    return df[['rut', 'nombre', 'cargo', 'area']]

def hhmm_a_min(val):
    """Convierte 'HH:MM' a minutos desde medianoche. Retorna None si inválido."""
    if val is None or (isinstance(val, float) and np.isnan(val)):
        return None
    s = str(val).strip()
    if ':' not in s:
        return None
    try:
        h, m = s.split(':')[:2]
        return int(h) * 60 + int(m)
    except Exception:
        return None

# ===========================================================
# 2. CONSTRUCCIÓN DEL DATAFRAME PRINCIPAL
# ===========================================================
def construir_df(path_general, path_semanal, path_nomina):
    df_gen = cargar_general(path_general)
    df     = cargar_semanal(path_semanal)
    df_nom = cargar_nomina(path_nomina)

    # Mapas desde la nómina (fuente de verdad)
    nom_map   = df_nom.set_index('rut')['nombre'].to_dict()
    area_map  = df_nom.set_index('rut')['area'].to_dict()
    cargo_map = df_nom.set_index('rut')['cargo'].to_dict()

    df['nombre'] = df['rut'].map(nom_map)
    df['area']   = df['rut'].map(area_map)
    df['cargo']  = df['rut'].map(cargo_map)

    # Convertir horas a minutos
    df['et_min'] = df['entrada_turno'].apply(hhmm_a_min)
    df['st_min'] = df['salida_turno'].apply(hhmm_a_min)
    df['er_min'] = df['entrada_real'].apply(hhmm_a_min)
    df['sr_min'] = df['salida_real'].apply(hhmm_a_min)
    df['colacion_min'] = df['colacion_h'].fillna(0) * 60

    # Flag permiso: justificados == 1 (semanal) cruzado con cols M/N/O/Q (general)
    df['con_permiso'] = df['justificados'].fillna(0) > 0

    return df, df_gen, df_nom

# ===========================================================
# 3. ANÁLISIS
# ===========================================================
def analizar_permisos(df_gen, df_nom):
    nom_map  = df_nom.set_index('rut')['nombre'].to_dict()
    area_map = df_nom.set_index('rut')['area'].to_dict()

    perms = df_gen.groupby('rut').agg(
        perm_goce     = ('Permisos con goce de Sueldo','sum'),
        vacaciones    = ('Vacaciones y días Admin.',   'sum'),
        perm_sin_goce = ('Permiso sin goce de Sueldo', 'sum'),
        lic_medica    = ('Licencia Médica',            'sum'),
    ).reset_index()
    perms['nombre'] = perms['rut'].map(nom_map)
    perms['area']   = perms['rut'].map(area_map)
    perms['total']  = perms[['perm_goce','vacaciones','perm_sin_goce','lic_medica']].sum(axis=1)
    return perms[perms['total'] > 0].sort_values(['area','total'], ascending=[True, False])

def analizar_atrasos(df):
    """Atrasos de entrada. Excluye permisos y marcaciones anómalas (er >= st)."""
    mask = (
        ~df['con_permiso'] &
        df['et_min'].notna() &
        df['er_min'].notna() &
        ~((df['er_min'].notna()) & (df['st_min'].notna()) & (df['er_min'] >= df['st_min']))
    )
    da = df[mask].copy()
    da['atraso_min'] = da['er_min'] - da['et_min']
    da['tuvo_atraso'] = da['atraso_min'] > 0

    # Filtrar anomalías por umbral
    da = da[~(da['tuvo_atraso'] & (da['atraso_min'] > UMBRAL_ATRASO_MAX))]

    personas = (
        da.groupby(['rut','nombre','cargo','area'])['tuvo_atraso']
        .any().reset_index()
        .query('tuvo_atraso == True')
        .sort_values('area')
    )

    por_area = (
        da[da['tuvo_atraso']]
        .groupby('area')
        .agg(
            n_personas       = ('rut',        'nunique'),
            n_registros      = ('atraso_min', 'count'),
            prom_atraso_min  = ('atraso_min', lambda x: round(x.mean(), 1)),
            total_min_atraso = ('atraso_min', 'sum'),
        )
        .reset_index()
        .sort_values('prom_atraso_min', ascending=False)
    )

    max_por_area = (
        da[da['tuvo_atraso']]
        .loc[da[da['tuvo_atraso']].groupby('area')['atraso_min'].idxmax()]
        [['area','nombre','cargo','fecha_str','entrada_real','atraso_min']]
        .round({'atraso_min': 1})
        .set_index('area')
    )

    return da, personas, por_area, max_por_area

def analizar_salidas(df):
    mask = (
        ~df['con_permiso'] &
        df['st_min'].notna() &
        df['sr_min'].notna()
    )
    ds = df[mask].copy()
    ds['extra_min'] = ds['sr_min'] - ds['st_min']
    ds['se_quedo'] = ds['extra_min'] > 0

    por_area = (
        ds[ds['se_quedo']]
        .groupby('area')
        .agg(
            n_personas_extra = ('rut',       'nunique'),
            prom_extra_min   = ('extra_min', lambda x: round(x.mean(), 1)),
            n_registros      = ('extra_min', 'count'),
        )
        .reset_index()
        .sort_values('prom_extra_min', ascending=False)
    )
    return por_area

def analizar_horas(df):
    mask = (
        ~df['con_permiso'] &
        df['er_min'].notna() &
        df['sr_min'].notna()
    )
    dh = df[mask].copy()
    dh['horas_brutas']  = (dh['sr_min'] - dh['er_min']) / 60
    dh['horas_asist_m'] = (dh['horas_brutas'] - dh['colacion_min'] / 60).clip(lower=0)
    dh['diff_horas']    = dh['horas_asist_m'] - dh['horas_asign'].fillna(0)

    por_persona = (
        dh.groupby(['area','rut','nombre','cargo'])
        .agg(
            horas_asignadas = ('horas_asign',  'sum'),
            horas_asistidas = ('horas_asist_m','sum'),
            diff_horas      = ('diff_horas',   'sum'),
        )
        .reset_index()
        .round(2)
        .sort_values(['area','diff_horas'])
    )

    por_area = (
        dh.groupby('area')
        .agg(
            horas_asign_total = ('horas_asign',  'sum'),
            horas_asist_total = ('horas_asist_m','sum'),
            diff_total        = ('diff_horas',   'sum'),
            n_personas        = ('rut',          'nunique'),
        )
        .reset_index()
    )
    por_area['diff_prom_pp'] = (por_area['diff_total'] / por_area['n_personas']).round(2)
    por_area = por_area.round(2).sort_values('diff_total')

    dh['deficit_min'] = (dh['horas_asign'].fillna(0) - dh['horas_asist_m']).clip(lower=0) * 60
    perdida = (
        dh.groupby('area')
        .agg(
            minutos_perdidos      = ('deficit_min', 'sum'),
            registros_con_perdida = ('deficit_min', lambda x: (x > 0).sum()),
            personas_afectadas    = ('rut',         lambda x: x[dh.loc[x.index,'deficit_min'] > 0].nunique()),
        )
        .reset_index()
        .round(1)
        .sort_values('minutos_perdidos', ascending=False)
    )

    return por_persona, por_area, perdida

def calcular_criticidad(df_atr):
    crit = (
        df_atr.groupby('area')
        .agg(
            n_personas           = ('rut',        'nunique'),
            n_registros_atraso   = ('tuvo_atraso','sum'),
            minutos_atraso_total = ('atraso_min', lambda x: x[x > 0].sum()),
        )
        .reset_index()
    )
    crit['freq_pp']   = crit['n_registros_atraso'] / crit['n_personas'].clip(lower=1)
    crit['dur_ppmin'] = crit['minutos_atraso_total'] / crit['n_personas'].clip(lower=1)
    max_f = crit['freq_pp'].max()
    max_d = crit['dur_ppmin'].max()
    crit['freq_norm'] = crit['freq_pp']   / max_f if max_f > 0 else 0
    crit['dur_norm']  = crit['dur_ppmin'] / max_d if max_d > 0 else 0
    crit['indice_0_10'] = (
        10 * (PESO_FRECUENCIA * crit['freq_norm'] + PESO_DURACION * crit['dur_norm'])
    ).round(2)

    def cal(x):
        if x < 2:  return 'Muy baja'
        elif x < 4: return 'Baja'
        elif x < 6: return 'Media'
        elif x < 8: return 'Alta'
        else:       return 'Muy alta'

    crit['calificacion'] = crit['indice_0_10'].apply(cal)
    crit = crit.sort_values('indice_0_10', ascending=False).reset_index(drop=True)
    crit['rank'] = crit.index + 1
    return crit

# ===========================================================
# 4. GENERACIÓN DEL EXCEL
# ===========================================================
CAL_COLORS = {
    'Muy alta': "FCE4D6",
    'Alta':     "FCAE7C",
    'Media':    AMARILLO,
    'Baja':     VERDE,
    'Muy baja': "E8F5E9",
}

def _desc(ws, texto, n_cols=10):
    ws.merge_cells(f'A2:{get_column_letter(n_cols)}2')
    c = ws['A2']
    c.value = texto
    c.font = Font(name='Arial', size=9, italic=True, color="444444")
    c.fill = PatternFill('solid', start_color="F5F5F5")
    c.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
    ws.row_dimensions[2].height = 30

def _titulo(ws, texto, n_cols=10, h=28):
    ws.merge_cells(f'A1:{get_column_letter(n_cols)}1')
    c = ws['A1']
    c.value = texto
    c.font  = Font(name='Arial', bold=True, size=12, color="FFFFFF")
    c.fill  = PatternFill('solid', start_color=COLOR_OSC)
    c.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = h
    ws.sheet_view.showGridLines = False

def sheet_resumen(wb, crit, n_anomalias):
    ws = wb.create_sheet("Resumen Ejecutivo")
    _titulo(ws, "ANÁLISIS DE ASISTENCIA · AUSTRAL PACK S.A.", n_cols=10, h=35)

    ws.merge_cells('A2:J2')
    nota = ws['A2']
    nota.value = (
        f"Excluidas {n_anomalias} marcaciones anómalas (entrada_real ≥ salida_turno o atraso > {UMBRAL_ATRASO_MAX} min)"
    )
    nota.font = Font(name='Arial', size=8, italic=True, color="888888")
    nota.alignment = Alignment(horizontal='center', vertical='center')

    ws.merge_cells('A3:J3')
    disc = ws['A3']
    disc.value = (
        "El índice se calcula según la frecuencia y duración de atrasos por persona, "
        "destacando las áreas con mayor impacto real en tiempo perdido."
    )
    disc.font = Font(name='Arial', size=9, italic=True, color="444444")
    disc.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
    ws.row_dimensions[3].height = 36

    row = 5
    ws.merge_cells(f'A{row}:J{row}')
    hdr(ws, row, 1, "◆  INDICADORES GLOBALES", bg=COLOR_MED, sz=10)
    ws.row_dimensions[row].height = 18
    row += 1

    kpis = [
        ("Área más crítica (atrasos)", f"{crit.iloc[0]['area']}  (índice {crit.iloc[0]['indice_0_10']} / 10)"),
    ]
    hdr(ws, row, 1, "Indicador", bg=COLOR_CLAR, fg="000000", sz=9, halign='left')
    ws.merge_cells(f'B{row}:J{row}')
    hdr(ws, row, 2, "Valor", bg=COLOR_CLAR, fg="000000", sz=9)
    row += 1
    for k, v in kpis:
        dat(ws, row, 1, k, halign='left')
        ws.merge_cells(f'B{row}:J{row}')
        dat(ws, row, 2, v, halign='left', bold=True)
        ws.row_dimensions[row].height = 16
        row += 1

    row += 1
    ws.merge_cells(f'A{row}:J{row}')
    hdr(ws, row, 1, "◆  ÍNDICE DE CRITICIDAD POR ÁREA (Atrasos — 0 a 10)", bg=COLOR_MED, sz=10)
    ws.row_dimensions[row].height = 18
    row += 1

    hdrs = ["Rank","Área","N° Personas","Reg. c/Atraso","Min Atraso","Frec/Persona","Min/Persona","Índice","Calificación"]
    for ci, h in enumerate(hdrs, 1):
        hdr(ws, row, ci, h, bg=COLOR_CLAR, fg="000000", sz=9, wrap=True)
    ws.row_dimensions[row].height = 28
    row += 1

    start = row
    for _, rw in crit.iterrows():
        dat(ws, row, 1, int(rw['rank']))
        dat(ws, row, 2, rw['area'], halign='left')
        dat(ws, row, 3, int(rw['n_personas']))
        dat(ws, row, 4, int(rw['n_registros_atraso']))
        dat(ws, row, 5, round(rw['minutos_atraso_total'], 1), fmt='#,##0.0')
        dat(ws, row, 6, round(rw['freq_pp'], 2),   fmt='0.00')
        dat(ws, row, 7, round(rw['dur_ppmin'], 1),  fmt='#,##0.0')
        dat(ws, row, 8, rw['indice_0_10'],           fmt='0.00', bold=True)
        c = dat(ws, row, 9, rw['calificacion'], bold=True)
        bg = CAL_COLORS.get(rw['calificacion'])
        if bg:
            c.fill = PatternFill('solid', start_color=bg)
        ws.row_dimensions[row].height = 15
        row += 1

    borders(ws, start - 1, row - 1, 1, 9)
    widths(ws, {'A':6,'B':22,'C':11,'D':13,'E':14,'F':12,'G':12,'H':10,'I':13,'J':5})

def sheet_permisos(wb, perms):
    ws = wb.create_sheet("Permisos")
    _titulo(ws, "PERMISOS Y VACACIONES", n_cols=7)
    _desc(ws, "Detalle de días de permisos, vacaciones y licencias médicas utilizados por cada persona durante el período. Los totales iguales o mayores a 5 días se destacan en rojo.", n_cols=7)

    hdrs = ["Área","Nombre","Perm. con Goce","Vacac. y Días Admin.","Perm. sin Goce","Lic. Médica","TOTAL"]
    for ci, h in enumerate(hdrs, 1):
        hdr(ws, 3, ci, h, bg=COLOR_MED, sz=9, wrap=True)
    ws.row_dimensions[3].height = 28

    for ri, (_, rw) in enumerate(perms.iterrows(), 4):
        bg = GRIS_CLAR if ri % 2 == 0 else None
        dat(ws, ri, 1, rw['area'],   bg=bg, halign='left')
        dat(ws, ri, 2, rw['nombre'], bg=bg, halign='left')
        dat(ws, ri, 3, int(rw['perm_goce']),     bg=bg)
        dat(ws, ri, 4, int(rw['vacaciones']),     bg=bg)
        dat(ws, ri, 5, int(rw['perm_sin_goce']),  bg=bg)
        dat(ws, ri, 6, int(rw['lic_medica']),     bg=bg)
        c7 = dat(ws, ri, 7, int(rw['total']), bg=bg, bold=True)
        if rw['total'] >= 5:
            c7.fill = PatternFill('solid', start_color=ROJO_ALERT)
        ws.row_dimensions[ri].height = 15

    last = 3 + len(perms)
    tot  = last + 1
    ws.merge_cells(f'A{tot}:B{tot}')
    hdr(ws, tot, 1, "TOTAL", bg=COLOR_OSC, sz=9)
    for ci in range(3, 8):
        cl = get_column_letter(ci)
        c  = ws.cell(row=tot, column=ci, value=f'=SUM({cl}4:{cl}{last})')
        c.font      = Font(name='Arial', bold=True, size=9, color="FFFFFF")
        c.fill      = PatternFill('solid', start_color=COLOR_OSC)
        c.alignment = Alignment(horizontal='center', vertical='center')
    borders(ws, 3, tot, 1, 7)
    widths(ws, {'A':18,'B':34,'C':14,'D':20,'E':16,'F':14,'G':10})

def sheet_atrasos(wb, por_area, max_por_area, personas):
    ws = wb.create_sheet("Atrasos")
    _titulo(ws, "ANÁLISIS DE ATRASOS  (excluye permisos y marcaciones anómalas)", n_cols=8)
    _desc(ws, "Registra los atrasos de entrada por área y persona. Se excluyen días con permiso autorizado y marcaciones anómalas. Promedio superior a 20 min se marca en rojo; entre 10 y 20 min, en amarillo.", n_cols=8)

    ws.merge_cells('A3:H3')
    hdr(ws, 3, 1, "Resumen por Área", bg=COLOR_MED, sz=10)

    hdrs = ["Área","N° Personas","Reg. c/Atraso","Prom. Atraso (min)","Total Min","Máx (min)","Persona Máx","Fecha"]
    for ci, h in enumerate(hdrs, 1):
        hdr(ws, 4, ci, h, bg=COLOR_CLAR, fg="000000", sz=9, wrap=True)
    ws.row_dimensions[4].height = 28

    for ri, (_, rw) in enumerate(por_area.iterrows(), 5):
        bg   = GRIS_CLAR if ri % 2 == 0 else None
        area = rw['area']
        dat(ws, ri, 1, area,                    bg=bg, halign='left')
        dat(ws, ri, 2, int(rw['n_personas']),   bg=bg)
        dat(ws, ri, 3, int(rw['n_registros']),  bg=bg)
        c4 = dat(ws, ri, 4, rw['prom_atraso_min'], bg=bg, fmt='0.0')
        if rw['prom_atraso_min'] > 20:
            c4.fill = PatternFill('solid', start_color=ROJO_ALERT)
        elif rw['prom_atraso_min'] > 10:
            c4.fill = PatternFill('solid', start_color=AMARILLO)
        dat(ws, ri, 5, round(rw['total_min_atraso'], 1), bg=bg, fmt='#,##0.0')
        mx = max_por_area.loc[area] if area in max_por_area.index else None
        dat(ws, ri, 6, round(mx['atraso_min'], 1) if mx is not None else '-', fmt='0.0')
        dat(ws, ri, 7, mx['nombre']    if mx is not None else '-', halign='left')
        dat(ws, ri, 8, mx['fecha_str'] if mx is not None else '-')
        ws.row_dimensions[ri].height = 15

    last3 = 4 + len(por_area)
    borders(ws, 4, last3, 1, 8)

    dr = last3 + 2
    ws.merge_cells(f'A{dr}:H{dr}')
    hdr(ws, dr, 1, "Detalle — Personas con al menos 1 atraso", bg=COLOR_MED, sz=10)
    dr += 1
    for ci, h in enumerate(["Área","Nombre","Cargo"], 1):
        hdr(ws, dr, ci, h, bg=COLOR_CLAR, fg="000000", sz=9)
    dr += 1
    for _, rw in personas.iterrows():
        dat(ws, dr, 1, rw['area'],  halign='left')
        dat(ws, dr, 2, rw['nombre'], halign='left')
        dat(ws, dr, 3, rw.get('cargo', ''), halign='left')
        ws.row_dimensions[dr].height = 15
        dr += 1
    borders(ws, last3 + 3, dr - 1, 1, 3)
    widths(ws, {'A':20,'B':12,'C':14,'D':18,'E':16,'F':15,'G':30,'H':14})

def sheet_horas(wb, por_persona, por_area, perdida):
    ws = wb.create_sheet("Horas Asig vs Asist")
    _titulo(ws, "HORAS ASIGNADAS vs. ASISTIDAS  (excluye permisos · colación según turno)", n_cols=6)
    _desc(ws, "Compara las horas de turno asignadas con las horas efectivamente trabajadas, descontando colación. Diferencia positiva indica horas extra (verde); negativa, déficit (rojo).", n_cols=6)

    ws.merge_cells('A3:F3')
    hdr(ws, 3, 1, "Resumen por Área", bg=COLOR_MED, sz=10)
    hdrs = ["Área","N° Pers","Hrs Asign","Hrs Asist","Dif (hrs)","Pers Afect"]
    for ci, h in enumerate(hdrs, 1):
        hdr(ws, 4, ci, h, bg=COLOR_CLAR, fg="000000", sz=9, wrap=True)
    ws.row_dimensions[4].height = 28

    perd_idx = perdida.set_index('area')
    for ri, (_, rw) in enumerate(por_area.iterrows(), 5):
        bg   = GRIS_CLAR if ri % 2 == 0 else None
        area = rw['area']
        dat(ws, ri, 1, area,                              bg=bg, halign='left')
        dat(ws, ri, 2, int(rw['n_personas']),              bg=bg)
        dat(ws, ri, 3, round(rw['horas_asign_total'], 2),  bg=bg, fmt='0.00')
        dat(ws, ri, 4, round(rw['horas_asist_total'], 2),  bg=bg, fmt='0.00')
        c5 = dat(ws, ri, 5, round(rw['diff_total'], 2), bg=bg, fmt='0.00', bold=True)
        if rw['diff_total'] < 0:
            c5.fill = PatternFill('solid', start_color=ROJO_ALERT)
        elif rw['diff_total'] > 0:
            c5.fill = PatternFill('solid', start_color=VERDE)
        p = perd_idx.loc[area] if area in perd_idx.index else None
        dat(ws, ri, 6, int(p['personas_afectadas']) if p is not None else 0, bg=bg)
        ws.row_dimensions[ri].height = 15

    last4 = 4 + len(por_area)
    tot4  = last4 + 1
    ws.merge_cells(f'A{tot4}:B{tot4}')
    hdr(ws, tot4, 1, "TOTAL", bg=COLOR_OSC, sz=9)
    for ci, cl, fmt in [(3,'C','0.00'),(4,'D','0.00'),(5,'E','0.00'),(6,'F','#,##0')]:
        c = ws.cell(row=tot4, column=ci, value=f'=SUM({cl}5:{cl}{last4})')
        c.font = Font(name='Arial', bold=True, size=9, color="FFFFFF")
        c.fill = PatternFill('solid', start_color=COLOR_OSC)
        c.alignment = Alignment(horizontal='center', vertical='center')
        c.number_format = fmt
    borders(ws, 4, tot4, 1, 6)

    dr = tot4 + 2
    ws.merge_cells(f'A{dr}:F{dr}')
    hdr(ws, dr, 1, "Detalle por Persona", bg=COLOR_MED, sz=10)
    dr += 1
    for ci, h in enumerate(["Área","Nombre","Cargo","Hrs Asign","Hrs Asist","Dif (hrs)"], 1):
        hdr(ws, dr, ci, h, bg=COLOR_CLAR, fg="000000", sz=9)
    dr += 1
    for _, rw in por_persona.iterrows():
        dat(ws, dr, 1, rw['area'],   halign='left')
        dat(ws, dr, 2, rw['nombre'], halign='left')
        dat(ws, dr, 3, rw['cargo'],  halign='left')
        dat(ws, dr, 4, round(rw['horas_asignadas'], 2), fmt='0.00')
        dat(ws, dr, 5, round(rw['horas_asistidas'], 2), fmt='0.00')
        dc = dat(ws, dr, 6, round(rw['diff_horas'], 2), fmt='0.00', bold=True)
        if rw['diff_horas'] < 0:
            dc.fill = PatternFill('solid', start_color=ROJO_ALERT)
        elif rw['diff_horas'] > 5:
            dc.fill = PatternFill('solid', start_color=VERDE)
        ws.row_dimensions[dr].height = 15
        dr += 1
    borders(ws, tot4 + 3, dr - 1, 1, 6)
    widths(ws, {'A':20,'B':32,'C':30,'D':12,'E':12,'F':14})

def sheet_minutos_perdidos(wb, perdida):
    ws = wb.create_sheet("Minutos Perdidos")
    _titulo(ws, "MINUTOS PERDIDOS POR ÁREA  (hrs asignadas > hrs asistidas, excluye permisos)", n_cols=5)
    _desc(ws, "Muestra el déficit acumulado por área: minutos en que las horas trabajadas fueron inferiores a las del turno asignado. Superior a 200 min se marca en rojo; entre 100 y 200 min, en amarillo.", n_cols=5)

    hdrs = ["Área","Min Perdidos","Hrs Perdidas","Reg. c/Déficit","Pers. Afectadas"]
    for ci, h in enumerate(hdrs, 1):
        hdr(ws, 3, ci, h, bg=COLOR_MED, sz=9, wrap=True)
    ws.row_dimensions[3].height = 28

    for ri, (_, rw) in enumerate(perdida.iterrows(), 4):
        bg = GRIS_CLAR if ri % 2 == 0 else None
        dat(ws, ri, 1, rw['area'], bg=bg, halign='left')
        mp = rw['minutos_perdidos']
        c2 = dat(ws, ri, 2, mp, bg=bg, fmt='#,##0.0', bold=True)
        if mp > 200:
            c2.fill = PatternFill('solid', start_color=ROJO_ALERT)
        elif mp > 100:
            c2.fill = PatternFill('solid', start_color=AMARILLO)
        dat(ws, ri, 3, round(mp / 60, 2), bg=bg, fmt='0.00')
        dat(ws, ri, 4, int(rw['registros_con_perdida']), bg=bg)
        dat(ws, ri, 5, int(rw['personas_afectadas']),    bg=bg)
        ws.row_dimensions[ri].height = 15

    last5 = 3 + len(perdida)
    tot5  = last5 + 1
    hdr(ws, tot5, 1, "TOTAL", bg=COLOR_OSC, sz=9)
    for ci, cl, fmt in [(2,'B','#,##0.0'),(4,'D','#,##0'),(5,'E','#,##0')]:
        c = ws.cell(row=tot5, column=ci, value=f'=SUM({cl}4:{cl}{last5})')
        c.font = Font(name='Arial', bold=True, size=9, color="FFFFFF")
        c.fill = PatternFill('solid', start_color=COLOR_OSC)
        c.alignment = Alignment(horizontal='center', vertical='center')
        c.number_format = fmt
    c3 = ws.cell(row=tot5, column=3, value=f'=B{tot5}/60')
    c3.font = Font(name='Arial', bold=True, size=9, color="FFFFFF")
    c3.fill = PatternFill('solid', start_color=COLOR_OSC)
    c3.alignment = Alignment(horizontal='center', vertical='center')
    c3.number_format = '0.00'
    borders(ws, 3, tot5, 1, 5)
    widths(ws, {'A':22,'B':16,'C':15,'D':18,'E':18})

def sheet_dashboard(wb, crit, por_area_atr, perdida, area_hrs, perms):
    ws = wb.create_sheet("Dashboard", 0)
    ws.sheet_view.showGridLines = False

    # Título
    ws.merge_cells('A1:R1')
    c = ws['A1']
    c.value = "DASHBOARD DE ASISTENCIA · AUSTRAL PACK S.A."
    c.font  = Font(name='Arial', bold=True, size=14, color="FFFFFF")
    c.fill  = PatternFill('solid', start_color=COLOR_OSC)
    c.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 36

    # KPI boxes
    kpi_items = [
        ("Áreas analizadas",    str(len(crit))),
        ("Con atrasos",         f"{int(crit['n_personas'].sum())} pers."),
        ("Min atraso total",    f"{int(crit['minutos_atraso_total'].sum())} min"),
        ("Min perdidos total",  f"{int(perdida['minutos_perdidos'].sum())} min"),
        ("Área más crítica",    crit.iloc[0]['area']),
        ("Índice criticidad",   f"{crit.iloc[0]['indice_0_10']} / 10"),
    ]
    col_starts = [1, 3, 5, 7, 9, 11]
    for (label, val), cs in zip(kpi_items, col_starts):
        ce = cs + 1
        ws.merge_cells(start_row=3, start_column=cs, end_row=3, end_column=ce)
        ws.merge_cells(start_row=4, start_column=cs, end_row=4, end_column=ce)
        ws.merge_cells(start_row=5, start_column=cs, end_row=5, end_column=ce)
        cl = ws.cell(row=3, column=cs, value=label)
        cl.font = Font(name='Arial', size=8, color="FFFFFF")
        cl.fill = PatternFill('solid', start_color=COLOR_MED)
        cl.alignment = Alignment(horizontal='center', vertical='center')
        cv = ws.cell(row=4, column=cs, value=val)
        cv.font = Font(name='Arial', bold=True, size=13, color=COLOR_OSC)
        cv.fill = PatternFill('solid', start_color="FFFFFF")
        cv.alignment = Alignment(horizontal='center', vertical='center')
        cb = ws.cell(row=5, column=cs)
        cb.fill = PatternFill('solid', start_color=COLOR_OSC)
    ws.row_dimensions[3].height = 16
    ws.row_dimensions[4].height = 26
    ws.row_dimensions[5].height = 5

    for col in [get_column_letter(i) for i in range(1, 19)]:
        ws.column_dimensions[col].width = 12

    # ---- Datos para gráficos (en el mismo Dashboard, fila 75+) ----
    DR = 200  # fila de encabezados

    # Bloque A: criticidad — cols A,B
    ws.cell(DR, 1, 'Área'); ws.cell(DR, 2, 'Índice')
    for i, (_, rw) in enumerate(crit.iterrows(), DR + 1):
        ws.cell(i, 1, rw['area']); ws.cell(i, 2, float(rw['indice_0_10']))
    n_crit = DR + len(crit)

    # Bloque B: minutos perdidos — cols E,F
    ws.cell(DR, 5, 'Área'); ws.cell(DR, 6, 'Min Perdidos')
    for i, (_, rw) in enumerate(perdida.iterrows(), DR + 1):
        ws.cell(i, 5, rw['area']); ws.cell(i, 6, float(rw['minutos_perdidos']))
    n_perd = DR + len(perdida)

    # Bloque C: horas asig vs asist — cols H,I,J
    ws.cell(DR, 8, 'Área'); ws.cell(DR, 9, 'Hrs Asign'); ws.cell(DR, 10, 'Hrs Asist')
    area_s = area_hrs.sort_values('horas_asign_total')
    for i, (_, rw) in enumerate(area_s.iterrows(), DR + 1):
        ws.cell(i, 8, rw['area'])
        ws.cell(i, 9,  float(rw['horas_asign_total']))
        ws.cell(i, 10, float(rw['horas_asist_total']))
    n_hrs = DR + len(area_hrs)

    # Bloque D: promedio atraso — cols L,M
    ws.cell(DR, 12, 'Área'); ws.cell(DR, 13, 'Prom Atraso')
    for i, (_, rw) in enumerate(por_area_atr.iterrows(), DR + 1):
        ws.cell(i, 12, rw['area']); ws.cell(i, 13, float(rw['prom_atraso_min']))
    n_atr = DR + len(por_area_atr)

    # Bloque E: permisos — cols O,P
    ws.cell(DR, 15, 'Nombre'); ws.cell(DR, 16, 'Vacaciones')
    for i, (_, rw) in enumerate(perms.iterrows(), DR + 1):
        ws.cell(i, 15, rw['nombre']); ws.cell(i, 16, float(rw['vacaciones']))
    n_perms = DR + len(perms)

    # ---- Gráfico 1: Índice criticidad (barras horiz) ----
    g1 = BarChart(); g1.type = "bar"; g1.style = 10
    g1.title = "Índice de Criticidad por Área (0–10)"
    g1.legend = None; g1.width = 14; g1.height = 12
    g1.add_data(Reference(ws, min_col=2, min_row=DR, max_row=n_crit), titles_from_data=True)
    g1.set_categories(Reference(ws, min_col=1, min_row=DR + 1, max_row=n_crit))
    g1.series[0].graphicalProperties.solidFill = COLOR_MED
    g1.series[0].graphicalProperties.line.solidFill = COLOR_OSC
    ws.add_chart(g1, "A7")

    # ---- Gráfico 2: Minutos perdidos ----
    g2 = BarChart(); g2.type = "col"; g2.style = 10
    g2.title = "Minutos Perdidos vs Turno por Área"
    g2.legend = None; g2.width = 14; g2.height = 12
    g2.add_data(Reference(ws, min_col=6, min_row=DR, max_row=n_perd), titles_from_data=True)
    g2.set_categories(Reference(ws, min_col=5, min_row=DR + 1, max_row=n_perd))
    g2.series[0].graphicalProperties.solidFill = "fa4440"
    g2.series[0].graphicalProperties.line.solidFill = COLOR_MED
    ws.add_chart(g2, "J7")

    # ---- Gráfico 3: Horas asig vs asist ----
    g3 = BarChart(); g3.type = "bar"; g3.style = 10
    g3.title = "Horas Asignadas vs. Asistidas por Área"
    g3.width = 14; g3.height = 12
    g3.add_data(Reference(ws, min_col=9, max_col=10, min_row=DR, max_row=n_hrs), titles_from_data=True)
    g3.set_categories(Reference(ws, min_col=8, min_row=DR + 1, max_row=n_hrs))
    g3.series[0].graphicalProperties.solidFill = COLOR_OSC
    g3.series[1].graphicalProperties.solidFill = "fa4440"
    ws.add_chart(g3, "A29")

    # ---- Gráfico 4: Promedio atraso ----
    g4 = BarChart(); g4.type = "col"; g4.style = 10
    g4.title = "Promedio Minutos de Atraso por Área"
    g4.legend = None; g4.width = 14; g4.height = 12
    g4.add_data(Reference(ws, min_col=13, min_row=DR, max_row=n_atr), titles_from_data=True)
    g4.set_categories(Reference(ws, min_col=12, min_row=DR + 1, max_row=n_atr))
    g4.series[0].graphicalProperties.solidFill = COLOR_MED
    g4.series[0].graphicalProperties.line.solidFill = COLOR_OSC
    ws.add_chart(g4, "J29")

    # ---- Gráfico 5: Vacaciones por persona ----
    g5 = BarChart(); g5.type = "col"; g5.style = 10
    g5.title = "Días de Vacaciones / Permisos por Persona"
    g5.legend = None; g5.width = 20; g5.height = 12
    g5.add_data(Reference(ws, min_col=16, min_row=DR, max_row=n_perms), titles_from_data=True)
    g5.set_categories(Reference(ws, min_col=15, min_row=DR + 1, max_row=n_perms))
    g5.series[0].graphicalProperties.solidFill = COLOR_OSC
    ws.add_chart(g5, "A51")

def generar_excel(resultados, path_salida):
    (df_gen, perms, df_atr, personas_atr, por_area_atr, max_atr,
     por_area_sal, por_persona_hrs, por_area_hrs, perdida, crit, n_anomalias) = resultados

    wb = Workbook()
    wb.remove(wb.active)

    sheet_dashboard(wb, crit, por_area_atr, perdida, por_area_hrs, perms)
    sheet_resumen(wb, crit, n_anomalias)
    sheet_permisos(wb, perms)
    sheet_atrasos(wb, por_area_atr, max_atr, personas_atr)
    sheet_horas(wb, por_persona_hrs, por_area_hrs, perdida)
    sheet_minutos_perdidos(wb, perdida)

    wb.save(path_salida)
    print(f"✓ Excel guardado en: {path_salida}")

# ===========================================================
# 5. MAIN
# ===========================================================
def main():
    if len(sys.argv) < 4:
        print(__doc__)
        sys.exit(1)

    path_general = sys.argv[1]
    path_semanal = sys.argv[2]
    path_nomina  = sys.argv[3]

    for p in [path_general, path_semanal, path_nomina]:
        if not os.path.exists(p):
            print(f"Error: no se encontró el archivo '{p}'")
            sys.exit(1)

    print("Cargando datos...")
    df, df_gen, df_nom = construir_df(path_general, path_semanal, path_nomina)

    print("Analizando permisos...")
    perms = analizar_permisos(df_gen, df_nom)

    print("Analizando atrasos...")
    df_atr, personas_atr, por_area_atr, max_atr = analizar_atrasos(df)

    # Contar anomalías excluidas
    mask_raw = ~df['con_permiso'] & df['et_min'].notna() & df['er_min'].notna()
    n_anomalias = int(
        ((df[mask_raw]['er_min'] >= df[mask_raw]['st_min'].fillna(9999)) |
         ((df[mask_raw]['er_min'] - df[mask_raw]['et_min']) > UMBRAL_ATRASO_MAX)).sum()
    )

    print("Analizando salidas...")
    por_area_sal = analizar_salidas(df)

    print("Analizando horas...")
    por_persona_hrs, por_area_hrs, perdida = analizar_horas(df)

    print("Calculando índice de criticidad...")
    crit = calcular_criticidad(df_atr)

    # Nombre de salida: siempre en la carpeta donde se ejecuta el script
    nombre_salida = "analisis_asistencia.xlsx"
    path_salida = os.path.join(os.getcwd(), nombre_salida)

    print("Generando Excel...")
    resultados = (
        df_gen, perms, df_atr, personas_atr, por_area_atr, max_atr,
        por_area_sal, por_persona_hrs, por_area_hrs, perdida, crit, n_anomalias,
    )
    generar_excel(resultados, path_salida)

    # Resumen en consola
    print("\n" + "="*55)
    print("RESUMEN")
    print("="*55)
    print(f"  Trabajadores únicos:       {df['rut'].nunique()}")
    print(f"  Días con permiso/vac:      {int(perms['total'].sum())}")
    print(f"  Personas con atrasos:      {len(personas_atr)}")
    print(f"  Total minutos atraso:      {int(por_area_atr['total_min_atraso'].sum())}")
    print(f"  Total minutos perdidos:    {int(perdida['minutos_perdidos'].sum())}")
    print(f"  Marcaciones anómalas excl: {n_anomalias}")
    print(f"  Área más crítica:          {crit.iloc[0]['area']} (índice {crit.iloc[0]['indice_0_10']})")
    print("="*55)

if __name__ == "__main__":
    main()
