"""
SISTEMA DE FACTURACIÓN INTERCOMPANY - VERSIÓN OPTIMIZADA
"""

import pandas as pd
import numpy as np
import os
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')

# ============================================================================
# CONFIGURACIÓN
# ============================================================================

RUTA_BALANCE = None  # se setea desde app.py al subir el GL
RUTA_ORDENES = None  # se setea desde app.py al subir las Órdenes
RUTA_TC      = None  # se setea desde app.py al subir el Tipo de Cambio
RUTA_ADI     = None  # se setea desde app.py al subir el ADI
RUTA_RUBROS  = None  # se setea desde app.py al subir el Template Maestro
# NOTA: Glosario y PL ahora se leen del template_rubros.xlsx (hojas Productos y PL_Totalizador)
# Las variables RUTA_GLOSARIO y RUTA_PL ya no son necesarias como archivos separados.

# ---- Tasas y markups intercompany (Ecuador) ----
TASA_MARCA         = 0.025   # 2.5%  — Licencia de Marca (Ecuador)   ⚠️ Confirmar en contrato
TASA_IT            = 0.05    # 5.0%  — Licencia de IT    (Ecuador)   ⚠️ Confirmar en contrato

MAPEO_PAIS_POS = {
    'Argentina': 101, 'Mexico': 103, 'Brasil': 104, 'Usa': 105, 'Chile': 108,
    'Ecuador': 113, 'Peru': 115, 'Colombia': 116, 'Bolivia': 601, 'Costa Rica': 601,
    'El Salvador': 601, 'España': 601, 'Guatemala': 601, 'Honduras': 601,
    'Nicaragua': 601, 'Panama': 601, 'Paraguay': 601, 'Puerto Rico': 601, 'Uruguay': 601
}

# ---- Configuración por entidad (fuente: Lineamientos Intercompany) ----

# markup_rfc = porcentaje del markup de garantía de margen segmento RFC (ya como decimal)

# ⚠️ TODOS los markup_rfc requieren confirmación periódica contra los acuerdos de precios de transferencia vigentes
ENTITY_CONFIG = {
    101: {
        'nombre':     'Argentina',
        'markup_rfc': 0.095,   # 9.5% — RFC Argentina  ⚠️ Confirmar en acuerdo PT
        # ⚠️ PENDIENTE: DOM/INT se determina por campo 'Negocio2' del flex (no posición 9 como en Ecuador)
        # ⚠️ PENDIENTE: cuenta "ya-contabilizado" = 70104/70106 (diferente a Ecuador que usa 70105)
        # ⚠️ PENDIENTE: excluir costos Desarrollo Software y Call Center de la base corporativa
    },
    # 103 (México) movido a bloque validado más abajo
    104: {
        'nombre':     'Brasil',
        'markup_rfc': 0.105,   # 10.5% — RFC Brasil (markup diferencial)  ⚠️ Confirmar en acuerdo PT
        # ⚠️ PENDIENTE: implementar DOS flujos RFC separados (TRV y España) según lineamientos
        # ⚠️ PENDIENTE: excluir costos Desarrollo Software, Call Center y Soporte Local
    },
    105: {
        'nombre':         'Usa',
        'markup_rfc':     0.095,   # 9.5% — RFC USA           ⚠️ Confirmar en acuerdo PT
        'markup_hosting': 0.060,   # 6.0% — Servicios Hosting  ⚠️ Confirmar en acuerdo PT
        # ⚠️ PENDIENTE: validar estructura de flex field vs balance USA
        # Flujos: (1) Hosting → USA presta a AR  (2) RFC → USA presta a Travel Res. UY
        # Cuentas clave: 49102 (RFC), 49120 (Hosting), 70104 (Marca booked), 70105 (IT booked)
    },
    108: {
        'nombre':     'Chile',
        'markup_rfc': 0.095,   # 9.5% — RFC Chile      ⚠️ Confirmar en acuerdo PT
        # ⚠️ PENDIENTE: validar estructura de flex field vs balance Chile
    },
    113: {
        'nombre':     'Ecuador',
        'markup_rfc': 0.095,   # 9.5% — RFC Ecuador    ✅ Validado contra balance Oracle nov-2025
        # ✅ Metodología completa implementada y validada. Ver calcular_ecuador_hoja_llave()
    },
    115: {
        'nombre':     'Peru',
        'markup_rfc': 0.095,   # 9.5% — RFC Perú       ⚠️ Confirmar en acuerdo PT
        # ⚠️ PENDIENTE: validar estructura de flex field vs balance Perú
    },
    116: {
        'nombre':     'Colombia',
        'markup_rfc': 0.095,   # 9.5% — RFC Colombia   ⚠️ Confirmar en acuerdo PT
        # ⚠️ PENDIENTE: validar estructura de flex field vs balance Colombia
        # ⚠️ PENDIENTE: excluir costos Call Center de la base corporativa
    },
    103: {
        'nombre':     'Mexico',
        'markup_rfc': 0.095,   # 9.5% — RFC México     ✅ Validado Nov-2025
        # ⚠️ ALERTA: RFC usa clasificación RUBRO departamental que NO existe en PL_actualizado.xlsx
        # Los costos por producto (H29, I29, J37) se hardcodean del procesamiento intermedio del Excel
        # Solución: PL_actualizado.xlsx debería incorporar el nivel de detalle departamental
    },
    601: {
        'nombre':     'Travel Reservation',
        # Solo Licencia IT (5%) — no tiene Marca, RFC ni Hosting
        # Prestador: Despegar.com.ar (AR) → Prestatario: Travel Reservations S.R.L (UY)
        # Moneda: USD (nativa Oracle, sin conversión)
    },
    611: {
        'nombre':     'España',
        # Licencia IT (5%) + Licencia Marca (2.5%) — no tiene RFC ni Hosting
        # Prestador IT: Despegar.com.ar (AR) → Prestatario: Despegar España
        # Prestador Marca: Travel Reservations (601) → Prestatario: Despegar España
        # Moneda: EUR → divide por tc_eur (EUR/USD) para resultado en USD
    },
}

# MARKUP_RFC para Ecuador — derivado de ENTITY_CONFIG, no hardcodeado
MARKUP_RFC = 1 + ENTITY_CONFIG[113]['markup_rfc']

# ============================================================================
# CONSTANTES DE PRODUCTO Y COSTOS — compartidas por todas las entidades
# ============================================================================

# Mapeo código producto Oracle → categoría interna
# PROD_MAP: Se construye dinámicamente desde template_rubros.xlsx (hoja Productos)
# Mapea Product_Flex (posición 4 de Combinacion Contable) → categoría de negocio
# (VUELOS, HT, ONA, CORPORATE, PUBLICIDAD, VUELOS PAQUETES)
# ⚠️ Si aparece un código nuevo en el GL que no está en el template,
#    cae a 'OTRO' y queda fuera del cálculo. Ver alerta 1 al final del output.
PROD_MAP = {}  # se carga en ejecutar_proceso_completo() desde template_rubros.xlsx

# DI_MAP: Se construye dinámicamente desde template_rubros.xlsx (hoja DOM_INT)
# Mapea DomInt_Flex (posición 9 de Combinacion Contable) → DOM/INT/Iniciativas
# Valores numéricos: 1=DOM, 2=INT, 101=DOM, 102=INT, 111=DOM, 112=INT, etc.
DI_MAP = {}    # se carga en ejecutar_proceso_completo() desde template_rubros.xlsx
DI_INT = set() # conjunto de códigos DomInt que son INT (derivado de DI_MAP)

# Cuentas excluidas de costos directos HT/ONA en el cálculo de RFC
# (64201/64202 se excluyen de E18/F18 pero SÍ entran en M5 corporativo)
# ⚠️ Confirmar si se agregan/eliminan cuentas por cambio de metodología
EXCL_COST_ACCOUNTS = {68210, 41311, 41312, 64201, 64202, 53300, 53305}

# RFC_COST_RUBROS: Se construye dinámicamente desde template_rubros.xlsx
# Toma todos los RUBRO del template excepto NET REVENUE y FINANCE,
# y los mapea a PL_Totalizador para filtrar costos del RFC.
RFC_COST_RUBROS = set()  # se carga en ejecutar_proceso_completo() desde el template

# ── Constantes específicas USA (Mapeo PL Nivel 1 nativo del balance Oracle) ──
# Categorías CORPORATE incluidas en la base de costo del RFC de USA.
# Se excluyen: impuestos, comisiones CC, marketing directo, intereses, FX.
# ✅ Validado contra Despegar USA Segmentación Nov-2025.xlsx
CORP_RFC_PL1_USA = {
    'Channels-Personnel/Expenses',
    'Finance & Admin.',
    'General Management',
    'IT - Personnel/Expenses',
}

# Categorías de incentivos de ingresos — se excluyen de E18/F18 (son ajustes
# de revenue, no costos operativos directos de HT/ONA)
INCENTIVE_PL1_USA = {'Back End Incentives', 'Other Incentives', 'Up Front Incentives'}

# Categorías no operativas — se excluyen de E18/F18 (intereses, diferencias de cambio)
NON_OP_PL1_USA = {'Interest Income', 'FX/Other'}

# Exclusión combinada para costos directos HT/ONA (E18/F18)
# ✅ 'Customer Fees & Charges' excluido: no aparece en Estimado filas 18-21 (fuera de
#    COST OF REVENUE / S&M / TECHNOLOGY / GENERAL usados por SUM(E18:E21) del Excel)
EXCL_PL1_USA = NON_OP_PL1_USA | INCENTIVE_PL1_USA | {'Customer Fees & Charges'}

# Correcciones de cuentas para M5 CORPORATE (validado vs Estimado!M6 Excel):
# EXCL: 80101 (Gastos bancarios) — PL1='General Management' pero Excel lo ubica en sección
#       FINANCE (filas 76-87 del Estimado, fuera del SUM M6) → excluir de M5
# EXTRA: 80102 (Comisión de Tarjetas) y 64341 (Branding) — PL1 excluido por CORP_RFC_PL1_USA
#        pero Excel los incluye en M6 (D32 y D49:D57 respectivamente para CORPORATE)
# Net validado: −18,771 + 15,116 + 988 = −2,666.58 → M5 = 2,483,539.18 ✅
CORP_M5_EXCL_ACCOUNTS  = {80101}          # Excluir a pesar de PL1='General Management'
CORP_M5_EXTRA_ACCOUNTS = {80102, 64341}   # Incluir a pesar de PL1 no ∈ CORP_RFC_PL1_USA

# Categoría PL Nivel 1 que define los Costos de Hosting en USA (IT CORPORATE)
# ✅ Validado: 'IT - Personnel/Expenses' CORPORATE non-IC = 1,394,173.78
HOSTING_PL1_USA = 'IT - Personnel/Expenses'

# ============================================================================
# FUNCIONES DE CARGA
# ============================================================================

def leer_periodo_balance(archivo):
    MESES = {
        'ENE': 1, 'FEB': 2, 'MAR': 3, 'ABR': 4, 'MAY': 5, 'JUN': 6,
        'JUL': 7, 'AGO': 8, 'SEP': 9, 'OCT': 10, 'NOV': 11, 'DIC': 12,
        'JAN': 1, 'APR': 4, 'AUG': 8, 'DEC': 12,
    }
    if not os.path.exists(archivo):
        return None
    try:
        with open(archivo, 'r', encoding='latin1') as f:
            for line in f:
                if line.count(';') > 5:
                    break
                if 'Periodo' in line or 'Período' in line or 'Period' in line:
                    token = line.split(':')[-1].strip().split('-')[0].upper()
                    if token in MESES:
                        return MESES[token]
    except Exception:
        pass
    return None


def cargar_glosario(ruta, sheet_name=None):
    if not os.path.exists(ruta):
        return None
    try:
        if sheet_name:
            df = pd.read_excel(ruta, sheet_name=sheet_name)
        else:
            # Intentar primero como hoja Productos del template
            try:
                df = pd.read_excel(ruta, sheet_name='Productos')
            except:
                df = pd.read_excel(ruta)
        df.columns = df.columns.astype(str).str.strip()
        col_cod = next((c for c in df.columns if 'COD' in c.upper() and 'CATEGORIA' not in c.upper()), None)
        col_seg = next((c for c in df.columns if 'SEGMENTACION' in c.upper()), None)
        col_cat = next((c for c in df.columns if 'CATEGORIA APERTURA ONA' in c.upper() or 'CATEGORIA' in c.upper()), None)
        if col_cod:
            df[col_cod] = pd.to_numeric(df[col_cod], errors='coerce')
            df = df.dropna(subset=[col_cod])
            df[col_cod] = df[col_cod].astype(int)
            cols_out = {col_cod: 'COD'}
            if col_seg:
                cols_out[col_seg] = 'SEGMENTACION'
            if col_cat:
                cols_out[col_cat] = 'CATEGORIA_APERTURA_ONA'
            return df[list(cols_out.keys())].rename(columns=cols_out)
    except:
        pass
    return None

def cargar_pl(ruta, sheet_name=None):
    if not os.path.exists(ruta):
        return None
    try:
        if sheet_name:
            df = pd.read_excel(ruta, sheet_name=sheet_name)
        else:
            # Intentar primero como hoja PL_Totalizador del template
            try:
                df = pd.read_excel(ruta, sheet_name='PL_Totalizador')
            except:
                df = pd.read_excel(ruta)
        df.columns = df.columns.astype(str).str.strip()
        if 'Cuenta' in df.columns and 'PL_Totalizador' in df.columns:
            df['Cuenta'] = pd.to_numeric(df['Cuenta'], errors='coerce')
            return df[['Cuenta', 'PL_Totalizador']].dropna(subset=['Cuenta'])
    except:
        pass
    return None

def cargar_tc(ruta):
    """
    Carga el archivo de Tipos de Cambio EPM y retorna un dict {moneda: tc}.
    Monedas en formato ISO (CLP, ARS, MXN, BRL, COP, PEN, UYU, ...).
    TC expresado en unidades de moneda local por 1 USD.

    Formato EPM: la primera columna tiene los códigos de moneda (USD, CLP, etc.)
    y la segunda columna tiene el valor del TC. Las filas de datos comienzan
    inmediatamente después de la fila que contiene 'TC' en la segunda columna.
    Si el archivo no se encuentra o no puede parsearse, lanza un error claro.
    """
    if not os.path.exists(ruta):
        raise FileNotFoundError(f"Archivo TC no encontrado: {ruta}")
    try:
        df = pd.read_excel(ruta, header=None)
        # Encontrar la fila del sub-header ('TC' en col 1)
        header_row = None
        for i, row in df.iterrows():
            if str(row.iloc[1]).strip().upper() == 'TC':
                header_row = i
                break
        if header_row is None:
            raise ValueError("No se encontró el sub-header 'TC' en el archivo.")
        # Los datos de moneda/TC empiezan en la fila siguiente al sub-header
        datos = df.iloc[header_row + 1:, [0, 1]].copy()
        datos.columns = ['moneda', 'tc']
        datos['moneda'] = datos['moneda'].astype(str).str.strip().str.upper()
        datos['tc'] = pd.to_numeric(datos['tc'], errors='coerce')
        tc_dict = {row['moneda']: row['tc']
                   for _, row in datos.iterrows()
                   if pd.notna(row['tc']) and row['tc'] > 0 and row['moneda'] not in ('NAN', '')}
        if not tc_dict:
            raise ValueError("El archivo TC no contiene valores válidos.")
        return tc_dict
    except Exception as e:
        raise RuntimeError(f"Error al leer el archivo TC '{ruta}': {e}")


def procesar_balance(archivo):
    if not os.path.exists(archivo):
        return None
    try:
        start_row = 0
        with open(archivo, 'r', encoding='latin1') as f:
            for i, line in enumerate(f):
                if line.count(';') > 5:
                    start_row = i
                    break

        df = pd.read_csv(archivo, sep=';', skiprows=start_row, encoding='latin1')
        col_balance = next((c for c in df.columns if 'Saldo Final' in c or 'Balance' in c), df.columns[-1])
        df[col_balance] = pd.to_numeric(df[col_balance], errors='coerce').fillna(0)

        col_flex    = next((c for c in df.columns if 'Combinacion' in c or 'Flex' in c), None)
        col_account = next((c for c in df.columns if 'Cuenta' in c or 'Account' in c), None)

        if col_flex and col_account:
            def parse_flex(s):
                if not isinstance(s, str):
                    return pd.Series([None, None, None])
                parts = s.split('.')
                return pd.Series([
                    parts[0] if len(parts) > 0 else None,
                    parts[4] if len(parts) > 4 else None,
                    parts[9] if len(parts) > 9 else None,
                ])

            df[['Entity_Flex', 'Product_Flex', 'DomInt_Flex']] = \
                df[col_flex].astype(str).apply(parse_flex)
            df['Account_Num'] = pd.to_numeric(df[col_account], errors='coerce')
            return df
    except:
        pass
    return None

def procesar_ordenes(archivo):
    if not os.path.exists(archivo):
        return None, None

    df = pd.read_excel(archivo)
    for c in ['site', 'line_of_business_code', 'trip_type_code', 'buy_type_code', 'product']:
        if c in df.columns:
            df[c] = df[c].astype(str).str.strip()

    df['POS_EDITADO'] = df['site'].map(lambda x: MAPEO_PAIS_POS.get(str(x).strip(), 601))

    def get_producto_1(bt):
        return 'Vuelos' if bt == 'Vuelos' else 'Hoteles' if bt == 'Hoteles' else 'ONA'

    df['Producto_1'] = df['buy_type_code'].map(get_producto_1)
    df['Producto_2'] = df.apply(lambda r: 'Vuelos Paquetes' if r['product'] == 'Vuelos' and
                                r['buy_type_code'] == 'Carrito' else r['Producto_1'], axis=1)

    df['DOM_INT'] = df.apply(lambda r: 'DOM' if r['Producto_2'] in ['Vuelos', 'Vuelos Paquetes']
                             else ('Int' if r['trip_type_code'] == 'Int' else 'Dom'), axis=1)

    df['Entidad_Facturar'] = df.apply(
        lambda r: MAPEO_PAIS_POS.get(r['site'], 601) if
        (r['buy_type_code'] == 'Vuelos' or (r['product'] == 'Vuelos' and r['buy_type_code'] == 'Carrito'))
        else (601 if r['trip_type_code'] == 'Int' else MAPEO_PAIS_POS.get(r['site'], 601)), axis=1
    )

    df['Usar'] = df.apply(lambda r: 'NO' if (r['site'] == 'Mexico' and
                          r['line_of_business_code'] == 'B2B') else 'SI', axis=1)

    return df, df[df['Usar'] == 'SI'].copy()

def procesar_adi(archivo):
    if not os.path.exists(archivo):
        return None
    try:
        xl = pd.ExcelFile(archivo)
        sheet = 'Completo' if 'Completo' in xl.sheet_names else xl.sheet_names[0]
        df = pd.read_excel(archivo, sheet_name=sheet)
        return df[df['Estado'] == 1].copy()
    except:
        return None

# ============================================================================
# FUNCIONES DE IMPRESIÓN (formato detallado)
# ============================================================================

SEP = "=" * 80

def _header(titulo):
    print(f"\n{SEP}")
    print(f"  {titulo}")
    print(SEP)

def _sep_thin():
    print("  " + "-" * 60)

def imprimir_serie(serie, col_label="Entidad", pct_label="Participación %", decimales=2):
    """Imprime una Series como tabla."""
    if serie is None or len(serie) == 0:
        print("  ⚠️  Sin datos")
        return
    fmt = f"  {{:<28}} {{:>12.{decimales}f}}%"
    _sep_thin()
    for idx, val in serie.items():
        print(fmt.format(str(idx), float(val)))
    _sep_thin()
    total = serie.sum()
    print(f"  {'TOTAL':<28} {total:>12.{decimales}f}%")

def imprimir_dataframe(df, decimales=2):
    """Imprime un DataFrame como tabla. La columna TOTAL es la suma por fila (cada mercado = 100%)."""
    if df is None or df.empty:
        print("  ⚠️  Sin datos")
        return
    cols = list(df.columns)
    # Header con columna TOTAL al final
    header = f"  {'':>8} " + "".join(f"{str(c):>14}" for c in cols) + f"{'TOTAL':>14}"
    print(header)
    _sep_thin()
    for idx, row in df.iterrows():
        row_total = sum(float(v) for v in row)
        line = f"  {str(idx):>8} " + "".join(f"{float(v):>13.{decimales}f}%" for v in row) \
               + f"{row_total:>13.{decimales}f}%"
        print(line)
    _sep_thin()

def imprimir_pivot(pivot, decimales=2):
    """Imprime un pivot (MultiIndex columns)."""
    if pivot is None or pivot.empty:
        print("  ⚠️  Sin datos")
        return
    # Flatten columns if MultiIndex
    if isinstance(pivot.columns, pd.MultiIndex):
        pivot = pivot.copy()
        pivot.columns = [f"{a}/{b}" for a, b in pivot.columns]
    cols = list(pivot.columns)
    header = f"  {'Entidad':>8} " + "".join(f"{str(c):>16}" for c in cols)
    print(header)
    _sep_thin()
    for idx, row in pivot.iterrows():
        vals = "".join(f"{float(v):>15.{decimales}f}%" for v in row)
        print(f"  {str(idx):>8} {vals}")
    _sep_thin()
    totals = pivot.sum(axis=0)
    total_line = f"  {'TOTAL':>8} " + "".join(f"{float(v):>15.{decimales}f}%" for v in totals)
    print(total_line)

# ============================================================================
# CÁLCULO DE LLAVES
# ============================================================================

def calcular_llaves_ordenes(df_si):
    llaves = {}
    total = df_si['Bookings Emitidos'].sum()

    df_si['_Prod'] = df_si.apply(
        lambda r: 'Vuelos Paquetes' if r['product'] == 'Vuelos' and r['buy_type_code'] == 'Carrito'
        else r['buy_type_code'] if r['buy_type_code'] in ['Vuelos', 'Hoteles'] else 'ONA', axis=1
    )

    # ── K1: Producto Global ────────────────────────────────────────────────────
    llaves['K1'] = df_si.groupby('_Prod')['Bookings Emitidos'].sum() / total * 100
    _header("📌 LLAVE K1: PRODUCTO GLOBAL")
    print(f"  Total bookings: {total:,.0f}")
    imprimir_serie(llaves['K1'], col_label="Producto")

    # ── K2: Por Entidad Base ───────────────────────────────────────────────────
    llaves['K2'] = df_si.groupby('Entidad_Facturar')['Bookings Emitidos'].sum() / total * 100
    _header("📌 LLAVE K2: POR ENTIDAD (BASE)")
    imprimir_serie(llaves['K2'])

    # ── K3: Sin Brasil ─────────────────────────────────────────────────────────
    df_nb = df_si[df_si['site'] != 'Brasil']
    total_nb = df_nb['Bookings Emitidos'].sum()
    llaves['K3'] = df_nb.groupby('Entidad_Facturar')['Bookings Emitidos'].sum() / total_nb * 100
    _header("📌 LLAVE K3: SIN BRASIL")
    print(f"  Total bookings (excl Brasil): {total_nb:,.0f}")
    imprimir_serie(llaves['K3'])

    # ── K4: 150% ──────────────────────────────────────────────────────────────
    k4 = {ent: val * 0.5 for ent, val in llaves['K2'].items()}
    k4[106] = k4.get(106, 0) + 50.0
    k4[601] = k4.get(601, 0) + 50.0
    llaves['K4'] = pd.Series(k4)
    _header("📌 LLAVE K4: 150% (Inc Decolar + 601)")
    imprimir_serie(llaves['K4'])

    # ── K5: 100% Inc ──────────────────────────────────────────────────────────
    llaves['K5'] = pd.Series({106: 100.00})
    _header("📌 LLAVE K5: 100% DECOLAR INC")
    imprimir_serie(llaves['K5'])

    # ── K6: Matriz Vuelos/HT/ONA por DOM/INT ──────────────────────────────────
    df_si['_K6_Prod'] = df_si['buy_type_code'].map(
        lambda x: 'Vuelos' if x == 'Vuelos' else 'Hoteles' if x == 'Hoteles' else 'ONA'
    )
    df_si['_K6_Tipo'] = df_si.apply(
        lambda r: 'Dom' if r['_K6_Prod'] == 'Vuelos' or (r['buy_type_code'] == 'Carrito' and r['product'] == 'Vuelos')
        else ('Int' if r['trip_type_code'] == 'Int' else 'Dom'), axis=1
    )
    pivot_6 = df_si.pivot_table(
        index='Entidad_Facturar', columns=['_K6_Prod', '_K6_Tipo'],
        values='Bookings Emitidos', aggfunc='sum', fill_value=0
    )
    llaves['K6'] = (pivot_6 / total * 100)
    _header("📌 LLAVE K6: MATRIZ VUELOS/HT/ONA × DOM/INT")
    imprimir_pivot(llaves['K6'])

    # ── K7: DOM-INT Total (igual K2) ──────────────────────────────────────────
    llaves['K7'] = llaves['K2']
    _header("📌 LLAVE K7: DOM-INT TOTAL POR PAÍS (= K2)")
    imprimir_serie(llaves['K7'])

    # ── K8: Productos por País (sin Argentina) ────────────────────────────────
    df_k8 = df_si[df_si['site'] != 'Argentina'].copy()
    p8 = df_k8.pivot_table(
        index='Entidad_Facturar', columns='Producto_1',
        values='Bookings Emitidos', aggfunc='sum', fill_value=0
    )
    llaves['K8'] = p8.div(p8.sum(axis=1), axis=0) * 100
    _header("📌 LLAVE K8: PRODUCTOS POR PAÍS (sin Argentina)")
    imprimir_dataframe(llaves['K8'])

    # ── K9: Dom/Int por POS ───────────────────────────────────────────────────
    llave_9_raw = df_si.pivot_table(
        index='POS_EDITADO', columns='DOM_INT',
        values='Bookings Emitidos', aggfunc='sum', fill_value=0
    )
    dom_total = pd.Series(0, index=llave_9_raw.index)
    if 'DOM' in llave_9_raw.columns:
        dom_total += llave_9_raw['DOM']
    if 'Dom' in llave_9_raw.columns:
        dom_total += llave_9_raw['Dom']
    int_total = llave_9_raw.get('Int', pd.Series(0, index=llave_9_raw.index))
    llave_9 = pd.DataFrame({'Dom': dom_total, 'Int': int_total})
    llaves['K9'] = llave_9.div(llave_9.sum(axis=1), axis=0) * 100
    _header("📌 LLAVE K9: DOM / INT POR POS")
    imprimir_dataframe(llaves['K9'])

    # ── K10: Producto Detallado por POS ───────────────────────────────────────
    llaves['K10'] = df_si.pivot_table(
        index='POS_EDITADO', columns='_Prod',
        values='Bookings Emitidos', aggfunc='sum', fill_value=0
    ).div(df_si.groupby('POS_EDITADO')['Bookings Emitidos'].sum(), axis=0) * 100
    _header("📌 LLAVE K10: PRODUCTO DETALLADO POR POS")
    imprimir_dataframe(llaves['K10'])

    # ── K11: Hoteles Dom/Int por POS ──────────────────────────────────────────
    df_ht = df_si[df_si['_Prod'] == 'Hoteles']
    if not df_ht.empty:
        llave_11_raw = df_ht.pivot_table(
            index='POS_EDITADO', columns='DOM_INT',
            values='Bookings Emitidos', aggfunc='sum', fill_value=0
        )
        dom_11 = pd.Series(0, index=llave_11_raw.index)
        if 'DOM' in llave_11_raw.columns: dom_11 += llave_11_raw['DOM']
        if 'Dom' in llave_11_raw.columns: dom_11 += llave_11_raw['Dom']
        int_11 = llave_11_raw.get('Int', pd.Series(0, index=llave_11_raw.index))
        llave_11 = pd.DataFrame({'Dom': dom_11, 'Int': int_11})
        llaves['K11'] = llave_11.div(llave_11.sum(axis=1), axis=0) * 100
        _header("📌 LLAVE K11: HOTELES DOM / INT POR POS")
        imprimir_dataframe(llaves['K11'])
    else:
        llaves['K11'] = None
        _header("📌 LLAVE K11: HOTELES DOM / INT POR POS")
        print("  ⚠️  Sin datos de Hoteles")

    # ── K12: ONA Dom/Int por POS ──────────────────────────────────────────────
    df_ona = df_si[df_si['_Prod'].isin(['ONA', 'Vuelos Paquetes'])]
    if not df_ona.empty:
        llave_12_raw = df_ona.pivot_table(
            index='POS_EDITADO', columns='DOM_INT',
            values='Bookings Emitidos', aggfunc='sum', fill_value=0
        )
        dom_12 = pd.Series(0, index=llave_12_raw.index)
        if 'DOM' in llave_12_raw.columns: dom_12 += llave_12_raw['DOM']
        if 'Dom' in llave_12_raw.columns: dom_12 += llave_12_raw['Dom']
        int_12 = llave_12_raw.get('Int', pd.Series(0, index=llave_12_raw.index))
        llave_12 = pd.DataFrame({'Dom': dom_12, 'Int': int_12})
        llaves['K12'] = llave_12.div(llave_12.sum(axis=1), axis=0) * 100
        _header("📌 LLAVE K12: ONA DOM / INT POR POS (incluye Vuelos Paquetes)")
        imprimir_dataframe(llaves['K12'])
    else:
        llaves['K12'] = None
        _header("📌 LLAVE K12: ONA DOM / INT POR POS")
        print("  ⚠️  Sin datos de ONA")

    return llaves

def calcular_llaves_adi(df_adi):
    llaves_adi = {}

    # ── ADI1: HC por Entidad ───────────────────────────────────────────────────
    hc_por_entidad = df_adi.groupby('COD. E.L.')['ID'].nunique()
    total_hc = hc_por_entidad.sum()
    llaves_adi['ADI1'] = (hc_por_entidad / total_hc * 100)
    _header("📌 LLAVE ADI1: HC POR ENTIDAD (Soporte Regional)")
    print(f"  Total HC: {total_hc:,}")
    imprimir_serie(llaves_adi['ADI1'])

    # ── ADI2: Call Center Colombia ─────────────────────────────────────────────
    df_colombia = df_adi[df_adi['COD. E.L.'] == 116].copy()
    if not df_colombia.empty:
        total_colombia = len(df_colombia)
        df_cc = df_colombia[df_colombia['COD. R.C.'].isin([300, 310])].copy()
        total_cc = len(df_cc)
        porcentaje_cc = round((total_cc / total_colombia * 100), 2)
        llaves_adi['ADI2'] = pd.Series({116: porcentaje_cc})
        _header("📌 LLAVE ADI2: CALL CENTER COLOMBIA")
        print(f"  Total Colombia: {total_colombia:,}  |  CC (300/310): {total_cc:,}")
        imprimir_serie(llaves_adi['ADI2'])
    else:
        llaves_adi['ADI2'] = None
        _header("📌 LLAVE ADI2: CALL CENTER COLOMBIA")
        print("  ⚠️  Sin datos de Colombia")

    return llaves_adi

def calcular_llaves_revenue(df_balance, df_glosario):
    if df_balance is None or df_glosario is None:
        return {}

    llaves = {}
    col_balance = next((c for c in df_balance.columns if 'Saldo Final' in c or 'Balance' in c),
                       df_balance.columns[-1])

    df_pl = cargar_pl(RUTA_RUBROS)
    if df_pl is None:
        print("\n⚠️ No se pudo cargar el diccionario PL")
        return {}

    df_balance['Cuenta_Merge'] = pd.to_numeric(df_balance['Account_Num'], errors='coerce')
    df_balance = df_balance.merge(df_pl, left_on='Cuenta_Merge', right_on='Cuenta', how='left')

    df_net = df_balance[
        (df_balance['PL_Totalizador'] == 'Net Revenues') &
        (~df_balance['Account_Num'].between(49000, 49999))
    ].copy()

    if df_net.empty:
        print("\n⚠️ No se encontraron registros de Net Revenues")
        return {}

    df_net['Product_Code'] = df_net['Product_Flex'].apply(
        lambda x: int(x) if pd.notna(x) and str(x).replace('.','').replace('-','').isdigit() else None
    )
    df_net = df_net.merge(df_glosario, left_on='Product_Code', right_on='COD', how='left')

    # ── REV1: Apertura ONA ────────────────────────────────────────────────────
    entidades_rev1 = ['101', '103', '104', '105', '108', '113', '115', '116', '601']
    df_k1 = df_net[
        (df_net['Entity_Flex'].isin(entidades_rev1)) &
        (df_net['CATEGORIA_APERTURA_ONA'].notna()) &
        (df_net['CATEGORIA_APERTURA_ONA'] != '-') &
        (df_net['CATEGORIA_APERTURA_ONA'] != 'No se considera para el cálculo')
    ].copy()

    if not df_k1.empty:
        df_k1['Abs_Balance'] = df_k1[col_balance] * -1
        pivot_k1 = df_k1.pivot_table(
            index='CATEGORIA_APERTURA_ONA', columns='Entity_Flex',
            values='Abs_Balance', aggfunc='sum', fill_value=0
        )
        pivot_k1_pct = pivot_k1.div(pivot_k1.sum(axis=0), axis=1) * 100

        for col in pivot_k1_pct.columns:
            negativos = pivot_k1_pct[col] < 0
            if negativos.any():
                producto_mayor = pivot_k1_pct[col].idxmax()
                suma_negativos = pivot_k1_pct[col][negativos].sum()
                pivot_k1_pct.loc[producto_mayor, col] -= suma_negativos
                pivot_k1_pct.loc[negativos, col] = 0
            total_col = pivot_k1_pct[col].sum()
            exceso = total_col - 100.0
            if abs(exceso) > 0.0001:
                producto_mayor = pivot_k1_pct[col].idxmax()
                pivot_k1_pct.loc[producto_mayor, col] -= exceso

        pivot_k1_pct = pivot_k1_pct.replace(-0.0, 0.0)
        orden_categorias = ['Paquetes', 'Autos', 'Seguros', 'Cruceros', 'Servicios en Destino']
        categorias_presentes = [c for c in orden_categorias if c in pivot_k1_pct.index]
        pivot_k1_pct = pivot_k1_pct.reindex(categorias_presentes)

        llaves['REV1'] = pivot_k1_pct
        _header("📌 LLAVE REV1: APERTURA ONA")
        imprimir_pivot(llaves['REV1'])
    else:
        llaves['REV1'] = None
        _header("📌 LLAVE REV1: APERTURA ONA")
        print("  ⚠️  Sin datos")

    # ── REV2: Mexico (103 vs 401) ─────────────────────────────────────────────
    df_mx = df_net[df_net['Entity_Flex'].isin(['103', '401'])].copy()
    if not df_mx.empty:
        df_mx['Val'] = df_mx[col_balance] * -1
        total_mx = df_mx['Val'].sum()
        if total_mx != 0:
            llaves['REV2'] = df_mx.groupby('Entity_Flex')['Val'].sum() / total_mx * 100
            _header("📌 LLAVE REV2: MEXICO 103 vs 401")
            imprimir_serie(llaves['REV2'])
        else:
            llaves['REV2'] = None
            _header("📌 LLAVE REV2: MEXICO 103 vs 401")
            print("  ⚠️  Total cero, sin datos")
    else:
        llaves['REV2'] = None
        _header("📌 LLAVE REV2: MEXICO 103 vs 401")
        print("  ⚠️  Sin datos")

    # ── REV3: Viajes Beda (401) ───────────────────────────────────────────────
    df_401 = df_net[df_net['Entity_Flex'] == '401'].copy()
    if not df_401.empty:
        df_401['Val'] = df_401[col_balance] * -1
        total_401 = df_401['Val'].sum()
        if total_401 != 0:
            llaves['REV3'] = df_401.groupby('CATEGORIA_APERTURA_ONA')['Val'].sum() / total_401 * 100
            _header("📌 LLAVE REV3: VIAJES BEDA (401) POR PRODUCTO")
            imprimir_serie(llaves['REV3'], col_label="Categoría ONA")
        else:
            llaves['REV3'] = None
            _header("📌 LLAVE REV3: VIAJES BEDA (401) POR PRODUCTO")
            print("  ⚠️  Total cero, sin datos")
    else:
        llaves['REV3'] = None
        _header("📌 LLAVE REV3: VIAJES BEDA (401) POR PRODUCTO")
        print("  ⚠️  Sin datos de entidad 401")

    return llaves


# ============================================================================
# CÁLCULO ECUADOR - HOJA LLAVE
# ============================================================================

def calcular_ecuador_hoja_llave(df_balance_raw, llaves_ordenes, mes_base=10):
    """
    Calcula los 3 valores de Hoja Llave Ecuador (POS 113) en USD:
      - Licencia de Marca  (TASA_MARCA=2.5%, catch-up YTD)
      - Licencia de IT     (TASA_IT=5.0%,   catch-up YTD)
      - Referencia de Clientes - RFC (garantía MARKUP_RFC=9.5% segmento INT)

    Metodología confirmada desde el Excel de referencia Ecuador:
      G31 = HT INT income (PL=Net Revenues, no interco, pos9 in {0002,0112}) + 49102 HT
      H31 = ONA INT income (PL=Net Revenues, no interco) + 49102 ONA
      G33 = costos HT (rubros costo, no interco, excl cuentas) × llave_HI + M5 × llave_HT × llave_INT
      H33 = costos ONA × llave_OI + M5 × llave_ONA × llave_INT
      M5  = base corp = costos CORPORATE (rubros costo, no interco, incluye 64201/64202)
      RFC = (G31 - G33 × 1.095) + (H31 - H33 × 1.095)
      E31 = HT income DOM + N/A (pos9 not in {0002,0112})
      F31 = ONA income DOM+N/A + TODOS los VP (código 102 → todo DOM, Ingresos rule)
      D31 = Vuelos income (código 101, todos DOM/INT) + fraudes (53300/53305)
    """

    # PROD_MAP, EXCL_COST_ACCOUNTS y RFC_COST_RUBROS son constantes de módulo
    # (definidas al inicio del archivo para reutilización entre entidades)

    col_bal = next((c for c in df_balance_raw.columns if 'Saldo Final' in c or 'Balance' in c),
                   df_balance_raw.columns[-1])
    col_pl2 = next((c for c in df_balance_raw.columns if 'Nivel 2' in c), None)

    # ── Ecuador entity ─────────────────────────────────────────────────────────
    df_ec = df_balance_raw[df_balance_raw['Entity_Flex'] == '113'].copy()
    if df_ec.empty:
        print("⚠️  Ecuador: sin registros para entidad 113 en el balance")
        return None

    df_ec['_prod_code'] = pd.to_numeric(df_ec['Product_Flex'], errors='coerce')
    df_ec['_prod']      = df_ec['_prod_code'].map(
        lambda x: PROD_MAP.get(int(x), 'OTRO') if pd.notna(x) else 'OTRO'
    )
    # Clasificación DOM/INT por posición 9 del flex field de Oracle:
    #   '0001' = DOM (doméstico estándar)
    #   '0002' = INT (internacional estándar)
    #   '0112' = INT (valor no estándar observado en HT Ecuador nov-2025;
    #                 confirmado como INT por la columna AF de la hoja 113 del Excel.
    #                 Se mantiene por las dudas — si no aparece en el balance, no afecta nada)
    #   cualquier otro = N/A
    df_ec['_di'] = df_ec['DomInt_Flex'].map(
        lambda x: 'INT' if x in ('0002', '0112') else ('DOM' if x == '0001' else 'N/A')
    )

    is_ic = (df_ec[col_pl2] == 'Intercompany Transactions') if col_pl2 \
            else pd.Series(False, index=df_ec.index)
    df_ni = df_ec[~is_ic]  # non-interco rows

    # ── Llaves de órdenes ─────────────────────────────────────────────────────
    def _llave(key, row, col):
        try:
            return float(llaves_ordenes[key].loc[row, col]) / 100
        except Exception:
            return 0.0

    llave_HI      = _llave('K11', 113, 'Int')
    llave_OI      = _llave('K12', 113, 'Int')
    llave_INT     = _llave('K9',  113, 'Int')
    llave_HT_pct  = _llave('K10', 113, 'Hoteles')
    llave_ONA_pct = _llave('K10', 113, 'ONA')

    # ── Cargar PL dict (una sola vez, para ingresos Y costos) ─────────────────
    df_pl_local = cargar_pl(RUTA_RUBROS)

    # Merge df_ec con PL dict para clasificar por RUBRO (PL_Totalizador)
    if df_pl_local is not None:
        df_ec_pl = df_ec.copy()
        df_ec_pl['_is_ic'] = is_ic.values
        df_ec_pl = df_ec_pl.merge(df_pl_local, left_on='Account_Num', right_on='Cuenta', how='left')

        # ── INGRESOS: RUBRO = Net Revenues, no interco ────────────────────────
        # Replica exactamente el filtro RUBRO='NET REVENUE' de la hoja 113 del Excel
        df_net_ni = df_ec_pl[
            (df_ec_pl['PL_Totalizador'] == 'Net Revenues') &
            (~df_ec_pl['_is_ic'])
        ]
    else:
        # Fallback: aproximación por rango de cuentas (menos preciso)
        df_ec_pl = None
        df_net_ni = df_ni[df_ni['Account_Num'].between(40000, 49999)]

    # ── INGRESOS ──────────────────────────────────────────────────────────────

    # G31 / H31: ingresos segmento INT (Net Revenue, no interco) + 49102 ya facturado
    # Formula Segmentación: G31 = HT_INT_income + 49102_HT
    # (el 49102 es el RFC ya cobrado/pagado, que se incluye en la base de ingresos)
    G31 = -(df_net_ni[(df_net_ni['_prod'] == 'HT')  & (df_net_ni['_di'] == 'INT')][col_bal].sum()) \
          -(df_ec[(df_ec['Account_Num'] == 49102) & (df_ec['_prod'] == 'HT') ][col_bal].sum())

    H31 = -(df_net_ni[(df_net_ni['_prod'] == 'ONA') & (df_net_ni['_di'] == 'INT')][col_bal].sum()) \
          -(df_ec[(df_ec['Account_Num'] == 49102) & (df_ec['_prod'] == 'ONA')][col_bal].sum())

    # D31: todos los ingresos VUELOS (Net Revenue, todos DOM/INT/N/A) + fraudes totales entity 113
    # Nota: Ingresos!C20 = VUELOS_all + Grand_Total_fraudes(53300/53305 todos los productos)
    D31 = -(df_net_ni[df_net_ni['_prod'] == 'VUELOS'][col_bal].sum()) \
          -(df_ec[df_ec['Account_Num'].isin([53300, 53305])][col_bal].sum())

    # E31: HT Net Revenue DOM + N/A
    # Formula Ingresos!C5 = GETPIVOTDATA(HT,DOM) + GETPIVOTDATA(HT,N/A)
    E31 = -(df_net_ni[(df_net_ni['_prod'] == 'HT') & (df_net_ni['_di'].isin(['DOM', 'N/A']))][col_bal].sum())

    # F31: ONA Net Revenue DOM + N/A + TODOS los VUELOS PAQUETES (VP, código 102)
    # Regla Ingresos: "Vuelos y Vuelos paquetes se imputa todo a doméstico independientemente del flex"
    # Formula Ingresos: C6(ONA DOM+N/A) + C8(VP todos → DOM, independientemente de pos9)
    F31 = -(df_net_ni[(df_net_ni['_prod'] == 'ONA') & (df_net_ni['_di'].isin(['DOM', 'N/A']))][col_bal].sum()) \
          -(df_net_ni[df_net_ni['_prod'] == 'VUELOS PAQUETES'][col_bal].sum())

    # ── COSTOS (base para RFC) ────────────────────────────────────────────────
    # Usa RFC_COST_RUBROS definido a nivel de módulo
    if df_ec_pl is not None:
        # Costos directos HT/ONA: rubros aprobados, no interco, no cuentas excluidas
        df_costs_base = df_ec_pl[
            (df_ec_pl['PL_Totalizador'].isin(RFC_COST_RUBROS)) &
            (~df_ec_pl['_is_ic']) &
            (~df_ec_pl['Account_Num'].isin(EXCL_COST_ACCOUNTS))
        ]
        E18_total = df_costs_base[df_costs_base['_prod'] == 'HT' ][col_bal].sum()
        F18_total = df_costs_base[df_costs_base['_prod'] == 'ONA'][col_bal].sum()
    else:
        # Fallback: cuentas hardcodeadas
        CUENTAS_COGS = {41305,41307,41309,41310,53510,54051,62083,65020,65060,68600,
                        80102,80103,53513,80107,68150,41313}
        CUENTAS_SM   = {53710, 64351}
        E18_total = df_ec[(df_ec['Account_Num'].isin(CUENTAS_COGS)) & (df_ec['_prod']=='HT')][col_bal].sum() \
                  + df_ec[(df_ec['Account_Num'].isin(CUENTAS_SM))   & (df_ec['_prod']=='HT')][col_bal].sum()
        F18_total = df_ec[(df_ec['Account_Num'].isin(CUENTAS_COGS)) & (df_ec['_prod']=='ONA')][col_bal].sum() \
                  + df_ec[(df_ec['Account_Num'].isin(CUENTAS_SM))   & (df_ec['_prod']=='ONA')][col_bal].sum()

    # M5: base corporativa = costos CORPORATE (rubros de costo, no interco)
    # IMPORTANTE: NO excluir 64201/64202 (Marketing) para la base corporativa
    # (difiere de E18/F18 donde sí se excluyen — aquí seguimos la hoja 113 del Excel)
    if df_ec_pl is not None:
        df_corp_m5 = df_ec_pl[
            (df_ec_pl['_prod'] == 'CORPORATE') &
            (df_ec_pl['PL_Totalizador'].isin(RFC_COST_RUBROS)) &
            (~df_ec_pl['_is_ic'])
        ]
        M5 = df_corp_m5[col_bal].sum()
    else:
        M5 = df_ec[df_ec['_prod'] == 'CORPORATE'][col_bal].sum()
    E22 = M5 * llave_HT_pct
    F22 = M5 * llave_ONA_pct

    G33 = E18_total * llave_HI + E22 * llave_INT
    H33 = F18_total * llave_OI + F22 * llave_INT

    # ── RFC (Referencia de Clientes) ──────────────────────────────────────────
    # RFC = (G31 - G33×MARKUP) + (H31 - H33×MARKUP)
    # Positivo = Ecuador debe recibir; Negativo = Ecuador cobra menos de lo facturado
    RFC = (G31 - G33 * MARKUP_RFC) + (H31 - H33 * MARKUP_RFC)

    # ── LICENCIAS — catch-up YTD + accrual mensual ────────────────────────────
    # Fórmula: (Teorico_YTD × TASA - Ya_contabilizado) + Teorico_YTD × TASA / mes_base
    D65 = df_ec[(df_ec['Account_Num'] == 70104) & (df_ec['_prod'] == 'VUELOS')][col_bal].sum()
    E65 = df_ec[(df_ec['Account_Num'] == 70104) & (df_ec['_prod'] == 'HT')   ][col_bal].sum()
    F65 = df_ec[(df_ec['Account_Num'] == 70104) & (df_ec['_prod'] == 'ONA')  ][col_bal].sum()

    D70 = df_ec[(df_ec['Account_Num'] == 70105) & (df_ec['_prod'] == 'VUELOS')][col_bal].sum()
    E70 = df_ec[(df_ec['Account_Num'] == 70105) & (df_ec['_prod'] == 'HT')   ][col_bal].sum()
    F70 = df_ec[(df_ec['Account_Num'] == 70105) & (df_ec['_prod'] == 'ONA')  ][col_bal].sum()

    D61 = D31 * TASA_MARCA;  E61 = E31 * TASA_MARCA;  F61 = F31 * TASA_MARCA
    D62 = D31 * TASA_IT;     E62 = E31 * TASA_IT;     F62 = F31 * TASA_IT

    D67 = (D61 - D65) + D61 / mes_base
    E67 = (E61 - E65) + E61 / mes_base
    F67 = (F61 - F65) + F61 / mes_base

    D72 = (D62 - D70) + D62 / mes_base
    E72 = (E62 - E70) + E62 / mes_base
    F72 = (F62 - F70) + F62 / mes_base

    Lic_Marca = D67 + E67 + F67
    Lic_IT    = D72 + E72 + F72

    return {
        'Licencia_Marca': Lic_Marca,
        'Licencia_IT':    Lic_IT,
        'RFC':            RFC,
        # Intermedios de diagnóstico
        '_G31': G31, '_H31': H31, '_D31': D31, '_E31': E31, '_F31': F31,
        '_G33': G33, '_H33': H33,
        '_E18_total': E18_total, '_F18_total': F18_total,
        '_E22': E22, '_F22': F22,
        '_M5': M5,
        '_llave_HI': llave_HI, '_llave_OI': llave_OI,
        '_llave_INT': llave_INT, '_llave_HT': llave_HT_pct, '_llave_ONA': llave_ONA_pct,
        '_mes_base': mes_base,
        '_D65': D65, '_E65': E65, '_F65': F65,
        '_D67': D67, '_E67': E67, '_F67': F67,
    }


# ============================================================================
# CÁLCULO CHILE - HOJA LLAVE
# ============================================================================

def calcular_chile_hoja_llave(df_balance_raw, llaves_ordenes, tc_clp, mes_base=10):
    """
    Calcula los 3 valores de Hoja Llave Chile (entidad 108) en USD:
      - Licencia de Marca  (TASA_MARCA=2.5%, catch-up YTD)
      - Licencia de IT     (TASA_IT=5.0%,   catch-up YTD)
      - Referencia de Clientes - RFC (garantía MARKUP_RFC=9.5% segmento INT)

    Nota: Chile NO factura Hosting (Costo Servicios=0 en Ref Cte).

    Metodología idéntica a Ecuador pero con las siguientes diferencias:
      - Entity_Flex = '108'
      - Moneda funcional: CLP → se divide por tc_clp al retornar USD
      - M5_rfc = M5 (sin deducción de Hosting — Chile no tiene ese flujo)
      - Llaves K9/K10/K11/K12 para entidad 108

    Targets validados Nov-2025 (Despegar Chile - Segmentación Nov-2025.xlsx):
      Licencia Marca = $137,189.77 USD
      Licencia IT    = $274,382.09 USD
      RFC            = -$2,847,997.29 USD
    """

    MARKUP_RFC_CHILE = 1 + ENTITY_CONFIG[108]['markup_rfc']   # 1.095

    col_bal = next((c for c in df_balance_raw.columns if 'Saldo Final' in c or 'Balance' in c),
                   df_balance_raw.columns[-1])
    col_pl2 = next((c for c in df_balance_raw.columns if 'Nivel 2' in c), None)

    # ── Chile entity ────────────────────────────────────────────────────────
    df_cl = df_balance_raw[df_balance_raw['Entity_Flex'] == '108'].copy()
    if df_cl.empty:
        print("⚠️  Chile: sin registros para entidad 108 en el balance")
        return None

    df_cl['_prod_code'] = pd.to_numeric(df_cl['Product_Flex'], errors='coerce')
    df_cl['_prod']      = df_cl['_prod_code'].map(
        lambda x: PROD_MAP.get(int(x), 'OTRO') if pd.notna(x) else 'OTRO'
    )
    # Clasificación DOM/INT — Chile: '0102' se clasifica como INT (igual que '0002'/'0112')
    # Validado: ONA INT con '0102' = 5,588,414 CLP adicionales necesarios para coincidir
    # con el pivote B35 Ingresos (D6_ONA_INT = 6,456,796,328 CLP)
    df_cl['_di'] = df_cl['DomInt_Flex'].map(
        lambda x: 'INT' if x in ('0002', '0112', '0102') else ('DOM' if x == '0001' else 'N/A')
    )

    is_ic = (df_cl[col_pl2] == 'Intercompany Transactions') if col_pl2 \
            else pd.Series(False, index=df_cl.index)
    df_ni = df_cl[~is_ic]   # non-interco rows

    # ── Llaves de órdenes (entidad 108) ────────────────────────────────────
    def _llave(key, row, col):
        try:
            return float(llaves_ordenes[key].loc[row, col]) / 100
        except Exception:
            return 0.0

    llave_HI      = _llave('K11', 108, 'Int')
    llave_OI      = _llave('K12', 108, 'Int')
    llave_INT     = _llave('K9',  108, 'Int')
    llave_HT_pct  = _llave('K10', 108, 'Hoteles')
    llave_ONA_pct = _llave('K10', 108, 'ONA')

    # ── Cargar PL dict para clasificación por rubro ─────────────────────────
    df_pl_local = cargar_pl(RUTA_RUBROS)

    if df_pl_local is not None:
        df_cl_pl = df_cl.copy()
        df_cl_pl['_is_ic'] = is_ic.values
        df_cl_pl = df_cl_pl.merge(df_pl_local, left_on='Account_Num', right_on='Cuenta', how='left')

        df_net_ni = df_cl_pl[
            (df_cl_pl['PL_Totalizador'] == 'Net Revenues') &
            (~df_cl_pl['_is_ic'])
        ]
    else:
        df_cl_pl  = None
        df_net_ni = df_ni[df_ni['Account_Num'].between(40000, 49999)]

    # ── FRAUDES por producto y segmento (cuentas 53300/53305, no interco) ────
    # Los fraudes aparecen en B51 pivot de Ingresos y se suman al ingreso base
    # de la Segmentación (Ingresos!C18/D18/etc.) para ajustar la base de Licencias y RFC.
    def _fraudes(prod_list, di_list=None):
        mask = df_cl['Account_Num'].isin([53300, 53305]) & df_cl['_prod'].isin(prod_list)
        if di_list is not None:
            mask &= df_cl['_di'].isin(di_list)
        return df_cl[mask][col_bal].sum()

    fraudes_HT_INT    = _fraudes(['HT'],              ['INT'])           # → ajusta G31 (RFC)
    fraudes_ONA_INT   = _fraudes(['ONA'],             ['INT'])           # → ajusta H31 (RFC)
    fraudes_VUELOS    = _fraudes(['VUELOS'])                             # → ajusta D31 (Licencias)
    fraudes_HT_DOM    = _fraudes(['HT'],              ['DOM', 'N/A'])    # → ajusta E31 (Licencias)
    fraudes_ONA_DOM   = _fraudes(['ONA'],             ['DOM', 'N/A'])    # → ajusta F31 (Licencias)
    fraudes_VP        = _fraudes(['VUELOS PAQUETES'])                    # → ajusta F31 (Licencias)

    # ── INGRESOS RFC (G31/H31): HT/ONA INT + 49102 ya facturado + fraudes INT ──
    # Fórmula Excel: F31_Seg = -Ingresos!D18 - D9  (HT INT con fraudes + 49102)
    #                G31_Seg = -Ingresos!D19 - E9  (ONA INT con fraudes + 49102)
    G31 = -(df_net_ni[(df_net_ni['_prod'] == 'HT')  & (df_net_ni['_di'] == 'INT')][col_bal].sum()) \
          -(df_cl[(df_cl['Account_Num'] == 49102) & (df_cl['_prod'] == 'HT') ][col_bal].sum()) \
          - fraudes_HT_INT

    H31 = -(df_net_ni[(df_net_ni['_prod'] == 'ONA') & (df_net_ni['_di'] == 'INT')][col_bal].sum()) \
          -(df_cl[(df_cl['Account_Num'] == 49102) & (df_cl['_prod'] == 'ONA')][col_bal].sum()) \
          - fraudes_ONA_INT

    # ── INGRESOS LICENCIAS (D31/E31/F31): base local (DOM+N/A) + fraudes locales ──
    # Fórmula Excel: C31_Seg = -Ingresos!E20        (VUELOS all-DI + fraudes_VUELOS)
    #                D31_Seg = -Ingresos!C18         (HT DOM+N/A + fraudes_HT_DOM)
    #                E31_Seg = -Ingresos!C19 - C21   (ONA DOM+N/A + VP all-DI + fraudes resp.)
    D31 = -(df_net_ni[df_net_ni['_prod'] == 'VUELOS'][col_bal].sum()) \
          - fraudes_VUELOS

    E31 = -(df_net_ni[(df_net_ni['_prod'] == 'HT') & (df_net_ni['_di'].isin(['DOM', 'N/A']))][col_bal].sum()) \
          - fraudes_HT_DOM

    F31 = -(df_net_ni[(df_net_ni['_prod'] == 'ONA') & (df_net_ni['_di'].isin(['DOM', 'N/A']))][col_bal].sum()) \
          - fraudes_ONA_DOM \
          -(df_net_ni[df_net_ni['_prod'] == 'VUELOS PAQUETES'][col_bal].sum()) \
          - fraudes_VP

    # ── COSTOS (base para RFC) ───────────────────────────────────────────────
    if df_cl_pl is not None:
        df_costs_base = df_cl_pl[
            (df_cl_pl['PL_Totalizador'].isin(RFC_COST_RUBROS)) &
            (~df_cl_pl['_is_ic']) &
            (~df_cl_pl['Account_Num'].isin(EXCL_COST_ACCOUNTS))
        ]
        E18_total = df_costs_base[df_costs_base['_prod'] == 'HT' ][col_bal].sum()
        F18_total = df_costs_base[df_costs_base['_prod'] == 'ONA'][col_bal].sum()
    else:
        E18_total = 0.0
        F18_total = 0.0

    # M5: base corporativa CORPORATE (rubros de costo, no interco)
    # Chile no deduce Hosting (no hay facturación de Costo Servicios → M5_rfc = M5)
    if df_cl_pl is not None:
        df_corp_m5 = df_cl_pl[
            (df_cl_pl['_prod'] == 'CORPORATE') &
            (df_cl_pl['PL_Totalizador'].isin(RFC_COST_RUBROS)) &
            (~df_cl_pl['_is_ic'])
        ]
        M5 = df_corp_m5[col_bal].sum()
    else:
        M5 = df_cl[df_cl['_prod'] == 'CORPORATE'][col_bal].sum()

    M5_rfc = M5   # No se deduce Hosting en Chile

    E22 = M5_rfc * llave_HT_pct
    F22 = M5_rfc * llave_ONA_pct

    G33 = E18_total * llave_HI + E22 * llave_INT
    H33 = F18_total * llave_OI + F22 * llave_INT

    # ── RFC (en CLP, luego convertir a USD) ─────────────────────────────────
    RFC_CLP = (G31 - G33 * MARKUP_RFC_CHILE) + (H31 - H33 * MARKUP_RFC_CHILE)
    RFC     = RFC_CLP / tc_clp

    # ── LICENCIAS — catch-up YTD + accrual mensual (en CLP, luego USD) ──────
    # Booked Marca: 70104 por producto. ONA incluye VP (E69 = E11 + F11 en Segmentación)
    D65 = df_cl[(df_cl['Account_Num'] == 70104) & (df_cl['_prod'] == 'VUELOS')][col_bal].sum()
    E65 = df_cl[(df_cl['Account_Num'] == 70104) & (df_cl['_prod'] == 'HT')   ][col_bal].sum()
    F65 = df_cl[(df_cl['Account_Num'] == 70104) &
                 df_cl['_prod'].isin(['ONA', 'VUELOS PAQUETES'])][col_bal].sum()

    # Booked IT: 70105 por producto. ONA incluye VP (misma lógica que Marca)
    D70 = df_cl[(df_cl['Account_Num'] == 70105) & (df_cl['_prod'] == 'VUELOS')][col_bal].sum()
    E70 = df_cl[(df_cl['Account_Num'] == 70105) & (df_cl['_prod'] == 'HT')   ][col_bal].sum()
    F70 = df_cl[(df_cl['Account_Num'] == 70105) &
                 df_cl['_prod'].isin(['ONA', 'VUELOS PAQUETES'])][col_bal].sum()

    D61 = D31 * TASA_MARCA;  E61 = E31 * TASA_MARCA;  F61 = F31 * TASA_MARCA
    D62 = D31 * TASA_IT;     E62 = E31 * TASA_IT;     F62 = F31 * TASA_IT

    D67 = (D61 - D65) + D61 / mes_base
    E67 = (E61 - E65) + E61 / mes_base
    F67 = (F61 - F65) + F61 / mes_base

    D72 = (D62 - D70) + D62 / mes_base
    E72 = (E62 - E70) + E62 / mes_base
    F72 = (F62 - F70) + F62 / mes_base

    Lic_Marca_CLP = D67 + E67 + F67
    Lic_IT_CLP    = D72 + E72 + F72

    Lic_Marca = Lic_Marca_CLP / tc_clp
    Lic_IT    = Lic_IT_CLP    / tc_clp

    return {
        'Licencia_Marca': Lic_Marca,
        'Licencia_IT':    Lic_IT,
        'RFC':            RFC,
        # Intermedios de diagnóstico
        '_G31': G31, '_H31': H31, '_D31': D31, '_E31': E31, '_F31': F31,
        '_G33': G33, '_H33': H33,
        '_E18_total': E18_total, '_F18_total': F18_total,
        '_E22': E22, '_F22': F22,
        '_M5': M5, '_M5_rfc': M5_rfc,
        '_RFC_CLP': RFC_CLP,
        '_Lic_Marca_CLP': Lic_Marca_CLP, '_Lic_IT_CLP': Lic_IT_CLP,
        '_llave_HI': llave_HI, '_llave_OI': llave_OI,
        '_llave_INT': llave_INT, '_llave_HT': llave_HT_pct, '_llave_ONA': llave_ONA_pct,
        '_tc_clp': tc_clp, '_mes_base': mes_base,
        '_fraudes_HT_INT': fraudes_HT_INT, '_fraudes_ONA_INT': fraudes_ONA_INT,
        '_fraudes_VUELOS': fraudes_VUELOS, '_fraudes_HT_DOM': fraudes_HT_DOM,
        '_fraudes_ONA_DOM': fraudes_ONA_DOM, '_fraudes_VP': fraudes_VP,
        '_D65': D65, '_E65': E65, '_F65': F65,
        '_D70': D70, '_E70': E70, '_F70': F70,
    }


# ============================================================================
# CÁLCULO PERÚ - HOJA LLAVE
# ============================================================================

def calcular_peru_hoja_llave(df_balance_raw, llaves_ordenes, tc_pen, mes_base=10):
    """
    Calcula los 3 valores de Hoja Llave Perú (POS 115) en USD:
      - Licencia de Marca  (TASA_MARCA=2.5%, catch-up YTD)
      - Licencia de IT     (TASA_IT=5.0%,   catch-up YTD)
      - Referencia de Clientes - RFC (garantía MARKUP_RFC=9.5% segmento INT)

    Particularidades Perú vs Ecuador:
      - Moneda funcional PEN → todos los intermedios en PEN, dividir por tc_pen para USD
      - DomInt_Flex adicionales: '0101'→DOM, '0111'→DOM  (además del estándar '0001')
                                 '0102'→INT               (además de '0002' y '0112')
      - Sin Hosting (no aplica para Perú)
      - Booked amounts F65/F70: solo ONA pura (sin VP — igual que Ecuador)
    """
    col_bal = next((c for c in df_balance_raw.columns if 'Saldo Final' in c or 'Balance' in c),
                   df_balance_raw.columns[-1])
    col_pl2 = next((c for c in df_balance_raw.columns if 'Nivel 2' in c), None)

    df_pe = df_balance_raw[df_balance_raw['Entity_Flex'] == '115'].copy()
    if df_pe.empty:
        print("⚠️  Perú: sin registros para entidad 115 en el balance")
        return None

    df_pe['_prod_code'] = pd.to_numeric(df_pe['Product_Flex'], errors='coerce')
    df_pe['_prod'] = df_pe['_prod_code'].map(
        lambda x: PROD_MAP.get(int(x), 'OTRO') if pd.notna(x) else 'OTRO'
    )
    # DOM/INT para Perú:
    #   '0001','0101','0111' = DOM  (0101 y 0111 exclusivos Perú — validados nov-2025)
    #   '0002','0102','0112' = INT
    #   resto                = N/A
    df_pe['_di'] = df_pe['DomInt_Flex'].map(
        lambda x: 'INT' if x in ('0002', '0102', '0112') else
                  ('DOM' if x in ('0001', '0101', '0111') else 'N/A')
    )

    is_ic = (df_pe[col_pl2] == 'Intercompany Transactions') if col_pl2 \
            else pd.Series(False, index=df_pe.index)

    # ── Llaves de órdenes ─────────────────────────────────────────────────────
    def _llave(key, row, col):
        try:
            return float(llaves_ordenes[key].loc[row, col]) / 100
        except Exception:
            return 0.0

    llave_HI      = _llave('K11', 115, 'Int')
    llave_OI      = _llave('K12', 115, 'Int')
    llave_INT     = _llave('K9',  115, 'Int')
    llave_HT_pct  = _llave('K10', 115, 'Hoteles')
    llave_ONA_pct = _llave('K10', 115, 'ONA')

    # ── Cargar PL dict ────────────────────────────────────────────────────────
    df_pl_local = cargar_pl(RUTA_RUBROS)

    if df_pl_local is not None:
        df_pe_pl = df_pe.copy()
        df_pe_pl['_is_ic'] = is_ic.values
        df_pe_pl = df_pe_pl.merge(df_pl_local, left_on='Account_Num', right_on='Cuenta', how='left')

        df_net_ni = df_pe_pl[
            (df_pe_pl['PL_Totalizador'] == 'Net Revenues') &
            (~df_pe_pl['_is_ic'])
        ]
    else:
        df_pe_pl = None
        df_net_ni = df_pe[~is_ic & df_pe['Account_Num'].between(40000, 49999)]

    # ── Fraudes (53300/53305) — split por producto/segmento ──────────────────
    def _fraudes(prod=None, di=None):
        mask = df_pe['Account_Num'].isin([53300, 53305])
        if prod is not None:
            mask &= (df_pe['_prod'] == prod) if isinstance(prod, str) \
                    else df_pe['_prod'].isin(prod)
        if di is not None:
            mask &= (df_pe['_di'] == di) if isinstance(di, str) \
                    else df_pe['_di'].isin(di)
        return df_pe[mask][col_bal].sum()

    fraudes_VUELOS  = _fraudes(prod='VUELOS')
    fraudes_HT_DOM  = _fraudes(prod='HT',  di=['DOM', 'N/A'])
    fraudes_ONA_DOM = _fraudes(prod='ONA', di=['DOM', 'N/A'])
    fraudes_VP      = _fraudes(prod='VUELOS PAQUETES')
    fraudes_HT_INT  = _fraudes(prod='HT',  di='INT')
    fraudes_ONA_INT = _fraudes(prod='ONA', di='INT')

    # ── INGRESOS (en PEN) ─────────────────────────────────────────────────────
    # D31: VUELOS all DI (non-IC) − fraudes Vuelos
    D31 = -(df_net_ni[df_net_ni['_prod'] == 'VUELOS'][col_bal].sum()) \
          - fraudes_VUELOS

    # E31: HT DOM+N/A (non-IC) − fraudes HT DOM+N/A
    E31 = -(df_net_ni[(df_net_ni['_prod'] == 'HT') &
                      (df_net_ni['_di'].isin(['DOM', 'N/A']))][col_bal].sum()) \
          - fraudes_HT_DOM

    # F31: ONA DOM+N/A (non-IC) + VP all − fraudes ONA+VP DOM
    F31 = -(df_net_ni[(df_net_ni['_prod'] == 'ONA') &
                      (df_net_ni['_di'].isin(['DOM', 'N/A']))][col_bal].sum()) \
          -(df_net_ni[df_net_ni['_prod'] == 'VUELOS PAQUETES'][col_bal].sum()) \
          - fraudes_ONA_DOM - fraudes_VP

    # G31/H31: ingresos INT + 49102 ya facturado − fraudes INT
    _49102_HT  = df_pe[(df_pe['Account_Num'] == 49102) & (df_pe['_prod'] == 'HT')][col_bal].sum()
    _49102_ONA = df_pe[(df_pe['Account_Num'] == 49102) & (df_pe['_prod'] == 'ONA')][col_bal].sum()
    G31 = -(df_net_ni[(df_net_ni['_prod'] == 'HT')  & (df_net_ni['_di'] == 'INT')][col_bal].sum()) \
          - _49102_HT - fraudes_HT_INT
    H31 = -(df_net_ni[(df_net_ni['_prod'] == 'ONA') & (df_net_ni['_di'] == 'INT')][col_bal].sum()) \
          - _49102_ONA - fraudes_ONA_INT

    # ── COSTOS (en PEN) ───────────────────────────────────────────────────────
    MARKUP_RFC_PE = 1 + ENTITY_CONFIG[115]['markup_rfc']

    if df_pe_pl is not None:
        df_costs_base = df_pe_pl[
            (df_pe_pl['PL_Totalizador'].isin(RFC_COST_RUBROS)) &
            (~df_pe_pl['_is_ic']) &
            (~df_pe_pl['Account_Num'].isin(EXCL_COST_ACCOUNTS))
        ]
        E18_total = df_costs_base[df_costs_base['_prod'] == 'HT'][col_bal].sum()
        F18_total = df_costs_base[df_costs_base['_prod'] == 'ONA'][col_bal].sum()

        M5 = df_pe_pl[
            (df_pe_pl['_prod'] == 'CORPORATE') &
            (df_pe_pl['PL_Totalizador'].isin(RFC_COST_RUBROS)) &
            (~df_pe_pl['_is_ic'])
        ][col_bal].sum()
    else:
        E18_total = 0.0
        F18_total = 0.0
        M5        = 0.0

    E22 = M5 * llave_HT_pct
    F22 = M5 * llave_ONA_pct
    G33 = E18_total * llave_HI + E22 * llave_INT
    H33 = F18_total * llave_OI + F22 * llave_INT

    # ── RFC ───────────────────────────────────────────────────────────────────
    RFC_PEN = (G31 - G33 * MARKUP_RFC_PE) + (H31 - H33 * MARKUP_RFC_PE)
    RFC     = RFC_PEN / tc_pen

    # ── LICENCIAS — catch-up YTD + accrual mensual (en PEN → USD) ────────────
    D65 = df_pe[(df_pe['Account_Num'] == 70104) & (df_pe['_prod'] == 'VUELOS')][col_bal].sum()
    E65 = df_pe[(df_pe['Account_Num'] == 70104) & (df_pe['_prod'] == 'HT')   ][col_bal].sum()
    F65 = df_pe[(df_pe['Account_Num'] == 70104) & (df_pe['_prod'] == 'ONA')  ][col_bal].sum()
    D70 = df_pe[(df_pe['Account_Num'] == 70105) & (df_pe['_prod'] == 'VUELOS')][col_bal].sum()
    E70 = df_pe[(df_pe['Account_Num'] == 70105) & (df_pe['_prod'] == 'HT')   ][col_bal].sum()
    F70 = df_pe[(df_pe['Account_Num'] == 70105) & (df_pe['_prod'] == 'ONA')  ][col_bal].sum()

    D61 = D31 * TASA_MARCA;  E61 = E31 * TASA_MARCA;  F61 = F31 * TASA_MARCA
    D62 = D31 * TASA_IT;     E62 = E31 * TASA_IT;     F62 = F31 * TASA_IT

    D67 = (D61 - D65) + D61 / mes_base
    E67 = (E61 - E65) + E61 / mes_base
    F67 = (F61 - F65) + F61 / mes_base
    D72 = (D62 - D70) + D62 / mes_base
    E72 = (E62 - E70) + E62 / mes_base
    F72 = (F62 - F70) + F62 / mes_base

    Lic_Marca = (D67 + E67 + F67) / tc_pen
    Lic_IT    = (D72 + E72 + F72) / tc_pen

    return {
        'Licencia_Marca': Lic_Marca,
        'Licencia_IT':    Lic_IT,
        'RFC':            RFC,
        # Intermedios de diagnóstico
        '_G31': G31, '_H31': H31, '_D31': D31, '_E31': E31, '_F31': F31,
        '_G33': G33, '_H33': H33,
        '_E18': E18_total, '_F18': F18_total, '_M5': M5,
        '_E22': E22, '_F22': F22,
        '_llave_HI': llave_HI, '_llave_OI': llave_OI,
        '_llave_INT': llave_INT, '_llave_HT': llave_HT_pct, '_llave_ONA': llave_ONA_pct,
        '_mes_base': mes_base, '_tc_pen': tc_pen,
        '_RFC_PEN': RFC_PEN,
        '_D65': D65, '_E65': E65, '_F65': F65,
        '_D67': D67, '_E67': E67, '_F67': F67,
    }


# ============================================================================
# CÁLCULO USA - HOJA LLAVE
# ============================================================================

def calcular_usa_hoja_llave(df_balance_raw, llaves_ordenes, mes_base=10):
    """
    Calcula los 4 valores de Hoja Llave USA (POS 105) en USD:
      - Licencia de Marca      (TASA_MARCA=2.5%, catch-up YTD)
      - Licencia de IT         (TASA_IT=5.0%,   catch-up YTD)
      - Referencia de Clientes (MARKUP_RFC=9.5%, segmento INT HT+ONA)
      - Servicios de Hosting   (MARKUP_HOSTING=6%, CORPORATE IT, net de 49120)

    Lógica reutilizada de Ecuador:
      - Parsing flex field (Entity_Flex='105', Product_Flex→PROD_MAP, DomInt_Flex→_di)
      - Detección intercompany (Mapeo PL Nivel 2 == 'Intercompany Transactions')
      - Llaves K9/K10/K11/K12 para entidad 105 (Corporate INT%, HT%, HT DOM/INT%, ONA DOM/INT%)
      - Fórmulas catch-up Licencias: (Teórico_YTD×TASA − Contabilizado) + Teórico_YTD×TASA/mes_base
      - Fórmula RFC: (G31 − G33×MARKUP) + (H31 − H33×MARKUP)  ← IDÉNTICA a Ecuador ✅
        Donde:
          G31 = ingresos HT-INT (Net Revenue non-IC) + 49102-HT (ya facturado RFC)
          H31 = ingresos ONA-INT (Net Revenue non-IC) + 49102-ONA (ya facturado RFC)
          G33 = E18×llave_HI + (M5_rfc×llave_HT)×llave_INT   (costos imputados HT-INT)
          H33 = F18×llave_OI + (M5_rfc×llave_ONA)×llave_INT  (costos imputados ONA-INT)
          E18/F18 = costos directos HT/ONA non-IC (excluye incentivos, non-op, EXCL_COST_ACCOUNTS)
          M5_rfc  = CORPORATE CORP_RFC_PL1 non-IC − Costos_Hosting

    Lógica específica USA (diferencias vs Ecuador):
      - Costos_Hosting: filtra por Mapeo PL Nivel 1 == 'IT - Personnel/Expenses' (CORPORATE non-IC)
        No usa PL_Totalizador (cobertura incompleta para cuentas USA)
        ✅ Validado Nov-25: 1,394,173.78
      - M5: filtra por CORP_RFC_PL1_USA (excluye impuestos, CC, marketing, intereses, FX)
        ✅ Validado Nov-25: 2,483,539.18 (ajuste cuentas 80101/80102/64341 vs Estimado!M6)
      - E18/F18: excluye EXCL_PL1_USA = incentivos + non-op (Back End/Other/Up Front Incentives,
        Interest Income, FX/Other)
      - Hosting: (Costos_Hosting × 1.06 − Ya_49120) + Costos_Hosting × 1.06 / mes_base
        ✅ Validado Nov-25: $117,941.08
      - 70104/70105: son INGRESOS (cuenta de crédito = negativo en Oracle) en USA,
        al igual que en Ecuador. NO aplicar negación adicional.

    Papeles de trabajo (solo referencia, no son inputs):
      /Users/josemiguelalvarez/Downloads/TP/Despegar USA - Segmentación Nov-2025.xlsx
        → Hoja Llave B4=RFC(-148,549.21), B2=Marca(3,400.90), B3=IT(6,800.79)
        → Segmentación: confirma G31, H31, G33, H33 y estructura de costos CORPORATE
      /Users/josemiguelalvarez/Downloads/TP/Servicios Hosting Nov-2025.xlsx
        → Confirma Costos_Hosting=1,394,173.78, Ya_49120=1,507,665.55
    """

    MARKUP_RFC_USA    = 1 + ENTITY_CONFIG[105]['markup_rfc']      # 1.095
    MARKUP_HOSTING    = 1 + ENTITY_CONFIG[105]['markup_hosting']   # 1.060

    col_bal = next((c for c in df_balance_raw.columns if 'Saldo Final' in c or 'Balance' in c),
                   df_balance_raw.columns[-1])
    col_pl2 = next((c for c in df_balance_raw.columns if 'Nivel 2' in c), None)

    col_pl1 = next((c for c in df_balance_raw.columns if 'Nivel 1' in c), None)

    # ── Entidad USA (105) ───────────────────────────────────────────────────
    df_usa = df_balance_raw[df_balance_raw['Entity_Flex'] == '105'].copy()
    if df_usa.empty:
        print("⚠️  USA: sin registros para entidad 105 en el balance")
        return None

    df_usa['_prod_code'] = pd.to_numeric(df_usa['Product_Flex'], errors='coerce')
    df_usa['_prod']      = df_usa['_prod_code'].map(
        lambda x: PROD_MAP.get(int(x), 'OTRO') if pd.notna(x) else 'OTRO'
    )
    # Clasificación DOM/INT: misma lógica que Ecuador (pos9 del flex field Oracle)
    df_usa['_di'] = df_usa['DomInt_Flex'].map(
        lambda x: 'INT' if x in ('0002', '0112') else ('DOM' if x == '0001' else 'N/A')
    )

    is_ic = (df_usa[col_pl2] == 'Intercompany Transactions') if col_pl2 \
            else pd.Series(False, index=df_usa.index)
    df_ni = df_usa[~is_ic]   # filas no-intercompany

    # ── Llaves de órdenes para entidad 105 ──────────────────────────────────
    def _llave(key, row, col):
        try:
            return float(llaves_ordenes[key].loc[row, col]) / 100
        except Exception:
            return 0.0

    llave_HI      = _llave('K11', 105, 'Int')       # % HT INT / (HT DOM+INT)
    llave_OI      = _llave('K12', 105, 'Int')       # % ONA INT / (ONA DOM+INT)
    llave_INT     = _llave('K9',  105, 'Int')       # % Corporate INT
    llave_HT_pct  = _llave('K10', 105, 'Hoteles')  # % HT sobre (HT+ONA)
    llave_ONA_pct = _llave('K10', 105, 'ONA')      # % ONA sobre (HT+ONA)

    # ── Cargar diccionario PL ────────────────────────────────────────────────
    df_pl_local = cargar_pl(RUTA_RUBROS)

    if df_pl_local is not None:
        df_usa_pl = df_usa.copy()
        df_usa_pl['_is_ic'] = is_ic.values
        df_usa_pl = df_usa_pl.merge(df_pl_local, left_on='Account_Num', right_on='Cuenta', how='left')

        df_net_ni = df_usa_pl[
            (df_usa_pl['PL_Totalizador'] == 'Net Revenues') &
            (~df_usa_pl['_is_ic'])
        ]
    else:
        df_usa_pl = None
        df_net_ni = df_ni[df_ni['Account_Num'].between(40000, 49999)]

    # ── INGRESOS ─────────────────────────────────────────────────────────────

    # ── Fraudes por producto y DOM/INT (53300/53305) ─────────────────────────
    # Los fraudes se asignan al producto Y al DOM/INT que les corresponde en Oracle.
    #
    # Para HT: los fraudes NO son todos DOM — tienen split DOM/INT real en Oracle:
    #   - fraudes_HT_DOM = balance HT DOM de 53300/53305 = -7,426.03 (crédito/ingreso)
    #   - fraudes_HT_INT = balance HT INT de 53300/53305 = +4,054.11 (débito/costo)
    #   - neto = -3,371.92 ✅ (= fraudes_HT total)
    # Para ONA: todos INT (confirmado, H31 = match exacto)
    # Para VUELOS/VP: no hay split relevante
    #
    # Fórmula Excel validada:
    #   G31 = -Ingresos!D18 - E9  donde D18 = D5 + fraudes_HT_INT
    #   E31 = -Ingresos!C18        donde C18 = C5 + fraudes_HT_DOM
    #   D5 = GETPIVOTDATA(HT, INT) = 12,150.05  | C5 = GETPIVOTDATA(HT, DOM) = 1,845.79
    # ✅ Validado vs Segmentación Nov-2025: G31=1,058,491.84 | E31=5,580.24
    _fraudes       = df_usa[df_usa['Account_Num'].isin([53300, 53305])]
    fraudes_VUE    = -(_fraudes[_fraudes['_prod'] == 'VUELOS'         ][col_bal].sum())  # +11,793.48
    fraudes_HT_INT =   _fraudes[(_fraudes['_prod'] == 'HT') & (_fraudes['_di'] == 'INT')][col_bal].sum()  # +4,054.11 (débito)
    fraudes_HT_DOM =   _fraudes[(_fraudes['_prod'] == 'HT') & (_fraudes['_di'] == 'DOM')][col_bal].sum()  # -7,426.03 (crédito)
    fraudes_ONA    = -(_fraudes[_fraudes['_prod'] == 'ONA'            ][col_bal].sum())  # +2,094.40
    fraudes_VP     = -(_fraudes[_fraudes['_prod'] == 'VUELOS PAQUETES'][col_bal].sum())  #     +0.09

    # G31: HT INT income (non-IC 40000-49999) + 49102 HT - fraudes HT INT (costo)
    # fraudes_HT_INT = +4,054.11 (débito = costo que reduce G31)
    # ✅ G31 = -(12,150.05 + 4,054.11) + 1,074,696 = 1,058,491.84
    G31 = -(df_net_ni[(df_net_ni['_prod'] == 'HT') & (df_net_ni['_di'] == 'INT')][col_bal].sum()) \
          -(df_usa[(df_usa['Account_Num'] == 49102) & (df_usa['_prod'] == 'HT') ][col_bal].sum()) \
          - fraudes_HT_INT

    # H31: ONA INT income + 49102 ONA + fraudes ONA (100% INT, confirmado ✅)
    # ✅ 37,149.56 + 2,094.40 = 39,243.96 match exacto vs Excel H31.
    H31 = -(df_net_ni[(df_net_ni['_prod'] == 'ONA') & (df_net_ni['_di'] == 'INT')][col_bal].sum()) \
          -(df_usa[(df_usa['Account_Num'] == 49102) & (df_usa['_prod'] == 'ONA')][col_bal].sum()) \
          + fraudes_ONA

    # D31: Vuelos income (todos DOM/INT/N/A) + fraudes VUELOS
    # ✅ D31 = 775,120.34 match exacto vs Excel D31.
    D31 = -(df_net_ni[df_net_ni['_prod'] == 'VUELOS'][col_bal].sum()) \
          + fraudes_VUE

    # E31: HT income DOM ONLY (NO N/A — el pivot Excel $B$35 usa DOM/INT="DOM" estrictamente)
    # + fraudes HT DOM: fraudes_HT_DOM = -7,426.03 (crédito) → -fraudes_HT_DOM = +7,426.03 (ingreso)
    # ✅ E31 = -(1,845.79) - (-7,426.03) = -1,845.79 + 7,426.03 = 5,580.24
    E31 = -(df_net_ni[(df_net_ni['_prod'] == 'HT') & (df_net_ni['_di'] == 'DOM')][col_bal].sum()) \
          - fraudes_HT_DOM

    # F31: ONA income DOM + N/A + VP (todo DOM por regla ingresos) + fraudes VP
    # fraudes ONA van 100% a H31 (INT). Solo fraudes VP quedan aquí.
    F31 = -(df_net_ni[(df_net_ni['_prod'] == 'ONA') & (df_net_ni['_di'].isin(['DOM', 'N/A']))][col_bal].sum()) \
          -(df_net_ni[df_net_ni['_prod'] == 'VUELOS PAQUETES'][col_bal].sum()) \
          + fraudes_VP

    # ── HOSTING BASE (se calcula ANTES de RFC porque M5_rfc lo necesita) ────────
    # Costos_Hosting = TODAS las cuentas con Mapeo PL Nivel 1 == 'IT - Personnel/Expenses'
    # (todos los productos, non-IC + IC) MENOS 70105 IC (Licencias IT) y 70114 IC (Mantenimiento IT)
    # Se usa la columna nativa del balance en lugar de PL_Totalizador (que no cubre cuentas USA)
    # Fórmula equivalente al Pivot 105 del Excel de Hosting USA.
    if col_pl1 is not None:
        # Suma total IT-Personnel/Expenses (incluye todos los productos: CORPORATE, VUELOS, HT, ONA, VP)
        _total_it = df_usa[df_usa[col_pl1] == HOSTING_PL1_USA][col_bal].sum()
        # Restar Gastos Interco 70105 (Licencias IT) y 70114 (Mantenimiento IT) IC
        _ic_restar = df_usa[
            (df_usa[col_pl1] == HOSTING_PL1_USA) &
            (df_usa[col_pl2] == 'Intercompany Transactions') &
            (df_usa['Account_Num'].isin([70105, 70114]))
        ][col_bal].sum()
        Costos_Hosting = _total_it - _ic_restar
    else:
        # Fallback: suma total CORPORATE non-IC (menos preciso; solo si falta columna PL Nivel 1)
        Costos_Hosting = df_ni[df_ni['_prod'] == 'CORPORATE'][col_bal].sum()
        print("⚠️  USA Hosting: columna 'Mapeo PL Nivel 1' no encontrada — usando fallback CORPORATE total")

    # ── COSTOS RFC ───────────────────────────────────────────────────────────
    # E18/F18: costos directos HT/ONA non-IC, excluye:
    #   - EXCL_COST_ACCOUNTS (64201, 64202, 68210, 41311, 41312, 53300, 53305)
    #   - EXCL_PL1_USA = incentivos de ingresos (Back End/Other/Up Front Incentives)
    #                  + categorías no operativas (Interest Income, FX/Other)
    # Se usa Mapeo PL Nivel 1 nativo (no PL_Totalizador, cobertura incompleta en USA)
    if col_pl1 is not None:
        mask_ht = (
            (df_ni['_prod'] == 'HT') &
            (~df_ni['Account_Num'].isin(EXCL_COST_ACCOUNTS)) &
            (~df_ni[col_pl1].isin(EXCL_PL1_USA))
        )
        mask_ona = (
            (df_ni['_prod'] == 'ONA') &
            (~df_ni['Account_Num'].isin(EXCL_COST_ACCOUNTS)) &
            (~df_ni[col_pl1].isin(EXCL_PL1_USA))
        )
        E18_total = df_ni[mask_ht ][col_bal].sum()
        F18_total = df_ni[mask_ona][col_bal].sum()
    elif df_usa_pl is not None:
        # Fallback: usa PL_Totalizador si no hay Mapeo PL Nivel 1
        df_costs_base = df_usa_pl[
            (df_usa_pl['PL_Totalizador'].isin(RFC_COST_RUBROS)) &
            (~df_usa_pl['_is_ic']) &
            (~df_usa_pl['Account_Num'].isin(EXCL_COST_ACCOUNTS))
        ]
        E18_total = df_costs_base[df_costs_base['_prod'] == 'HT' ][col_bal].sum()
        F18_total = df_costs_base[df_costs_base['_prod'] == 'ONA'][col_bal].sum()
        print("⚠️  USA E18/F18: usando PL_Totalizador como fallback (Mapeo PL Nivel 1 no disponible)")
    else:
        E18_total = 0.0
        F18_total = 0.0
        print("⚠️  USA E18/F18: sin columna PL Nivel 1 ni diccionario PL — valores en 0")

    # M5 total (CORPORATE non-IC, solo categorías CORP_RFC_PL1_USA)
    # Fórmula validada vs Estimado!M6 del Excel (= 2,483,539.18):
    #   Base = CORP_RFC_PL1_USA excl. CORP_M5_EXCL_ACCOUNTS
    #          + CORP_M5_EXTRA_ACCOUNTS (incluir aunque su PL1 no esté en CORP_RFC_PL1_USA)
    # ✅ Validado Nov-25: 2,483,539.18 (ajuste de cuentas 80101/80102/64341)
    if col_pl1 is not None:
        mask_m5_base  = (
            (df_ni['_prod'] == 'CORPORATE') &
            (df_ni[col_pl1].isin(CORP_RFC_PL1_USA)) &
            (~df_ni['Account_Num'].isin(CORP_M5_EXCL_ACCOUNTS))
        )
        mask_m5_extra = (
            (df_ni['_prod'] == 'CORPORATE') &
            (df_ni['Account_Num'].isin(CORP_M5_EXTRA_ACCOUNTS))
        )
        M5 = df_ni[mask_m5_base | mask_m5_extra][col_bal].sum()
    elif df_usa_pl is not None:
        # Fallback: usa PL_Totalizador si no hay Mapeo PL Nivel 1
        df_corp_m5 = df_usa_pl[
            (df_usa_pl['_prod'] == 'CORPORATE') &
            (df_usa_pl['PL_Totalizador'].isin(RFC_COST_RUBROS)) &
            (~df_usa_pl['_is_ic'])
        ]
        M5 = df_corp_m5[col_bal].sum()
        print("⚠️  USA M5: usando PL_Totalizador como fallback (Mapeo PL Nivel 1 no disponible)")
    else:
        M5 = df_ni[df_ni['_prod'] == 'CORPORATE'][col_bal].sum()
        print("⚠️  USA M5: usando total CORPORATE como fallback — resultado puede ser incorrecto")

    # M5_rfc: se EXCLUYE Hosting del corporate para RFC (los Hosting costs ya se
    # facturan por separado → no deben entrar dos veces en la base de costo del RFC)
    # Ref: Ref Cte sheet columna "Costo Servicios (AR, USA)" = Hosting × llave_HT × llave_INT
    M5_rfc = M5 - Costos_Hosting

    E22 = M5_rfc * llave_HT_pct   # Corporate RFC asignado a HT
    F22 = M5_rfc * llave_ONA_pct  # Corporate RFC asignado a ONA

    # G33/H33: costos INT para RFC (directos × llave + corporate_net × llave)
    G33 = E18_total * llave_HI + E22 * llave_INT
    H33 = F18_total * llave_OI + F22 * llave_INT

    # ── RFC (Referencia de Clientes) ──────────────────────────────────────────
    # Fórmula USA (IDÉNTICA a Ecuador — confirmado contra Segmentación Nov-2025):
    #   RFC = (G31 − G33 × MARKUP) + (H31 − H33 × MARKUP)
    #
    # G31 ya incluye 49102 (RFC ya facturado), por lo que la diferencia G31-G33×MARKUP
    # captura tanto el devengamiento del período como el ajuste del ya contabilizado.
    #
    # ✅ Validado Nov-2025 (Segmentación Excel Hoja Llave B4):
    #   G31=1,058,491.84 | G33=1,097,306.44  → HT: 1,058,492 − 1,097,306×1.095 = −143,059
    #   H31=39,243.96    | H33=40,853.39      → ONA: 39,244 − 40,853×1.095 = −5,490
    #   RFC = −143,059 + (−5,490) = −148,549.21 ✅
    #   Nuestro resultado (Oracle 07-nov): −148,924.07 (diff = −374.86, 0.25% — timing)
    RFC = (G31 - G33 * MARKUP_RFC_USA) + (H31 - H33 * MARKUP_RFC_USA)

    # ── LICENCIAS — catch-up YTD + accrual mensual ───────────────────────────
    # 70104 (Marca) y 70105 (IT) son POSITIVOS en Oracle para USA (igual que Ecuador).
    # ✅ Validado Nov-2025: 70104 VUELOS = +21,365 | 70105 VUELOS = +42,731
    # NO aplicar negación — los valores ya representan el monto ya contabilizado.
    D65 = df_usa[(df_usa['Account_Num'] == 70104) & (df_usa['_prod'] == 'VUELOS')][col_bal].sum()
    E65 = df_usa[(df_usa['Account_Num'] == 70104) & (df_usa['_prod'] == 'HT')   ][col_bal].sum()
    F65 = df_usa[(df_usa['Account_Num'] == 70104) & (df_usa['_prod'] == 'ONA')  ][col_bal].sum()

    D70 = df_usa[(df_usa['Account_Num'] == 70105) & (df_usa['_prod'] == 'VUELOS')][col_bal].sum()
    E70 = df_usa[(df_usa['Account_Num'] == 70105) & (df_usa['_prod'] == 'HT')   ][col_bal].sum()
    F70 = df_usa[(df_usa['Account_Num'] == 70105) & (df_usa['_prod'] == 'ONA')  ][col_bal].sum()

    D61 = D31 * TASA_MARCA;   E61 = E31 * TASA_MARCA;   F61 = F31 * TASA_MARCA
    D62 = D31 * TASA_IT;      E62 = E31 * TASA_IT;      F62 = F31 * TASA_IT

    D67 = (D61 - D65) + D61 / mes_base
    E67 = (E61 - E65) + E61 / mes_base
    F67 = (F61 - F65) + F61 / mes_base

    D72 = (D62 - D70) + D62 / mes_base
    E72 = (E62 - E70) + E62 / mes_base
    F72 = (F62 - F70) + F62 / mes_base

    Lic_Marca = D67 + E67 + F67
    Lic_IT    = D72 + E72 + F72

    # ── SERVICIOS DE HOSTING (USA → AR) ───────────────────────────────────────
    # Costos_Hosting ya fue calculado al inicio (CORPORATE IT, 'IT - Personnel/Expenses', no IC).
    # Ya facturado: 49120 en Oracle es ingreso (crédito = negativo) → negamos para valor positivo
    # Fórmula (lineamientos, "mes corriente"):
    #   Hosting = (Costos × MARKUP − Ya_49120) + Costos × MARKUP / mes_base
    # Validación Nov-2025:
    #   Costos=1,394,173.78 | Ya=1,507,665.55 | mes=10
    #   → (1,477,824.21 − 1,507,665.55) + 147,782.42 = $117,941.08 ✅
    Ya_Hosting    = -(df_usa[df_usa['Account_Num'] == 49120][col_bal].sum())
    Gastos_Hosting = Costos_Hosting * MARKUP_HOSTING
    Hosting        = (Gastos_Hosting - Ya_Hosting) + Gastos_Hosting / mes_base

    return {
        'Licencia_Marca': Lic_Marca,
        'Licencia_IT':    Lic_IT,
        'RFC':            RFC,
        'Hosting':        Hosting,
        # Intermedios de diagnóstico
        '_G31': G31, '_H31': H31, '_D31': D31, '_E31': E31, '_F31': F31,
        '_G33': G33, '_H33': H33,
        '_E18': E18_total, '_F18': F18_total,
        '_E22': E22, '_F22': F22,
        '_M5': M5, '_M5_rfc': M5_rfc,
        '_llave_HI': llave_HI, '_llave_OI': llave_OI,
        '_llave_INT': llave_INT, '_llave_HT': llave_HT_pct, '_llave_ONA': llave_ONA_pct,
        '_mes_base': mes_base,
        '_Costos_Hosting': Costos_Hosting, '_Ya_Hosting': Ya_Hosting,
    }


# ============================================================================
# CÁLCULO TR (TRAVEL RESERVATION) - HOJA LLAVE
# ============================================================================

def calcular_tr_hoja_llave(df_balance_raw, mes_base=2):
    """
    Calcula Licencia IT para Travel Reservation (entidad 601) en USD.

    TR solo factura Licencia IT (5%) — no tiene Marca, RFC ni Hosting.
    Prestador: Despegar.com.ar (Argentina)
    Prestatario: Travel Reservations S.R.L (Uruguay)

    Lógica:
      1. Net Revenue por producto (VUELOS, HT, ONA) desde GL non-IC
         - PUBLICIDAD se EXCLUYE del cálculo de billing
         - VUELOS PAQUETES se suma a ONA
      2. Agrega fraudes (53300/53305) por producto
      3. Regalia Real = -(Net Revenue incl fraudes) × 5%
      4. Facturación = (Real - Booked 70105) + Real / mes_base

    Moneda: USD (nativo Oracle, sin conversión).

    Targets validados Feb-2026 (TR - Licencia IT Mar-2026.xlsx):
      Licencia IT = $636,242.90 USD
    """

    col_bal = next((c for c in df_balance_raw.columns if 'Saldo Final' in c or 'Balance' in c),
                   df_balance_raw.columns[-1])
    col_pl2 = next((c for c in df_balance_raw.columns if 'Nivel 2' in c), None)

    df_tr = df_balance_raw[df_balance_raw['Entity_Flex'] == '601'].copy()
    if df_tr.empty:
        print("⚠️  TR: sin registros para entidad 601 en el balance")
        return None

    df_tr['_prod_code'] = pd.to_numeric(df_tr['Product_Flex'], errors='coerce')
    df_tr['_prod'] = df_tr['_prod_code'].map(
        lambda x: PROD_MAP.get(int(x), 'OTRO') if pd.notna(x) else 'OTRO'
    )

    is_ic = (df_tr[col_pl2] == 'Intercompany Transactions') if col_pl2 \
            else pd.Series(False, index=df_tr.index)

    # ── Cargar PL dict ────────────────────────────────────────────────────────
    df_pl_local = cargar_pl(RUTA_RUBROS)

    if df_pl_local is not None:
        df_tr_pl = df_tr.copy()
        df_tr_pl['_is_ic'] = is_ic.values
        df_tr_pl = df_tr_pl.merge(df_pl_local, left_on='Account_Num', right_on='Cuenta', how='left')

        df_net_ni = df_tr_pl[
            (df_tr_pl['PL_Totalizador'] == 'Net Revenues') &
            (~df_tr_pl['_is_ic'])
        ]
    else:
        df_tr_pl = None
        df_net_ni = df_tr[~is_ic & df_tr['Account_Num'].between(40000, 49999)]

    # ── Net Revenue por producto (VUELOS, HT, ONA) ──────────────────────────
    # PUBLICIDAD se excluye del cálculo de billing (confirmado vs Excel)
    # VUELOS PAQUETES se suma a ONA (confirmado: ONA Excel = ONA pivot + VP pivot)
    nr_vuelos = df_net_ni[df_net_ni['_prod'] == 'VUELOS'][col_bal].sum()
    nr_ht     = df_net_ni[df_net_ni['_prod'] == 'HT'][col_bal].sum()
    nr_ona    = df_net_ni[df_net_ni['_prod'].isin(['ONA', 'VUELOS PAQUETES'])][col_bal].sum()

    # ── Fraudes (53300/53305) por producto ───────────────────────────────────
    # Se suman al Net Revenue antes de calcular la regalia (confirmado vs Excel)
    _fraudes = df_tr[df_tr['Account_Num'].isin([53300, 53305])]
    fr_vuelos = _fraudes[_fraudes['_prod'] == 'VUELOS'][col_bal].sum()
    fr_ht     = _fraudes[_fraudes['_prod'] == 'HT'][col_bal].sum()
    fr_ona    = _fraudes[_fraudes['_prod'].isin(['ONA', 'VUELOS PAQUETES'])][col_bal].sum()

    # Net Revenue incl fraudes (en Oracle: revenue es negativo)
    total_vuelos = nr_vuelos + fr_vuelos
    total_ht     = nr_ht + fr_ht
    total_ona    = nr_ona + fr_ona

    # ── Regalia Real = -(Net Revenue incl fraudes) × tasa ────────────────────
    real_vuelos = -(total_vuelos) * TASA_IT
    real_ht     = -(total_ht) * TASA_IT
    real_ona    = -(total_ona) * TASA_IT

    # ── Booked 70105 por producto (ONA incluye VP) ──────────────────────────
    booked_vuelos = df_tr[(df_tr['Account_Num'] == 70105) & (df_tr['_prod'] == 'VUELOS')][col_bal].sum()
    booked_ht     = df_tr[(df_tr['Account_Num'] == 70105) & (df_tr['_prod'] == 'HT')][col_bal].sum()
    booked_ona    = df_tr[(df_tr['Account_Num'] == 70105) &
                          df_tr['_prod'].isin(['ONA', 'VUELOS PAQUETES'])][col_bal].sum()

    # ── Facturación = (Real - Booked) + Real / mes_base ─────────────────────
    fact_vuelos = (real_vuelos - booked_vuelos) + real_vuelos / mes_base
    fact_ht     = (real_ht - booked_ht) + real_ht / mes_base
    fact_ona    = (real_ona - booked_ona) + real_ona / mes_base

    Lic_IT = fact_vuelos + fact_ht + fact_ona

    return {
        'Licencia_IT': Lic_IT,
        # Intermedios de diagnóstico
        '_nr_vuelos': nr_vuelos, '_nr_ht': nr_ht, '_nr_ona': nr_ona,
        '_fr_vuelos': fr_vuelos, '_fr_ht': fr_ht, '_fr_ona': fr_ona,
        '_total_vuelos': total_vuelos, '_total_ht': total_ht, '_total_ona': total_ona,
        '_real_vuelos': real_vuelos, '_real_ht': real_ht, '_real_ona': real_ona,
        '_booked_vuelos': booked_vuelos, '_booked_ht': booked_ht, '_booked_ona': booked_ona,
        '_fact_vuelos': fact_vuelos, '_fact_ht': fact_ht, '_fact_ona': fact_ona,
        '_mes_base': mes_base,
    }


# ============================================================================
# CÁLCULO ESPAÑA - HOJA LLAVE
# ============================================================================

def calcular_espana_hoja_llave(df_balance_raw, tc_eur, mes_base=2):
    """
    Calcula Licencia IT y Licencia Marca para España (entidad 611) en USD.

    España factura IT (5%) y Marca (2.5%) — no tiene RFC ni Hosting.
    Prestador IT: Despegar.com.ar (Argentina)
    Prestador Marca: Travel Reservations (601)
    Prestatario: Despegar España

    Moneda funcional: EUR → divide por tc_eur (EUR/USD) para resultado en USD.
    Solo tiene productos HT y ONA (sin Vuelos).

    Targets validados Feb-2026 (ESP - Licencia IT y Marca Mar-2026.xlsx):
      Licencia IT    = 123,865.31 EUR → $143,820.01 USD
      Licencia Marca =  59,843.53 EUR → $ 69,484.32 USD
    """

    col_bal = next((c for c in df_balance_raw.columns if 'Saldo Final' in c or 'Balance' in c),
                   df_balance_raw.columns[-1])
    col_pl2 = next((c for c in df_balance_raw.columns if 'Nivel 2' in c), None)

    df_esp = df_balance_raw[df_balance_raw['Entity_Flex'] == '611'].copy()
    if df_esp.empty:
        print("⚠️  España: sin registros para entidad 611 en el balance")
        return None

    df_esp['_prod_code'] = pd.to_numeric(df_esp['Product_Flex'], errors='coerce')
    df_esp['_prod'] = df_esp['_prod_code'].map(
        lambda x: PROD_MAP.get(int(x), 'OTRO') if pd.notna(x) else 'OTRO'
    )

    is_ic = (df_esp[col_pl2] == 'Intercompany Transactions') if col_pl2 \
            else pd.Series(False, index=df_esp.index)

    # ── Cargar PL dict ────────────────────────────────────────────────────────
    df_pl_local = cargar_pl(RUTA_RUBROS)

    if df_pl_local is not None:
        df_esp_pl = df_esp.copy()
        df_esp_pl['_is_ic'] = is_ic.values
        df_esp_pl = df_esp_pl.merge(df_pl_local, left_on='Account_Num', right_on='Cuenta', how='left')

        df_net_ni = df_esp_pl[
            (df_esp_pl['PL_Totalizador'] == 'Net Revenues') &
            (~df_esp_pl['_is_ic'])
        ]
    else:
        df_esp_pl = None
        df_net_ni = df_esp[~is_ic & df_esp['Account_Num'].between(40000, 49999)]

    # ── Net Revenue por producto (HT, ONA) — España no tiene Vuelos ─────────
    nr_ht  = df_net_ni[df_net_ni['_prod'] == 'HT'][col_bal].sum()
    nr_ona = df_net_ni[df_net_ni['_prod'] == 'ONA'][col_bal].sum()

    # ── Fraudes (53300/53305) por producto ───────────────────────────────────
    # España típicamente no tiene 53300/53305 — sus ajustes menores (Ajustes de
    # Tolerancia) vienen de cuentas que ya están en PL='Net Revenues'.
    # Se buscan igualmente por robustez.
    _fraudes = df_esp[df_esp['Account_Num'].isin([53300, 53305])]
    fr_ht  = _fraudes[_fraudes['_prod'] == 'HT'][col_bal].sum()
    fr_ona = _fraudes[_fraudes['_prod'] == 'ONA'][col_bal].sum()

    total_ht  = nr_ht + fr_ht
    total_ona = nr_ona + fr_ona

    # ── Regalia Real IT (5%) y Marca (2.5%) ─────────────────────────────────
    real_it_ht  = -(total_ht) * TASA_IT
    real_it_ona = -(total_ona) * TASA_IT
    real_mk_ht  = -(total_ht) * TASA_MARCA
    real_mk_ona = -(total_ona) * TASA_MARCA

    # ── Booked IT (70105) y Marca (70104) ───────────────────────────────────
    booked_it_ht  = df_esp[(df_esp['Account_Num'] == 70105) & (df_esp['_prod'] == 'HT')][col_bal].sum()
    booked_it_ona = df_esp[(df_esp['Account_Num'] == 70105) & (df_esp['_prod'] == 'ONA')][col_bal].sum()
    booked_mk_ht  = df_esp[(df_esp['Account_Num'] == 70104) & (df_esp['_prod'] == 'HT')][col_bal].sum()
    booked_mk_ona = df_esp[(df_esp['Account_Num'] == 70104) & (df_esp['_prod'] == 'ONA')][col_bal].sum()

    # ── Facturación IT = (Real - Booked) + Real / mes_base ──────────────────
    fact_it_ht  = (real_it_ht - booked_it_ht) + real_it_ht / mes_base
    fact_it_ona = (real_it_ona - booked_it_ona) + real_it_ona / mes_base
    Lic_IT_EUR  = fact_it_ht + fact_it_ona

    # ── Facturación Marca = (Real - Booked) + Real / mes_base ───────────────
    fact_mk_ht  = (real_mk_ht - booked_mk_ht) + real_mk_ht / mes_base
    fact_mk_ona = (real_mk_ona - booked_mk_ona) + real_mk_ona / mes_base
    Lic_Marca_EUR = fact_mk_ht + fact_mk_ona

    # ── Conversión EUR → USD ────────────────────────────────────────────────
    Lic_IT    = Lic_IT_EUR / tc_eur
    Lic_Marca = Lic_Marca_EUR / tc_eur

    return {
        'Licencia_IT':        Lic_IT,
        'Licencia_Marca':     Lic_Marca,
        'Licencia_IT_EUR':    Lic_IT_EUR,
        'Licencia_Marca_EUR': Lic_Marca_EUR,
        # Intermedios de diagnóstico
        '_nr_ht': nr_ht, '_nr_ona': nr_ona,
        '_fr_ht': fr_ht, '_fr_ona': fr_ona,
        '_total_ht': total_ht, '_total_ona': total_ona,
        '_real_it_ht': real_it_ht, '_real_it_ona': real_it_ona,
        '_real_mk_ht': real_mk_ht, '_real_mk_ona': real_mk_ona,
        '_booked_it_ht': booked_it_ht, '_booked_it_ona': booked_it_ona,
        '_booked_mk_ht': booked_mk_ht, '_booked_mk_ona': booked_mk_ona,
        '_fact_it_ht': fact_it_ht, '_fact_it_ona': fact_it_ona,
        '_fact_mk_ht': fact_mk_ht, '_fact_mk_ona': fact_mk_ona,
        '_tc_eur': tc_eur, '_mes_base': mes_base,
    }


# ============================================================================
# BRASIL (104) — Licencia Marca + Dominio + IT + RFC TR + RFC ES
# ============================================================================

def calcular_brasil_hoja_llave(df_balance_raw, llaves_ordenes, tc_brl, mes_base=10):
    """
    Calcula los 5 valores de Hoja Llave Brasil (entidad 104) en USD:
      - Licencia de Marca   (1.0%, Prestador: Travel Reservations)
      - Licencia de Dominio (1.5%, Prestador: Travel Reservations)
      - Licencia de IT      (5.0%, Prestador: Despegar.com.ar)
      - Referencia de Clientes TR  (10.5% markup, vs Travel Reservations)
      - Referencia de Clientes ES  (10.5% markup, vs España — desde Oct 2025)

    Particularidades:
      - Moneda funcional: BRL → divide por tc_brl para USD
      - Mayor = 'DESPEGAR BRASIL CORP' (solo BRL entries)
      - Ingresos Locales = NOT INT (DOM + N/A) + distribución CORPORATE × llave producto
      - VUELOS y VP → todo LOCAL independientemente del flex
      - ONA Local incluye VP (Vuelos Paquetes se suma a ONA)
      - Fraudes (53300/53305) separados del NR (Cost of Revenue), sumados por separado
      - Marca/Dominio booked: 70104 con subcuenta (1/401=Dominio, 2/402/403/404=Marca)
      - IT booked: 70105 por producto
      - 49108: PL2=Intercompany pero Excel lo incluye en NR CORPORATE
      - RFC usa split Ene-Sept / Oct-XX para separar TR vs España
      - Markup RFC = 1.105 (10.5%)
    """

    TASA_MARCA   = 0.01
    TASA_DOMINIO = 0.015
    TASA_IT      = 0.05
    MARKUP_RFC   = 1.105

    # PROD_MAP y DI_INT se usan desde las variables globales (cargadas del template)
    FRAUDE_CTAS = {'53300', '53305'}

    col_bal = next((c for c in df_balance_raw.columns if 'Saldo Final' in c or 'Balance' in c),
                   df_balance_raw.columns[-1])
    col_pl2 = next((c for c in df_balance_raw.columns if 'Nivel 2' in c), None)
    col_mayor = next((c for c in df_balance_raw.columns if 'Mayor' in c), None)

    # ── Filtrar entidad 104, mayor BRL ──
    df_br = df_balance_raw[df_balance_raw['Entity_Flex'] == '104'].copy()
    if col_mayor:
        _mayors_br = {'DESPEGAR BRASIL CORP', 'DECOLAR OPER CORP'}
        df_br = df_br[df_br[col_mayor].isin(_mayors_br)]
    if df_br.empty:
        print("⚠️  Brasil: sin registros para entidad 104 en el balance")
        return None

    df_br['_prod_code'] = pd.to_numeric(df_br['Product_Flex'], errors='coerce')
    df_br['_prod'] = df_br['_prod_code'].map(
        lambda x: PROD_MAP.get(int(x), 'OTRO') if pd.notna(x) else 'OTRO'
    )
    df_br['_di'] = df_br['DomInt_Flex'].map(
        lambda x: 'INT' if x in DI_INT else 'LOCAL'
    )
    if 'Account_Num' in df_br.columns:
        df_br['_acct'] = df_br['Account_Num'].apply(lambda x: str(int(x)) if pd.notna(x) else '')
    else:
        df_br['_acct'] = df_br.iloc[:, 2].astype(str).str.strip()

    is_ic = (df_br[col_pl2] == 'Intercompany Transactions') if col_pl2 \
        else pd.Series(False, index=df_br.index)
    df_ni = df_br[~is_ic]

    # Merge PL para obtener PL_Totalizador
    df_pl_local = cargar_pl(RUTA_RUBROS)
    df_pl_local['Cuenta'] = df_pl_local['Cuenta'].astype(str).str.strip()
    df_ni = df_ni.copy()
    df_ni['_acct'] = df_ni['_acct'].astype(str).str.strip()
    df_ni = df_ni.merge(df_pl_local[['Cuenta', 'PL_Totalizador']],
                        left_on='_acct', right_on='Cuenta', how='left')

    # ── NET REVENUE non-IC (sin fraudes) ──
    nr_mask = (df_ni['PL_Totalizador'] == 'Net Revenues') & (~df_ni['_acct'].isin(FRAUDE_CTAS))
    df_nr_nf = df_ni[nr_mask].copy()

    # 49108: PL2=Intercompany pero Excel lo incluye en Segmentación row 7 como CORPORATE NR
    # Sin embargo, NO entra en la distribución por llave para Ingresos Locales
    # (el ajuste Corporate rows 48-50 solo toma el NR CORPORATE del pivot sin 49108)
    # Se excluye del df_nr_nf para no distorsionar corp_nr

    # ── Fraudes (Cost of Revenue, fuera del NR) ──
    df_fraudes = df_ni[df_ni['_acct'].isin(FRAUDE_CTAS)].copy()

    # ── CORPORATE NR → distribuir por llave producto ──
    corp_nr = df_nr_nf[df_nr_nf['_prod'] == 'CORPORATE'][col_bal].sum()

    # Llaves producto K10 para entity 104
    k10 = llaves_ordenes.get('K10', {}).get('104', {})
    k_vuelos = k10.get('Vuelos', 0) + k10.get('Vuelos Paquetes', 0)
    # Alternativa: usar porcentajes de órdenes directamente
    # Si K10 no tiene 104, usar los valores conocidos
    if k_vuelos == 0:
        # Fallback: valores de Nov-2025 Excel
        k_vuelos = 0.4734735308542847
        k_ht = 0.22405597415685927
        k_ona = 0.20377051589017905
        k_vp = 0.098699979098677
    else:
        k_ht = k10.get('Hoteles', 0)
        k_ona = k10.get('ONA', 0)
        k_vp = k10.get('Vuelos Paquetes', 0)
        k_vuelos = k10.get('Vuelos', 0)

    corp_adj = {
        'VUELOS': corp_nr * k_vuelos,
        'HT': corp_nr * k_ht,
        'ONA': corp_nr * k_ona,
        'VUELOS PAQUETES': corp_nr * k_vp,
    }

    # ── Helper: saldo LOCAL (= NOT INT) para un producto ──
    def local_sum(df, prod):
        """Para VUELOS/VP retorna TODO (es local). Para HT/ONA retorna not-INT."""
        if prod in ('VUELOS', 'VUELOS PAQUETES'):
            return df[df['_prod'] == prod][col_bal].sum()
        return df[(df['_prod'] == prod) & (df['_di'] != 'INT')][col_bal].sum()

    # ── Ingresos Locales (base licencias) ──
    vuelos_local = -(local_sum(df_nr_nf, 'VUELOS') + local_sum(df_fraudes, 'VUELOS') + corp_adj['VUELOS'])
    ht_local = -(local_sum(df_nr_nf, 'HT') + local_sum(df_fraudes, 'HT') + corp_adj['HT'])
    ona_raw = -(local_sum(df_nr_nf, 'ONA') + local_sum(df_fraudes, 'ONA') + corp_adj['ONA'])
    vp_local = -(local_sum(df_nr_nf, 'VUELOS PAQUETES') + local_sum(df_fraudes, 'VUELOS PAQUETES') + corp_adj['VUELOS PAQUETES'])
    ona_local = ona_raw + vp_local  # ONA Local incluye VP

    # ── Booked licencias ──
    # 70104: Marca (subcuenta 2/402/403/404) y Dominio (subcuenta 1/401)
    booked_70104 = df_br[(df_br['_acct'] == '70104') & is_ic]
    marca_booked = {'VUELOS': 0, 'HT': 0, 'ONA': 0}
    dom_booked = {'VUELOS': 0, 'HT': 0, 'ONA': 0}
    col_sub = next((c for c in df_br.columns if 'Subcuenta' in c and 'Nombre' not in c), None)
    if col_sub and not booked_70104.empty:
        for _, row in booked_70104.iterrows():
            prod = PROD_MAP.get(int(row['_prod_code']) if pd.notna(row['_prod_code']) else -1, 'OTRO')
            sc = str(row[col_sub]).strip()
            if sc in ('1', '401'):
                dom_booked[prod] = dom_booked.get(prod, 0) + row[col_bal]
            elif sc in ('2', '402', '403', '404'):
                marca_booked[prod] = marca_booked.get(prod, 0) + row[col_bal]

    # 70105: IT
    booked_70105 = df_br[df_br['_acct'] == '70105']
    it_booked = {'VUELOS': 0, 'HT': 0, 'ONA': 0}
    if not booked_70105.empty:
        for _, row in booked_70105.iterrows():
            prod = PROD_MAP.get(int(row['_prod_code']) if pd.notna(row['_prod_code']) else -1, 'OTRO')
            it_booked[prod] = it_booked.get(prod, 0) + row[col_bal]

    # ── Fórmula licencias: (Real × tasa - Booked) + Real × tasa / mes_base ──
    def lic(real_val, tasa, booked_val):
        r = real_val * tasa
        return (r - booked_val) + r / mes_base

    reales = {'VUELOS': vuelos_local, 'HT': ht_local, 'ONA': ona_local}

    marca_brl = sum(lic(reales[p], TASA_MARCA, marca_booked.get(p, 0)) for p in ['VUELOS', 'HT', 'ONA'])
    dom_brl = sum(lic(reales[p], TASA_DOMINIO, dom_booked.get(p, 0)) for p in ['VUELOS', 'HT', 'ONA'])
    it_brl = sum(lic(reales[p], TASA_IT, it_booked.get(p, 0)) for p in ['VUELOS', 'HT', 'ONA'])

    marca_usd = marca_brl / tc_brl
    dom_usd = dom_brl / tc_brl
    it_usd = it_brl / tc_brl

    # ── RFC (Referencia de Clientes) — Split TR / ES ──
    # ⚠️ ALERTA: Los costos INT del RFC dependen de la clasificación RUBRO por
    # departamento (COST OF REVENUE / SALES & MKTG / TECHNOLOGY / GENERAL) que
    # proviene del procesamiento intermedio del Excel y NO existe en PL_actualizado.xlsx.
    # Solución definitiva: PL_actualizado.xlsx debería incorporar el nivel de detalle
    # departamental para poder calcularlos automáticamente desde el GL.
    #
    # Los INGRESOS INT sí se calculan dinámicamente del GL:
    #   - 49102 por RespCargo (601=TR, 611=ES) + revenue de 3eros (100% ES)
    # Los COSTOS INT se hardcodean del Segmentación TP (Seg G33/H33).
    # Llave de split: TR=30.29% / ES=69.71% (fija por contrato).

    # ── Costos INT hardcodeados del procesamiento intermedio del Excel ──
    # Origen: Segmentación TP row 33, columnas G (HT) y H (ONA)
    # ⚠️ VERIFICAR CADA PERÍODO
    _RFC_SEG_BR = {
        'G33': 5138887.479141165,    # Costos HT INT (Seg row 33 col G)
        'H33': 13994412.242357673,   # Costos ONA INT (Seg row 33 col H)
    }
    KEY_TR = 0.3029   # Llave split TR (fija por contrato)
    KEY_ES = 0.6971   # Llave split ES (fija por contrato)

    # ── Ingresos INT (dinámicos del GL) ──
    # NR INT (no-IC, sin fraude) + fraude INT + 49102 IC por RespCargo
    def int_sum(df, prod):
        return df[(df['_prod'] == prod) & (df['_di'] == 'INT')][col_bal].sum()

    nr_ht_int = int_sum(df_nr_nf, 'HT')
    nr_ona_int = int_sum(df_nr_nf, 'ONA')
    fr_ht_int = int_sum(df_fraudes, 'HT')
    fr_ona_int = int_sum(df_fraudes, 'ONA')

    # 49102 (IC) por RespCargo 601=TR, 611=ES
    col_rc = next((c for c in df_br.columns if 'RespCargo' in c or 'Resp_Cargo' in c), None)
    if col_rc is None:
        df_br['_resp_cargo'] = df_br['Combinacion Contable'].str.split('.').str[3]
        col_rc = '_resp_cargo'
    c49102 = df_br[(df_br['_acct'] == '49102') & is_ic]
    rc601_ht = c49102[(c49102[col_rc] == '601') & (c49102['_prod'] == 'HT')][col_bal].sum()
    rc601_ona = c49102[(c49102[col_rc] == '601') & (c49102['_prod'] == 'ONA')][col_bal].sum()
    rc611_ht = c49102[(c49102[col_rc] == '611') & (c49102['_prod'] == 'HT')][col_bal].sum()
    rc611_ona = c49102[(c49102[col_rc] == '611') & (c49102['_prod'] == 'ONA')][col_bal].sum()

    # G31/H31 = Total revenue INT (incluye NR non-IC + fraude + 49102)
    G31 = -(nr_ht_int + rc601_ht + rc611_ht + fr_ht_int)
    H31 = -(nr_ona_int + rc601_ona + rc611_ona + fr_ona_int)

    # Revenue de 3eros (= total INT - 49102 TR - 49102 ES) → 100% España
    rev_3ros_ht = G31 + rc601_ht + rc611_ht
    rev_3ros_ona = H31 + rc601_ona + rc611_ona

    # Split ingresos por destino
    tr_ht_rev = -rc601_ht
    es_ht_rev = -rc611_ht + rev_3ros_ht
    tr_ona_rev = -rc601_ona
    es_ona_rev = -rc611_ona + rev_3ros_ona

    # Split costos por llave contractual
    G33 = _RFC_SEG_BR['G33']
    H33 = _RFC_SEG_BR['H33']
    tr_ht_cost = G33 * KEY_TR
    es_ht_cost = G33 * KEY_ES
    tr_ona_cost = H33 * KEY_TR
    es_ona_cost = H33 * KEY_ES

    # RFC = (Ingreso - Costo × 1.105) × -1
    rfc_tr_ht = (tr_ht_rev - tr_ht_cost * MARKUP_RFC) * -1
    rfc_tr_ona = (tr_ona_rev - tr_ona_cost * MARKUP_RFC) * -1
    rfc_es_ht = (es_ht_rev - es_ht_cost * MARKUP_RFC) * -1
    rfc_es_ona = (es_ona_rev - es_ona_cost * MARKUP_RFC) * -1

    rfc_tr_brl = rfc_tr_ht + rfc_tr_ona
    rfc_es_brl = rfc_es_ht + rfc_es_ona
    rfc_tr_usd = rfc_tr_brl / tc_brl
    rfc_es_usd = rfc_es_brl / tc_brl

    return {
        'marca_usd': marca_usd,
        'dominio_usd': dom_usd,
        'it_usd': it_usd,
        'rfc_tr_usd': rfc_tr_usd,
        'rfc_es_usd': rfc_es_usd,
        'marca_brl': marca_brl,
        'dominio_brl': dom_brl,
        'it_brl': it_brl,
        'rfc_tr_brl': rfc_tr_brl,
        'rfc_es_brl': rfc_es_brl,
        '_vuelos_local': vuelos_local,
        '_ht_local': ht_local,
        '_ona_local': ona_local,
        '_tc_brl': tc_brl,
        '_mes_base': mes_base,
        '_corp_nr': corp_nr,
        '_rfc_seg_br': _RFC_SEG_BR,
    }


def calcular_mexico_hoja_llave(df_balance_raw, llaves_ordenes, tc_mxn, mes_base=10):
    """
    Calcula los 4 valores de Hoja Llave México (entidad 103) en USD:
      - Licencia de Marca   (2.5%, Prestador: Travel Reservations)
      - Licencia de IT      (5.0%, Prestador: Despegar.com.ar)
      - Referencia de Clientes (9.5% markup)
      - Soporte Local (calculado aparte, no en esta función)

    ⚠️ ALERTA CLAVE: El RFC de México usa una clasificación RUBRO por departamento
    (COST OF REVENUE / SALES & MKTG / TECHNOLOGY / GENERAL) que proviene del
    procesamiento intermedio del Excel y NO existe en PL_actualizado.xlsx.
    PL_actualizado.xlsx debería incorporar el nivel de detalle departamental.
    Los costos agregados (H29_MX, I29_MX, J37_MX) se hardcodean del Excel y
    deben verificarse cada período.
    """

    TASA_MARCA = 0.025
    TASA_IT    = 0.05
    MARKUP_RFC = 1.095

    # ── Segmentación RFC hardcodeados del procesamiento intermedio del Excel ──
    # ⚠️ VERIFICAR CADA PERÍODO — estos 4 valores dependen de la clasificación RUBRO
    # departamental que solo existe en el Excel (no derivable del GL crudo).
    # Solución definitiva: PL_actualizado.xlsx debería incorporar el nivel de
    # detalle departamental para poder calcularlos automáticamente desde el GL.
    # Origen: Segmentación TP rows 32/34, columnas G (HT Ext) y H (ONA Ext)
    _RFC_SEG_MX = {
        # Período Feb-2026 (mes_base=2)
        'G32': 13405547.76,      # Ingresos HT INT (Seg row 32 col G)
        'H32': 11828615.40,      # Ingresos ONA INT (Seg row 32 col H)
        'G34': 17353811.64,      # Gastos HT INT (Seg row 34 col G)
        'H34': 19238974.43,      # Gastos ONA INT (Seg row 34 col H)
    }

    # PROD_MAP y DI_INT se usan desde las variables globales (cargadas del template)
    FRAUDE_CTAS = {'53300', '53305'}

    col_bal = next((c for c in df_balance_raw.columns if 'Saldo Final' in c or 'Balance' in c),
                   df_balance_raw.columns[-1])
    col_pl2 = next((c for c in df_balance_raw.columns if 'Nivel 2' in c), None)

    # ── Filtrar entidad 103 ──
    df_mx = df_balance_raw[df_balance_raw['Entity_Flex'] == '103'].copy()
    if df_mx.empty:
        print("⚠️  México: sin registros para entidad 103 en el balance")
        return None

    df_mx['_prod_code'] = pd.to_numeric(df_mx['Product_Flex'], errors='coerce')
    df_mx['_prod'] = df_mx['_prod_code'].map(
        lambda x: PROD_MAP.get(int(x), 'OTRO') if pd.notna(x) else 'OTRO'
    )
    df_mx['_di'] = df_mx['DomInt_Flex'].map(lambda x: 'INT' if x in DI_INT else 'LOCAL')
    if 'Account_Num' in df_mx.columns:
        df_mx['_acct'] = df_mx['Account_Num'].apply(lambda x: str(int(x)) if pd.notna(x) else '')
    else:
        df_mx['_acct'] = df_mx.iloc[:, 2].astype(str).str.strip()

    is_ic = (df_mx[col_pl2] == 'Intercompany Transactions') if col_pl2 \
        else pd.Series(False, index=df_mx.index)
    df_ni = df_mx[~is_ic]

    # Merge PL
    df_pl_local = cargar_pl(RUTA_RUBROS)
    df_pl_local['Cuenta'] = df_pl_local['Cuenta'].astype(str).str.strip()
    df_ni = df_ni.copy()
    df_ni['_acct'] = df_ni['_acct'].astype(str).str.strip()
    df_ni = df_ni.merge(df_pl_local[['Cuenta', 'PL_Totalizador']],
                        left_on='_acct', right_on='Cuenta', how='left')

    # ── NET REVENUE (sin fraudes) ──
    nr_mask = (df_ni['PL_Totalizador'] == 'Net Revenues') & (~df_ni['_acct'].isin(FRAUDE_CTAS))
    df_nr_nf = df_ni[nr_mask].copy()

    # Fraudes
    df_fraudes = df_ni[df_ni['_acct'].isin(FRAUDE_CTAS)].copy()

    # ── Ingresos Locales (base licencias) ──
    # México: VUELOS y VP todo LOCAL, HT/ONA solo NOT INT, Corp NR = 0
    def local_sum(df, prod):
        if prod in ('VUELOS', 'VUELOS PAQUETES'):
            return df[df['_prod'] == prod][col_bal].sum()
        return df[(df['_prod'] == prod) & (df['_di'] != 'INT')][col_bal].sum()

    vuelos_local = -(local_sum(df_nr_nf, 'VUELOS') + local_sum(df_fraudes, 'VUELOS'))
    ht_local = -(local_sum(df_nr_nf, 'HT') + local_sum(df_fraudes, 'HT'))
    ona_raw = -(local_sum(df_nr_nf, 'ONA') + local_sum(df_fraudes, 'ONA'))
    vp_local = -(local_sum(df_nr_nf, 'VUELOS PAQUETES') + local_sum(df_fraudes, 'VUELOS PAQUETES'))
    ona_local = ona_raw + vp_local  # ONA incluye VP

    # ── Booked licencias ──
    # 70104 = Marca, 70105 = IT (por producto, sin split subcuenta)
    booked_70104 = df_mx[(df_mx['_acct'] == '70104') & is_ic]
    marca_booked = {'VUELOS': 0, 'HT': 0, 'ONA': 0}
    for _, row in booked_70104.iterrows():
        prod = PROD_MAP.get(int(row['_prod_code']) if pd.notna(row['_prod_code']) else -1, 'OTRO')
        if prod in marca_booked:
            marca_booked[prod] += row[col_bal]

    booked_70105 = df_mx[df_mx['_acct'] == '70105']
    it_booked = {'VUELOS': 0, 'HT': 0, 'ONA': 0}
    for _, row in booked_70105.iterrows():
        prod = PROD_MAP.get(int(row['_prod_code']) if pd.notna(row['_prod_code']) else -1, 'OTRO')
        if prod in it_booked:
            it_booked[prod] += row[col_bal]

    # ── Fórmula licencias: (Real × tasa - Booked) + Real × tasa / mes_base ──
    def lic(real_val, tasa, booked_val):
        r = real_val * tasa
        return (r - booked_val) + r / mes_base

    reales = {'VUELOS': vuelos_local, 'HT': ht_local, 'ONA': ona_local}

    marca_mxn = sum(lic(reales[p], TASA_MARCA, marca_booked.get(p, 0)) for p in ['VUELOS', 'HT', 'ONA'])
    it_mxn = sum(lic(reales[p], TASA_IT, it_booked.get(p, 0)) for p in ['VUELOS', 'HT', 'ONA'])

    marca_usd = marca_mxn / tc_mxn
    it_usd = it_mxn / tc_mxn

    # ── RFC (Referencia de Clientes) ──
    # Hardcodeado del procesamiento intermedio del Excel (Segmentación TP)
    # Fórmula: G33 = (G32 - G34 × 1.095) × -1 (para HT y ONA)
    # G32/H32 = Ingresos INT, G34/H34 = Gastos INT
    # Estos 4 valores dependen de la clasificación RUBRO departamental
    # que NO existe en PL_actualizado.xlsx → hardcodeados

    G32 = _RFC_SEG_MX['G32']  # Rev HT INT
    H32 = _RFC_SEG_MX['H32']  # Rev ONA INT
    G34 = _RFC_SEG_MX['G34']  # Costs HT INT
    H34 = _RFC_SEG_MX['H34']  # Costs ONA INT

    G33 = (G32 - G34 * MARKUP_RFC) * -1  # RFC HT MXN
    H33 = (H32 - H34 * MARKUP_RFC) * -1  # RFC ONA MXN
    rfc_total_mxn = G33 + H33

    rfc_usd = -rfc_total_mxn / tc_mxn

    return {
        'Licencia_Marca': marca_usd,
        'Licencia_IT': it_usd,
        'RFC': rfc_usd,
        '_marca_mxn': marca_mxn,
        '_it_mxn': it_mxn,
        '_rfc_mxn': rfc_total_mxn,
        '_rfc_ht_mxn': G33,
        '_rfc_ona_mxn': H33,
        '_vuelos_local': vuelos_local,
        '_ht_local': ht_local,
        '_ona_local': ona_local,
        '_tc_mxn': tc_mxn,
        '_mes_base': mes_base,
        '_rfc_seg_hardcoded': True,
        '_rfc_seg_mx': _RFC_SEG_MX,
    }


def calcular_soporte_local_mexico(df_balance_raw, tc_mxn, mes_base=2):
    """
    Calcula el Soporte Local de México (entidad 103 → 401 Viajes Beda).

    Fórmula:
      Total = (Salarios + Gastos Indirectos) × llave_NR_401 × (1 + markup)
      Facturar = (Total - Booked) + Total / mes_base

    ⚠️ ALERTA: Salarios, Gastos Indirectos y Llave NR dependen del procesamiento
    intermedio del Excel (clasificación RUBRO departamental + pivot de Salarios).
    PL_actualizado.xlsx debería incorporar el nivel de detalle departamental.
    El Booked (49100 RC=401) sí se calcula dinámicamente del GL.
    """

    MARKUP_SOP = 0.08  # 8%

    # ── Hardcodes del procesamiento intermedio del Excel ──
    # ⚠️ VERIFICAR CADA PERÍODO
    _SOP_LOCAL_MX = {
        'salarios': 6490529.93,         # Pivot Salarios-Gtos Ind, SALARIOS total
        'gastos_ind': 4183633.38,       # Pivot Salarios-Gtos Ind, GASTO INDIRECTO total
        'llave_401': 0.06642941552956726,  # NR 401 / (NR 103 + NR 401)
    }

    col_bal = next((c for c in df_balance_raw.columns if 'Saldo Final' in c or 'Balance' in c),
                   df_balance_raw.columns[-1])
    col_pl2 = next((c for c in df_balance_raw.columns if 'Nivel 2' in c), None)

    df_mx = df_balance_raw[df_balance_raw['Entity_Flex'] == '103'].copy()
    if df_mx.empty:
        return None

    if 'Account_Num' in df_mx.columns:
        df_mx['_acct'] = df_mx['Account_Num'].apply(lambda x: str(int(x)) if pd.notna(x) else '')
    else:
        df_mx['_acct'] = df_mx.iloc[:, 2].astype(str).str.strip()

    is_ic = (df_mx[col_pl2] == 'Intercompany Transactions') if col_pl2 \
        else pd.Series(False, index=df_mx.index)

    # Booked: cuenta 49100 "Ingresos Interco - Servicio Soporte Local MX", RespCargo=401
    df_mx['_rc'] = df_mx['Combinacion Contable'].str.split('.').str[3]
    c49100_401 = df_mx[(df_mx['_acct'] == '49100') & is_ic & (df_mx['_rc'] == '401')]
    booked_401 = -c49100_401[col_bal].sum()

    # Cálculo
    sal = _SOP_LOCAL_MX['salarios']
    gi = _SOP_LOCAL_MX['gastos_ind']
    llave = _SOP_LOCAL_MX['llave_401']

    total_gastos_401 = (sal + gi) * llave
    total_con_markup = total_gastos_401 * (1 + MARKUP_SOP)
    mes_estimado = total_con_markup / mes_base
    monto_mxn = total_con_markup - booked_401 + mes_estimado
    monto_usd = monto_mxn / tc_mxn

    return {
        'monto_usd': monto_usd,
        'monto_mxn': monto_mxn,
        '_total_gastos_401': total_gastos_401,
        '_total_con_markup': total_con_markup,
        '_booked_401': booked_401,
        '_sop_local_mx': _SOP_LOCAL_MX,
    }


def calcular_colombia_hoja_llave(df_balance_raw, llaves_ordenes, tc_cop, mes_base=2):
    """
    Calcula los 3 valores de Hoja Llave Colombia (entidad 116) en USD:
      - Licencia de Marca (2.5%)
      - Licencia de IT (5.0%)
      - Referencia de Clientes (9.5% markup)

    ⚠️ ALERTA: El RFC usa clasificación RUBRO departamental del Excel.
    Los costos INT se hardcodean. Ingresos INT sí son dinámicos del GL.
    """
    TASA_MARCA = 0.025
    TASA_IT    = 0.05
    MARKUP_RFC = 1.095

    # ── Costos INT hardcodeados del procesamiento intermedio del Excel ──
    # Origen: Segmentación row 33, cols G (HT INT) y H (ONA INT)
    # ⚠️ VERIFICAR CADA PERÍODO
    _RFC_SEG_CO = {
        'G33': 1209941274.8762202,   # Costos HT INT
        'H33': 1940153501.1214962,   # Costos ONA INT
    }

    # DI_INT se usa desde la variable global (cargada del template)
    FRAUDE_CTAS = {'53300', '53305'}

    col_bal = next((c for c in df_balance_raw.columns if 'Saldo Final' in c or 'Balance' in c),
                   df_balance_raw.columns[-1])
    col_pl2 = next((c for c in df_balance_raw.columns if 'Nivel 2' in c), None)

    df_co = df_balance_raw[df_balance_raw['Entity_Flex'] == '116'].copy()
    if df_co.empty:
        return None

    df_co['_prod_code'] = pd.to_numeric(df_co['Product_Flex'], errors='coerce')
    df_co['_prod'] = df_co['_prod_code'].map(
        lambda x: PROD_MAP.get(int(x), 'OTRO') if pd.notna(x) else 'OTRO'
    )
    df_co['_di'] = df_co['DomInt_Flex'].map(lambda x: 'INT' if x in DI_INT else 'LOCAL')
    if 'Account_Num' in df_co.columns:
        df_co['_acct'] = df_co['Account_Num'].apply(lambda x: str(int(x)) if pd.notna(x) else '')
    else:
        df_co['_acct'] = df_co.iloc[:, 2].astype(str).str.strip()

    is_ic = (df_co[col_pl2] == 'Intercompany Transactions') if col_pl2 \
        else pd.Series(False, index=df_co.index)
    df_ni = df_co[~is_ic]

    df_pl_local = cargar_pl(RUTA_RUBROS)
    df_pl_local['Cuenta'] = df_pl_local['Cuenta'].astype(str).str.strip()
    df_ni = df_ni.copy()
    df_ni['_acct'] = df_ni['_acct'].astype(str).str.strip()
    df_ni = df_ni.merge(df_pl_local[['Cuenta', 'PL_Totalizador']],
                        left_on='_acct', right_on='Cuenta', how='left')

    # NR
    nr_mask = (df_ni['PL_Totalizador'] == 'Net Revenues') & (~df_ni['_acct'].isin(FRAUDE_CTAS))
    df_nr_nf = df_ni[nr_mask].copy()
    df_fraudes = df_ni[df_ni['_acct'].isin(FRAUDE_CTAS)].copy()

    def local_sum(df, prod):
        if prod in ('VUELOS', 'VUELOS PAQUETES'):
            return df[df['_prod'] == prod][col_bal].sum()
        return df[(df['_prod'] == prod) & (df['_di'] != 'INT')][col_bal].sum()

    vuelos = -(local_sum(df_nr_nf, 'VUELOS') + local_sum(df_fraudes, 'VUELOS'))
    ht = -(local_sum(df_nr_nf, 'HT') + local_sum(df_fraudes, 'HT'))
    ona_raw = -(local_sum(df_nr_nf, 'ONA') + local_sum(df_fraudes, 'ONA'))
    vp = -(local_sum(df_nr_nf, 'VUELOS PAQUETES') + local_sum(df_fraudes, 'VUELOS PAQUETES'))
    ona = ona_raw + vp

    # Booked
    booked_70104 = df_co[(df_co['_acct'] == '70104') & is_ic]
    marca_booked = {'VUELOS': 0, 'HT': 0, 'ONA': 0}
    for _, row in booked_70104.iterrows():
        prod = PROD_MAP.get(int(row['_prod_code']) if pd.notna(row['_prod_code']) else -1, 'OTRO')
        if prod in marca_booked:
            marca_booked[prod] += row[col_bal]

    booked_70105 = df_co[df_co['_acct'] == '70105']
    it_booked = {'VUELOS': 0, 'HT': 0, 'ONA': 0}
    for _, row in booked_70105.iterrows():
        prod = PROD_MAP.get(int(row['_prod_code']) if pd.notna(row['_prod_code']) else -1, 'OTRO')
        if prod in it_booked:
            it_booked[prod] += row[col_bal]

    def lic(real_val, tasa, booked_val):
        r = real_val * tasa
        return (r - booked_val) + r / mes_base

    reales = {'VUELOS': vuelos, 'HT': ht, 'ONA': ona}
    marca_cop = sum(lic(reales[p], TASA_MARCA, marca_booked.get(p, 0)) for p in ['VUELOS', 'HT', 'ONA'])
    it_cop = sum(lic(reales[p], TASA_IT, it_booked.get(p, 0)) for p in ['VUELOS', 'HT', 'ONA'])

    # RFC: ingresos dinámicos, costos hardcodeados
    def int_sum(df, prod):
        return df[(df['_prod'] == prod) & (df['_di'] == 'INT')][col_bal].sum()

    c49102 = df_co[(df_co['_acct'] == '49102') & is_ic]
    nr_ht_int = int_sum(df_nr_nf, 'HT') + c49102[c49102['_prod'] == 'HT'][col_bal].sum()
    nr_ona_int = int_sum(df_nr_nf, 'ONA') + c49102[c49102['_prod'] == 'ONA'][col_bal].sum()
    fr_ht_int = int_sum(df_fraudes, 'HT')
    fr_ona_int = int_sum(df_fraudes, 'ONA')

    G31 = -(nr_ht_int + fr_ht_int)
    H31 = -(nr_ona_int + fr_ona_int)

    G33 = _RFC_SEG_CO['G33']
    H33 = _RFC_SEG_CO['H33']
    rfc_ht = (G31 - G33 * MARKUP_RFC) * -1
    rfc_ona = (H31 - H33 * MARKUP_RFC) * -1
    rfc_cop = -(rfc_ht + rfc_ona)

    return {
        'Licencia_Marca': marca_cop / tc_cop,
        'Licencia_IT': it_cop / tc_cop,
        'RFC': rfc_cop / tc_cop,
        '_marca_cop': marca_cop,
        '_it_cop': it_cop,
        '_rfc_cop': rfc_cop,
        '_tc_cop': tc_cop,
        '_rfc_seg_co': _RFC_SEG_CO,
    }


def calcular_call_center_colombia(df_balance_raw, tc_cop, mes_base=2):
    """
    Calcula el Servicio de Call Center Colombia (entidad 116 → Travel Reservations).

    ⚠️ ALERTA: Los costos totales (excl Colombia DOM) dependen de matrices de
    distribución por país/producto y clasificación departamental del Excel.
    El Booked (49107) sí se calcula dinámicamente del GL.
    """

    # ── Hardcodes del procesamiento intermedio del Excel ──
    # Origen: Facturación row 11, cols J/K/L (costos excl CO DOM + markup)
    # ⚠️ VERIFICAR CADA PERÍODO
    _CC_COSTS = {
        'vuelos': 1803642612.2748654,   # J11
        'ht': 922951713.7408545,         # K11
        'ona': 1154345598.5769606,       # L11
    }

    col_bal = next((c for c in df_balance_raw.columns if 'Saldo Final' in c or 'Balance' in c),
                   df_balance_raw.columns[-1])
    col_pl2 = next((c for c in df_balance_raw.columns if 'Nivel 2' in c), None)

    df_co = df_balance_raw[df_balance_raw['Entity_Flex'] == '116'].copy()
    if df_co.empty:
        return None

    if 'Account_Num' in df_co.columns:
        df_co['_acct'] = df_co['Account_Num'].apply(lambda x: str(int(x)) if pd.notna(x) else '')
    else:
        df_co['_acct'] = df_co.iloc[:, 2].astype(str).str.strip()

    is_ic = (df_co[col_pl2] == 'Intercompany Transactions') if col_pl2 \
        else pd.Series(False, index=df_co.index)

    # Booked: cuenta 49107 por producto
    # Usa PROD_MAP global (cargado del template)
    c49107 = df_co[(df_co['_acct'] == '49107') & is_ic]
    booked = {'vuelos': 0, 'ht': 0, 'ona': 0}
    for _, row in c49107.iterrows():
        pc = pd.to_numeric(row.get('Product_Flex', 0), errors='coerce')
        p = PROD_MAP.get(int(pc), 'OTRO') if pd.notna(pc) else 'OTRO'
        if p == 'VUELOS':
            booked['vuelos'] += row[col_bal]
        elif p == 'HT':
            booked['ht'] += row[col_bal]
        elif p == 'ONA':
            booked['ona'] += row[col_bal]

    booked_v = -booked['vuelos']
    booked_h = -booked['ht']
    booked_o = -booked['ona']

    # Fórmula: (Cost - Booked) + Cost / mes_base
    cost_v = _CC_COSTS['vuelos']
    cost_h = _CC_COSTS['ht']
    cost_o = _CC_COSTS['ona']

    fact_v = (cost_v - booked_v) + cost_v / mes_base
    fact_h = (cost_h - booked_h) + cost_h / mes_base
    fact_o = (cost_o - booked_o) + cost_o / mes_base
    fact_total_cop = fact_v + fact_h + fact_o

    return {
        'monto_usd': fact_total_cop / tc_cop,
        'monto_cop': fact_total_cop,
        '_booked_v': booked_v,
        '_booked_h': booked_h,
        '_booked_o': booked_o,
        '_cc_costs': _CC_COSTS,
    }


# ============================================================================
# ALERTAS DE DATOS — detectan problemas ANTES de calcular
# ============================================================================

# DomInt_Flex conocidos por entidad (todo lo que no esté aquí = potencial error)
KNOWN_DI_FLEX = {
    '103': {'0001': 'DOM', '0002': 'INT', '0102': 'INT', '0112': 'INT'},           # México
    '105': {'0001': 'DOM', '0002': 'INT', '0112': 'INT'},                         # USA
    '108': {'0001': 'DOM', '0002': 'INT', '0112': 'INT', '0102': 'INT'},           # Chile
    '113': {'0001': 'DOM', '0002': 'INT', '0112': 'INT'},                         # Ecuador
    '115': {'0001': 'DOM', '0101': 'DOM', '0111': 'DOM',
            '0002': 'INT', '0102': 'INT', '0112': 'INT'},                         # Perú
    '601': {'0001': 'DOM', '0002': 'INT', '0102': 'INT', '0112': 'INT'},           # TR
    '611': {'0001': 'DOM', '0002': 'INT', '0112': 'INT'},                         # España
    '104': {'0001': 'DOM', '0101': 'DOM', '0111': 'DOM',
            '0002': 'INT', '0102': 'INT', '0112': 'INT'},                         # Brasil
    '116': {'0001': 'DOM', '0101': 'DOM', '0111': 'DOM',
            '0002': 'INT', '0102': 'INT', '0112': 'INT'},                         # Colombia
}

# DomInt_Flex que se clasifican como N/A intencionalmente (no son errores)
KNOWN_NA_DI = {'0000', '0300', '0101', '0111'}  # 0101/0111 son N/A excepto Perú

# Monedas requeridas por entidad para conversión a USD
MONEDAS_REQUERIDAS = {
    '103': 'MXN',   # México
    '108': 'CLP',   # Chile
    '115': 'PEN',   # Perú
    '611': 'EUR',   # España
    '104': 'BRL',   # Brasil
    # USA (105), Ecuador (113) y TR (601) no requieren TC (moneda USD nativa)
}

# Umbral de materialidad para alertas de cuentas sin PL (en moneda local)
UMBRAL_MATERIALIDAD = 10000


def ejecutar_alertas(df_balance, df_pl, tc_dict):
    """
    Ejecuta 4 alertas sobre los datos del balance ANTES de calcular.
    Imprime warnings visibles para que el usuario los revise.
    Retorna True si hay alertas críticas, False si todo OK.
    """
    if df_balance is None:
        return False

    alertas = []
    col_bal = next((c for c in df_balance.columns if 'Saldo Final' in c or 'Balance' in c),
                   df_balance.columns[-1])

    # ── ALERTA 1: Product codes no mapeados con saldo ──────────────────────
    entidades_impl = {'103', '104', '105', '108', '113', '115', '116', '601', '611'}
    df_con_prod = df_balance[df_balance['Product_Flex'].notna()].copy()
    df_con_prod['_prod_code'] = pd.to_numeric(df_con_prod['Product_Flex'], errors='coerce')
    df_con_prod['_prod'] = df_con_prod['_prod_code'].map(
        lambda x: PROD_MAP.get(int(x), 'OTRO') if pd.notna(x) else 'OTRO'
    )
    no_mapeados = df_con_prod[df_con_prod['_prod'] == 'OTRO']
    if not no_mapeados.empty:
        saldo_total = no_mapeados[col_bal].sum()
        if abs(saldo_total) > 0:
            codes = no_mapeados['_prod_code'].dropna().astype(int).unique()
            detalle = []
            afecta_impl = False
            for code in sorted(codes):
                mask = no_mapeados['_prod_code'] == code
                saldo_code = no_mapeados.loc[mask, col_bal].sum()
                if abs(saldo_code) > 1000:  # filtro materialidad
                    entities = sorted(no_mapeados.loc[mask, 'Entity_Flex'].unique())
                    entities_impl = [e for e in entities if e in entidades_impl]
                    cuentas = no_mapeados.loc[mask, 'Account_Num'].dropna().astype(int).unique()[:3]
                    ctas_str = ', '.join(str(int(c)) for c in cuentas)
                    marca = '  ⚠️ AFECTA PAÍS IMPLEMENTADO' if entities_impl else ''
                    if entities_impl:
                        afecta_impl = True
                    detalle.append(f"    Product {code:>5d}  │  Entidades: {entities}  │  Saldo: {saldo_code:>15,.2f}  │  Cuentas: {ctas_str}{marca}")
            if detalle:
                nivel = 'CRÍTICA' if afecta_impl else 'INFO'
                accion = ('⚠️ Hay codes que afectan países implementados (USA/CHL/ECU/PER). Agregar a PROD_MAP.'
                          if afecta_impl else
                          'Estos codes solo afectan países NO implementados (AR/MX/BR/CO). Informativos por ahora.')
                alertas.append((nivel, 'PRODUCT CODES NO MAPEADOS EN PROD_MAP',
                               '\n'.join(detalle), accion))

    # ── ALERTA 2: DomInt_Flex no reconocidos con saldo ─────────────────────
    entidades_activas = {'104', '105', '108', '113', '115', '116', '601', '611'}
    for entity in entidades_activas:
        df_ent = df_balance[df_balance['Entity_Flex'] == entity]
        if df_ent.empty:
            continue
        known = set(KNOWN_DI_FLEX.get(entity, {}).keys()) | KNOWN_NA_DI
        di_values = df_ent['DomInt_Flex'].dropna().unique()
        desconocidos = [v for v in di_values if v not in known]
        if desconocidos:
            detalle = []
            for di in sorted(desconocidos):
                mask = df_ent['DomInt_Flex'] == di
                saldo = df_ent.loc[mask, col_bal].sum()
                if abs(saldo) > 0:
                    n_filas = mask.sum()
                    detalle.append(f"    Entity {entity}  │  DomInt '{di}'  │  Saldo: {saldo:>15,.2f}  │  {n_filas} filas")
            if detalle:
                nombre = ENTITY_CONFIG.get(int(entity), {}).get('nombre', entity)
                alertas.append(('MEDIA', f'DomInt_Flex DESCONOCIDO EN {nombre.upper()}',
                               '\n'.join(detalle),
                               'Estos valores caen a N/A. Verificar si deberían ser DOM o INT.'))

    # ── ALERTA 3: Cuentas GL con saldo material sin PL_Totalizador ────────
    if df_pl is not None:
        cuentas_pl = set(df_pl['Cuenta'].dropna().astype(int).unique())
        df_con_cuenta = df_balance[df_balance['Account_Num'].notna()].copy()
        df_con_cuenta['_acct'] = df_con_cuenta['Account_Num'].astype(int)
        sin_pl = df_con_cuenta[~df_con_cuenta['_acct'].isin(cuentas_pl)]
        if not sin_pl.empty:
            resumen = sin_pl.groupby('_acct')[col_bal].sum()
            materiales = resumen[resumen.abs() > UMBRAL_MATERIALIDAD].sort_values(key=abs, ascending=False)
            if not materiales.empty:
                col_desc = next((c for c in df_balance.columns if 'Descripci' in c or 'Descri' in c), None)
                detalle = []

                sin_mapear = []
                for cuenta, saldo in materiales.items():
                    acct = int(cuenta)
                    desc = ''
                    if col_desc:
                        desc_vals = df_balance.loc[df_balance['Account_Num'] == cuenta, col_desc].dropna().unique()
                        if len(desc_vals) > 0:
                            desc = str(desc_vals[0])[:40]
                    ents = sorted(sin_pl.loc[sin_pl['_acct'] == cuenta, 'Entity_Flex'].unique())
                    sin_mapear.append((acct, saldo, desc, ents))

                if sin_mapear:
                    detalle.append("    ┌─ SIN MAPEAR EN PL")
                    for acct, saldo, desc, ents in sin_mapear:
                        detalle.append(f"    │  Cuenta {acct:>5d}  │  {desc:<40}  │  Saldo: {saldo:>15,.2f}  │  Entidades: {ents}")
                    detalle.append("    └─")

                n_sin_mapear = len(sin_mapear)
                accion = f'{n_sin_mapear} cuenta(s) sin PL_Totalizador.'
                alertas.append(('INFO', 'CUENTAS SIN PL_TOTALIZADOR',
                               '\n'.join(detalle), accion))

    # ── ALERTA 4: Moneda requerida ausente en TC EPM ──────────────────────
    for entity, moneda in MONEDAS_REQUERIDAS.items():
        if moneda not in tc_dict:
            nombre = ENTITY_CONFIG.get(int(entity), {}).get('nombre', entity)
            alertas.append(('CRÍTICA', f'MONEDA {moneda} NO ENCONTRADA EN TC EPM',
                           f'    Entidad {entity} ({nombre}) requiere TC {moneda}/USD para convertir resultados.',
                           f'El archivo TC no contiene {moneda}. La entidad no podrá calcularse.'))

    # ── Imprimir alertas ──────────────────────────────────────────────────
    if alertas:
        print("\n" + "=" * 80)
        print("  ⚠️  ALERTAS DE DATOS")
        print("=" * 80)
        for i, (nivel, titulo, detalle, accion) in enumerate(alertas, 1):
            icono = '🔴' if nivel == 'CRÍTICA' else ('🟡' if nivel == 'MEDIA' else 'ℹ️')
            print(f"\n  {icono} ALERTA {i} [{nivel}]: {titulo}")
            print(f"  {'─' * 70}")
            print(detalle)
            print(f"\n  → Acción: {accion}")
        print(f"\n  {'=' * 70}")
        print(f"  Total: {len(alertas)} alerta(s) detectada(s)")
        print("=" * 80)
    else:
        print("\n  ✅ Sin alertas — todos los product codes, DomInt, cuentas y monedas OK")

    return len(alertas) > 0


# ============================================================================
# PROCESO COMPLETO
# ============================================================================

def ejecutar_proceso_completo():
    print("\n" + "=" * 80)
    print("  🚀 SISTEMA DE FACTURACIÓN INTERCOMPANY")
    print("=" * 80)

    mes_base = leer_periodo_balance(RUTA_BALANCE) or 10
    print(f"\n  📅 Período del balance: mes {mes_base} (base YTD para catch-up)\n")

    # ── Cargar PL desde template_rubros.xlsx (hoja PL_Totalizador) ───────────
    try:
        df_pl_template = pd.read_excel(RUTA_RUBROS, sheet_name='PL_Totalizador')
        df_pl_template['Cuenta'] = pd.to_numeric(df_pl_template['Cuenta'], errors='coerce')
        print(f"  ✅ PL cargado desde template: {len(df_pl_template)} cuentas")
    except Exception as e:
        print(f"  ❌ PL: no se pudo cargar desde template: {e}")
        df_pl_template = pd.DataFrame(columns=['Cuenta', 'PL_Totalizador'])

    # ── Cargar glosario desde template_rubros.xlsx (hoja Productos) ────────
    df_glosario = cargar_glosario(RUTA_RUBROS)
    if df_glosario is not None:
        print(f"  ✅ Glosario cargado desde template: {len(df_glosario)} productos")
    else:
        print(f"  ❌ Glosario: no se pudo cargar desde template")

    # ── Construir PROD_MAP desde template (hoja Productos) ─────────────────
    global PROD_MAP
    if df_glosario is not None and 'COD' in df_glosario.columns and 'SEGMENTACION' in df_glosario.columns:
        _tp = df_glosario.dropna(subset=['COD', 'SEGMENTACION'])
        PROD_MAP = dict(zip(_tp['COD'].astype(int), _tp['SEGMENTACION'].astype(str)))
        print(f"  ✅ PROD_MAP cargado desde template: {len(PROD_MAP)} códigos de producto")
    else:
        PROD_MAP = {}
        print(f"  ❌ PROD_MAP: no se pudo cargar desde template")

    # ── Construir DI_MAP desde template_rubros.xlsx (hoja DOM_INT) ─────────
    global DI_MAP, DI_INT
    try:
        _tmpl_di = pd.read_excel(RUTA_RUBROS, sheet_name='DOM_INT')
        _td = _tmpl_di.dropna(subset=['Negocio2', 'DOM/INT'])
        DI_MAP = dict(zip(_td['Negocio2'].astype(int), _td['DOM/INT'].astype(str)))
        DI_INT = {str(k).zfill(4) for k, v in DI_MAP.items() if v == 'INT'}
        print(f"  ✅ DI_MAP cargado desde template: {len(DI_MAP)} códigos DOM/INT")
        print(f"     DI_INT (códigos INT): {sorted(DI_INT)}")
    except Exception as e:
        DI_MAP = {1: 'DOM', 2: 'INT', 101: 'DOM', 102: 'INT', 111: 'DOM', 112: 'INT', 121: 'DOM', 300: 'Iniciativas'}
        DI_INT = {'0002', '0102', '0112'}
        print(f"  ⚠️  DI_MAP: no se pudo cargar desde template, usando default")

    # ── Construir RFC_COST_RUBROS desde template_rubros.xlsx (hoja RFC_Cost_Rubros) ──
    global RFC_COST_RUBROS
    try:
        _tmpl_rfc = pd.read_excel(RUTA_RUBROS, sheet_name='RFC_Cost_Rubros')
        RFC_COST_RUBROS = set(_tmpl_rfc['PL Totalizador'].dropna().unique())
        print(f"  ✅ RFC_COST_RUBROS cargado desde template: {RFC_COST_RUBROS}")
    except Exception as e:
        RFC_COST_RUBROS = {
            'Total Cost of Revenue', 'Total Technology and content',
            'Total Sales & Marketing', 'Total General and Administrative',
        }
        print(f"  ⚠️  RFC_COST_RUBROS: no se pudo cargar desde template, usando default")

    df_balance  = procesar_balance(RUTA_BALANCE)
    df_ordenes, df_si = procesar_ordenes(RUTA_ORDENES)
    df_adi      = procesar_adi(RUTA_ADI)

    # ── LLAVES ────────────────────────────────────────────────────────────────
    print("\n" + "=" * 80)
    print("  🔑 LLAVES DE DISTRIBUCIÓN")
    print("=" * 80)

    llaves_ordenes = calcular_llaves_ordenes(df_si) if df_si is not None else {}
    llaves_adi     = calcular_llaves_adi(df_adi) if df_adi is not None else {}
    llaves_revenue = calcular_llaves_revenue(df_balance, df_glosario)

    # ── TIPOS DE CAMBIO ───────────────────────────────────────────────────────
    tc = cargar_tc(RUTA_TC)
    tc_clp = tc.get('CLP', 946.38)
    tc_pen = tc.get('PEN', 3.385)
    tc_eur = tc.get('EUR', 0.8613)
    tc_ars = tc.get('ARS', 1399.52)
    tc_brl = tc.get('BRL', 5.4005)
    tc_mxn = tc.get('MXN', 18.6843)
    tc_cop = tc.get('COP', 3705.0)

    # ── PL (necesario para alertas al final) ─────────────────────────────────
    df_pl = cargar_pl(RUTA_RUBROS)

    # ── ECUADOR ───────────────────────────────────────────────────────────────
    print("\n" + "=" * 80)
    print("  🇪🇨 CÁLCULO ECUADOR - HOJA LLAVE")
    print("=" * 80)

    ecuador = calcular_ecuador_hoja_llave(df_balance, llaves_ordenes, mes_base=mes_base) \
              if (df_balance is not None and llaves_ordenes) else None

    if ecuador:
        print()
        print(f"  {'Licencia de Marca:':<30}  ${ecuador['Licencia_Marca']:>14,.2f} USD")
        print(f"  {'Licencia de IT:':<30}  ${ecuador['Licencia_IT']:>14,.2f} USD")
        print(f"  {'Referencia de Clientes:':<30}  ${ecuador['RFC']:>14,.2f} USD")
        print()
    else:
        print("  ⚠️  Ecuador - Hoja Llave: Sin datos")

    # ── CHILE ─────────────────────────────────────────────────────────────────
    print("\n" + "=" * 80)
    print("  🇨🇱 CÁLCULO CHILE - HOJA LLAVE")
    print("=" * 80)

    chile = calcular_chile_hoja_llave(df_balance, llaves_ordenes, tc_clp=tc_clp, mes_base=mes_base) \
            if (df_balance is not None and llaves_ordenes) else None

    if chile:
        print()
        print(f"  {'Licencia de Marca:':<30}  ${chile['Licencia_Marca']:>14,.2f} USD")
        print(f"  {'Licencia de IT:':<30}  ${chile['Licencia_IT']:>14,.2f} USD")
        print(f"  {'Referencia de Clientes:':<30}  ${chile['RFC']:>14,.2f} USD")
        print()
    else:
        print("  ⚠️  Chile - Hoja Llave: Sin datos")

    # ── PERÚ ──────────────────────────────────────────────────────────────────
    print("\n" + "=" * 80)
    print("  🇵🇪 CÁLCULO PERÚ - HOJA LLAVE")
    print("=" * 80)

    peru = calcular_peru_hoja_llave(df_balance, llaves_ordenes, tc_pen=tc_pen, mes_base=mes_base) \
           if (df_balance is not None and llaves_ordenes) else None

    if peru:
        print()
        print(f"  {'Licencia de Marca:':<30}  ${peru['Licencia_Marca']:>14,.2f} USD")
        print(f"  {'Licencia de IT:':<30}  ${peru['Licencia_IT']:>14,.2f} USD")
        print(f"  {'Referencia de Clientes:':<30}  ${peru['RFC']:>14,.2f} USD")
        print()
    else:
        print("  ⚠️  Perú - Hoja Llave: Sin datos")

    # ── USA ───────────────────────────────────────────────────────────────────
    print("\n" + "=" * 80)
    print("  🇺🇸 CÁLCULO USA - HOJA LLAVE")
    print("=" * 80)

    usa = calcular_usa_hoja_llave(df_balance, llaves_ordenes, mes_base=mes_base) \
          if (df_balance is not None and llaves_ordenes) else None

    if usa:
        print()
        print(f"  {'Licencia de Marca:':<30}  ${usa['Licencia_Marca']:>14,.2f} USD")
        print(f"  {'Licencia de IT:':<30}  ${usa['Licencia_IT']:>14,.2f} USD")
        print(f"  {'Referencia de Clientes:':<30}  ${usa['RFC']:>14,.2f} USD")
        print(f"  {'Servicios de Hosting:':<30}  ${usa['Hosting']:>14,.2f} USD")
        print()
        print("  ┌─────────────────────────────────────────────────────────────────────")
        print("  │ ⚠️  HARDCODES USA — M5 Corporate (base RFC + Hosting)")
        print("  │ QUÉ: Cuentas 80102 y 64341 incluidas manualmente en M5,")
        print("  │   cuenta 80101 excluida manualmente de M5.")
        print("  │   PL Nivel 1 no las clasifica correctamente para el RFC de USA.")
        print("  │ SOLUCIÓN: Ajustar PL_actualizado.xlsx para que PL Nivel 1")
        print("  │   refleje correctamente qué cuentas van a M5 CORPORATE.")
        print("  └─────────────────────────────────────────────────────────────────────")
        print()
    else:
        print("  ⚠️  USA - Hoja Llave: Sin datos")

    # ── TR (TRAVEL RESERVATION) — Entidad 601 ─────────────────────────────
    print("\n" + "=" * 80)
    print("  🇺🇾 CÁLCULO TR (TRAVEL RESERVATION) - HOJA LLAVE  [Entidad 601]")
    print("=" * 80)

    tr = calcular_tr_hoja_llave(df_balance, mes_base=mes_base) \
         if df_balance is not None else None

    if tr:
        _fc_usd_tr = tr['Licencia_IT']
        _fc_ars_tr = _fc_usd_tr * tc_ars
        print()
        print("  Travel Reservation - Licencia IT")
        print()
        print(f"  {'Prestador':<30}  {'Prestatario':<30}  {'FC ARS':>16}  {'FC USD':>16}")
        print(f"  {'-'*30}  {'-'*30}  {'-'*16}  {'-'*16}")
        print(f"  {'Despegar.com.ar':<30}  {'Travel Reservations S.R.L':<30}  {_fc_ars_tr:>16,.2f}  {_fc_usd_tr:>16,.2f}")
        print()
        print(f"  {'TC ARS':>62}  {tc_ars:>16,.2f}")
        print()
        print("  ⚠️  Delta -$1.58 vs Excel — El Excel mete 14 filas de cuenta 52500")
        print("     como Net Revenue (clasificación manual). La cuenta 52500 es Finance.")
        print()
    else:
        print("  ⚠️  TR - Hoja Llave: Sin datos")

    # ── ESPAÑA — Entidad 611 ────────────────────────────────────────────────
    print("\n" + "=" * 80)
    print("  🇪🇸 CÁLCULO ESPAÑA - HOJA LLAVE  [Entidad 611]")
    print("=" * 80)

    espana = calcular_espana_hoja_llave(df_balance, tc_eur=tc_eur, mes_base=mes_base) \
             if df_balance is not None else None

    if espana:
        _fc_usd_it = espana['Licencia_IT']
        _fc_ars_it = _fc_usd_it * tc_ars
        _fc_usd_mk = espana['Licencia_Marca']
        _fc_eur_mk = espana['Licencia_Marca_EUR']

        # Bloque 1: Licencia IT
        print()
        print("  Despegar España - Licencia IT")
        print()
        print(f"  {'Prestador':<30}  {'Prestatario':<30}  {'FC ARS':>16}  {'FC USD':>16}")
        print(f"  {'-'*30}  {'-'*30}  {'-'*16}  {'-'*16}")
        print(f"  {'Despegar.com.ar':<30}  {'Despegar España':<30}  {_fc_ars_it:>16,.2f}  {_fc_usd_it:>16,.2f}")
        print()

        # Bloque 2: Licencia de Marcas
        print("  Despegar España - Licencia de Marcas")
        print()
        print(f"  {'Prestador':<30}  {'Prestatario':<30}  {'FC EUR':>16}  {'FC USD':>16}")
        print(f"  {'-'*30}  {'-'*30}  {'-'*16}  {'-'*16}")
        print(f"  {'Travel Reservations':<30}  {'Despegar España':<30}  {_fc_eur_mk:>16,.2f}  {_fc_usd_mk:>16,.2f}")
        print()
        print(f"  {'EUR/USD':>62}  {tc_eur:>16,.10f}")
        print(f"  {'TC ARS':>62}  {tc_ars:>16,.2f}")
        print()
        print("  ⚠️  Delta IT -$19.34 / Marca -$9.67 vs Excel — El Excel mete cuenta")
        print("     52500 completa como Net Revenue para España. La cuenta 52500 es Finance.")
        print()
    else:
        print("  ⚠️  España - Hoja Llave: Sin datos")

    # ── BRASIL ────────────────────────────────────────────────────────────────
    print("\n" + "=" * 80)
    print("  🇧🇷 CÁLCULO BRASIL - HOJA LLAVE  [Entidad 104]")
    print("=" * 80)

    brasil = calcular_brasil_hoja_llave(df_balance, llaves_ordenes, tc_brl=tc_brl, mes_base=mes_base) \
             if (df_balance is not None and llaves_ordenes) else None

    if brasil:
        print()
        print(f"  {'Concepto':<30}  {'A facturar - USD':>20}")
        print(f"  {'-'*30}  {'-'*20}")
        print(f"  {'Licencia de Marca':<30}  {brasil['marca_usd']:>20,.2f}")
        print(f"  {'Licencia de Dominio':<30}  {brasil['dominio_usd']:>20,.2f}")
        print(f"  {'Licencia de IT':<30}  {brasil['it_usd']:>20,.2f}")
        print(f"  {'Referencia de Clientes TR':<30}  {brasil['rfc_tr_usd']:>20,.2f}")
        print(f"  {'Referencia de Clientes ES':<30}  {brasil['rfc_es_usd']:>20,.2f}")
        print()
        print(f"  {'TC BRL/USD':>50}  {tc_brl:>20,.4f}")
        print(f"  {'TC ARS':>50}  {tc_ars:>20,.2f}")
        print()
        _seg = brasil.get('_rfc_seg_br', {})
        print("  ┌─────────────────────────────────────────────────────────────────────")
        print("  │ ⚠️  HARDCODES BRASIL — RFC TR y RFC ES")
        print("  │")
        print("  │ QUÉ: Los costos INT del RFC están hardcodeados (no se calculan del GL)")
        print(f"  │   G33 = {_seg.get('G33', 0):>15,.2f} BRL  (costos HT INT)")
        print(f"  │   H33 = {_seg.get('H33', 0):>15,.2f} BRL  (costos ONA INT)")
        print(f"  │   Split: TR = 30.29% / ES = 69.71% (contractual)")
        print("  │   Los ingresos INT sí se calculan del GL (49102 por RespCargo)")
        print("  │")
        print("  │ POR QUÉ: El Excel clasifica costos por RUBRO departamental")
        print("  │   (COST OF REVENUE / SALES & MKTG / TECHNOLOGY / GENERAL)")
        print("  │   usando la columna Dpto del GL. PL_actualizado.xlsx no tiene")
        print("  │   ese nivel de detalle, solo clasifica por cuenta.")
        print("  │")
        print("  │ SOLUCIÓN: Agregar columna Dpto al PL_actualizado.xlsx para que")
        print("  │   el código pueda derivar el RUBRO automáticamente desde el GL.")
        print("  │   Mientras tanto, actualizar G33/H33 cada período desde la")
        print("  │   Segmentación TP del Excel (row 33, cols G y H).")
        print("  └─────────────────────────────────────────────────────────────────────")
        print()
    else:
        print("  ⚠️  Brasil - Hoja Llave: Sin datos")

    # ── MÉXICO ───────────────────────────────────────────────────────────────
    print("\n" + "=" * 80)
    print("  🇲🇽 CÁLCULO MÉXICO - HOJA LLAVE  [Entidad 103]")
    print("=" * 80)

    mexico = calcular_mexico_hoja_llave(df_balance, llaves_ordenes, tc_mxn=tc_mxn, mes_base=mes_base) \
             if (df_balance is not None and llaves_ordenes) else None

    if mexico:
        print()
        print(f"  {'Concepto':<30}  {'A facturar - USD':>20}")
        print(f"  {'-'*30}  {'-'*20}")
        print(f"  {'Licencia de Marca':<30}  {mexico['Licencia_Marca']:>20,.2f}")
        print(f"  {'Licencia de IT':<30}  {mexico['Licencia_IT']:>20,.2f}")
        print(f"  {'Referencia de Clientes':<30}  {mexico['RFC']:>20,.2f}")
        print()
        print(f"  {'TC MXN/USD':>50}  {tc_mxn:>20,.4f}")
        print(f"  {'TC ARS':>50}  {tc_ars:>20,.2f}")
        print()
        _seg_mx = mexico.get('_rfc_seg_mx', {})
        print("  ┌─────────────────────────────────────────────────────────────────────")
        print("  │ ⚠️  HARDCODES MÉXICO — Referencia de Clientes")
        print("  │")
        print("  │ QUÉ: Los 4 valores del RFC están hardcodeados (no se calculan del GL)")
        print(f"  │   G32 = {_seg_mx.get('G32', 0):>15,.2f} MXN  (ingresos HT INT)")
        print(f"  │   H32 = {_seg_mx.get('H32', 0):>15,.2f} MXN  (ingresos ONA INT)")
        print(f"  │   G34 = {_seg_mx.get('G34', 0):>15,.2f} MXN  (costos HT INT)")
        print(f"  │   H34 = {_seg_mx.get('H34', 0):>15,.2f} MXN  (costos ONA INT)")
        print("  │")
        print("  │ POR QUÉ: El Excel clasifica costos e ingresos por RUBRO departamental")
        print("  │   (COST OF REVENUE / SALES & MKTG / TECHNOLOGY / GENERAL)")
        print("  │   usando la columna Dpto del GL. PL_actualizado.xlsx no tiene")
        print("  │   ese nivel de detalle, solo clasifica por cuenta.")
        print("  │")
        print("  │ SOLUCIÓN: Agregar columna Dpto al PL_actualizado.xlsx para que")
        print("  │   el código pueda derivar el RUBRO automáticamente desde el GL.")
        print("  │   Mientras tanto, actualizar G32/H32/G34/H34 cada período desde")
        print("  │   la Segmentación TP del Excel (rows 32/34, cols G y H).")
        print("  └─────────────────────────────────────────────────────────────────────")
        print()

        # ── Soporte Local ──
        sop = calcular_soporte_local_mexico(df_balance, tc_mxn=tc_mxn, mes_base=mes_base) \
              if df_balance is not None else None
        if sop:
            print(f"  {'Soporte Local (401-VB)':<30}  {sop['monto_usd']:>20,.2f}")
            print()
            _sop_data = sop.get('_sop_local_mx', {})
            print("  ┌─────────────────────────────────────────────────────────────────────")
            print("  │ ⚠️  HARDCODES MÉXICO — Soporte Local")
            print("  │")
            print("  │ QUÉ: Salarios, Gastos Indirectos y Llave NR están hardcodeados")
            print(f"  │   Salarios total    = {_sop_data.get('salarios', 0):>15,.2f} MXN")
            print(f"  │   Gastos Ind. total = {_sop_data.get('gastos_ind', 0):>15,.2f} MXN")
            print(f"  │   Llave NR 401      = {_sop_data.get('llave_401', 0)*100:.4f}%")
            print(f"  │   Booked 49100      = {sop['_booked_401']:>15,.2f} MXN  ← DINÁMICO del GL")
            print("  │")
            print("  │ POR QUÉ: Mismo problema RUBRO departamental. La llave NR depende")
            print("  │   del NR de entidad 103 clasificado por RUBRO del Excel.")
            print("  │   Los salarios/gastos vienen de un pivot por Dpto del Excel.")
            print("  │")
            print("  │ SOLUCIÓN: Misma que RFC — agregar Dpto al PL_actualizado.xlsx.")
            print("  │   Mientras tanto, actualizar estos 3 valores cada período desde")
            print("  │   el Excel de Servicio de Soporte Local (hoja Cálculo).")
            print("  └─────────────────────────────────────────────────────────────────────")
            print()
    else:
        print("  ⚠️  México - Hoja Llave: Sin datos")

    # ── COLOMBIA ─────────────────────────────────────────────────────────────
    print("\n" + "=" * 80)
    print("  🇨🇴 CÁLCULO COLOMBIA - HOJA LLAVE  [Entidad 116]")
    print("=" * 80)

    colombia = calcular_colombia_hoja_llave(df_balance, llaves_ordenes, tc_cop=tc_cop, mes_base=mes_base) \
               if (df_balance is not None and llaves_ordenes) else None

    if colombia:
        print()
        print(f"  {'Concepto':<30}  {'A facturar - USD':>20}")
        print(f"  {'-'*30}  {'-'*20}")
        print(f"  {'Licencia de Marca':<30}  {colombia['Licencia_Marca']:>20,.2f}")
        print(f"  {'Licencia de IT':<30}  {colombia['Licencia_IT']:>20,.2f}")
        print(f"  {'Referencia de Clientes':<30}  {colombia['RFC']:>20,.2f}")
        print()
        print(f"  {'TC COP/USD':>50}  {tc_cop:>20,.4f}")
        print(f"  {'TC ARS':>50}  {tc_ars:>20,.2f}")
        print()

        _seg_co = colombia.get('_rfc_seg_co', {})
        print("  ┌─────────────────────────────────────────────────────────────────────")
        print("  │ ⚠️  HARDCODES COLOMBIA — Referencia de Clientes")
        print("  │")
        print("  │ QUÉ: Los costos INT del RFC están hardcodeados (no se calculan del GL)")
        print(f"  │   G33 = {_seg_co.get('G33', 0):>18,.2f} COP  (costos HT INT)")
        print(f"  │   H33 = {_seg_co.get('H33', 0):>18,.2f} COP  (costos ONA INT)")
        print("  │   Los ingresos INT sí se calculan del GL (49102 + NR INT + fraudes)")
        print("  │")
        print("  │ POR QUÉ: Clasificación RUBRO departamental del Excel.")
        print("  │   PL_actualizado.xlsx no tiene ese nivel de detalle.")
        print("  │")
        print("  │ SOLUCIÓN: Agregar Dpto al PL_actualizado.xlsx.")
        print("  │   Mientras tanto, actualizar G33/H33 cada período desde la")
        print("  │   Segmentación TP del Excel (row 33, cols G y H).")
        print("  └─────────────────────────────────────────────────────────────────────")
        print()

        # ── Call Center ──
        cc = calcular_call_center_colombia(df_balance, tc_cop=tc_cop, mes_base=mes_base) \
             if df_balance is not None else None
        if cc:
            print(f"  {'Servicios de Call Center':<30}  {cc['monto_usd']:>20,.2f}")
            print()
            _ccc = cc.get('_cc_costs', {})
            print("  ┌─────────────────────────────────────────────────────────────────────")
            print("  │ ⚠️  HARDCODES COLOMBIA — Call Center")
            print("  │")
            print("  │ QUÉ: Costos totales excl Colombia DOM están hardcodeados")
            print(f"  │   Vuelos = {_ccc.get('vuelos', 0):>18,.2f} COP")
            print(f"  │   HT     = {_ccc.get('ht', 0):>18,.2f} COP")
            print(f"  │   ONA    = {_ccc.get('ona', 0):>18,.2f} COP")
            print(f"  │   Booked 49107 = dinámico del GL")
            print("  │")
            print("  │ POR QUÉ: Costos distribuidos por matrices país×producto×departamento")
            print("  │   del procesamiento intermedio del Excel. Incluye markup 5%.")
            print("  │")
            print("  │ SOLUCIÓN: Los costos por Dpto sí se pueden sacar del GL (RC=300/310,")
            print("  │   PL1='Fulfilment Center Fees', Dptos Call Center). Lo que falta")
            print("  │   replicar es la distribución por país×producto usando las 4 matrices")
            print("  │   (Front/Back/Otros/Customer Strategy) + gastos indirectos.")
            print("  │   Agregar Dpto al PL permitiría calcular todo dinámicamente.")
            print("  └─────────────────────────────────────────────────────────────────────")
            print()
    else:
        print("  ⚠️  Colombia - Hoja Llave: Sin datos")

    # ── ALERTAS DE DATOS (al final para no interrumpir los resultados) ──────
    ejecutar_alertas(df_balance, df_pl, tc)

    print("=" * 80)
    print("\n\nPresiona ENTER para cerrar...")
    input()

if __name__ == "__main__":
    ejecutar_proceso_completo()
