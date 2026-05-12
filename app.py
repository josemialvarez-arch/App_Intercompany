"""
FACTURACIÓN INTERCOMPANY — App v2
"""
import os
import streamlit as st
import pandas as pd
import tempfile
import io

from motor import (
    leer_periodo_balance, cargar_glosario, cargar_pl, cargar_tc,
    procesar_balance, procesar_ordenes, procesar_adi,
    calcular_llaves_ordenes, calcular_llaves_adi, calcular_llaves_revenue,
    calcular_ecuador_hoja_llave, calcular_chile_hoja_llave,
    calcular_peru_hoja_llave, calcular_usa_hoja_llave,
    calcular_tr_hoja_llave, calcular_espana_hoja_llave,
)

st.set_page_config(page_title="Facturación Intercompany", page_icon="📊", layout="wide")

# ============================================================================
# CSS
# ============================================================================
st.markdown("""
<style>
    .block-container { padding-top: 1.5rem; max-width: 1200px; }

    /* Header */
    .app-header {
        background: linear-gradient(135deg, #0f2027 0%, #203a43 50%, #2c5364 100%);
        padding: 2.5rem 2rem; border-radius: 16px; color: white; margin-bottom: 2rem;
        position: relative; overflow: hidden;
    }
    .app-header::after {
        content: ''; position: absolute; top: -40%; right: -10%;
        width: 300px; height: 300px;
        background: radial-gradient(circle, rgba(255,255,255,0.06) 0%, transparent 70%);
        border-radius: 50%;
    }
    .app-header h1 { color: white; font-size: 2.2rem; font-weight: 700; margin: 0; position: relative; z-index: 1; }
    .app-header .subtitle { color: #94b8d0; font-size: 0.95rem; margin-top: 0.5rem; position: relative; z-index: 1; }
    .app-header .badge {
        display: inline-block; background: rgba(255,255,255,0.15);
        padding: 0.25rem 0.75rem; border-radius: 20px; font-size: 0.75rem;
        color: #cde; margin-top: 0.8rem; position: relative; z-index: 1;
    }

    /* Steps */
    .steps-container { display: flex; gap: 0; margin: 1.5rem 0; align-items: center; }
    .step-circle {
        width: 48px; height: 48px; border-radius: 50%;
        display: flex; align-items: center; justify-content: center;
        font-weight: 700; font-size: 1.1rem; flex-shrink: 0;
        transition: all 0.3s ease;
    }
    .step-circle.done { background: #10b981; color: white; }
    .step-circle.active { background: #2563eb; color: white; box-shadow: 0 0 0 4px rgba(37,99,235,0.25); transform: scale(1.1); }
    .step-circle.pending { background: #e5e7eb; color: #9ca3af; }
    .step-line { flex: 1; height: 3px; margin: 0 0.4rem; transition: background 0.3s; }
    .step-line.done { background: #10b981; }
    .step-line.active { background: linear-gradient(90deg, #10b981, #2563eb); }
    .step-line.pending { background: #e5e7eb; }
    .step-label { text-align: center; font-size: 0.72rem; color: #6b7280; margin-top: 0.4rem; width: 48px; }
    .step-label.active { color: #2563eb; font-weight: 600; }
    .step-label.done { color: #10b981; }
    .step-item { display: flex; flex-direction: column; align-items: center; }

    /* Upload */
    .upload-card {
        background: white; border: 2px solid #2563eb; border-radius: 14px;
        padding: 1.5rem; margin: 1.5rem 0; box-shadow: 0 4px 16px rgba(37,99,235,0.1);
    }
    .upload-card h3 { margin: 0 0 0.3rem 0; color: #1e3a5f; font-size: 1.15rem; }
    .upload-card p { margin: 0; color: #6b7280; font-size: 0.85rem; }

    /* Country cards */
    div[data-testid="stVerticalBlockBorderWrapper"] {
        border-radius: 14px !important;
        transition: transform 0.2s, box-shadow 0.2s;
    }
    div[data-testid="stVerticalBlockBorderWrapper"]:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 20px rgba(0,0,0,0.08) !important;
    }

    /* Metric styling */
    div[data-testid="stMetricValue"] { font-size: 1.3rem !important; }

    /* Download buttons */
    .stDownloadButton > button {
        border-radius: 10px !important;
        font-weight: 500 !important;
        padding: 0.6rem 1.2rem !important;
    }

    /* Footer */
    .app-footer {
        text-align: center; color: #9ca3af; font-size: 0.75rem;
        padding: 2rem 0 1rem 0; border-top: 1px solid #f0f0f0; margin-top: 2rem;
    }
    .app-footer a { color: #6b7280; text-decoration: none; }

    /* Status badges */
    .status-ok { color: #10b981; font-weight: 600; }
    .status-warn { color: #f59e0b; font-weight: 600; }
    .status-err { color: #ef4444; font-weight: 600; }
</style>
""", unsafe_allow_html=True)

# ============================================================================
# FUNCIONES
# ============================================================================

def generar_bajada_trabajada(rutas, r):
    """Genera la Bajada Trabajada: GL + columnas derivadas"""
    gl = pd.read_csv(rutas["gl"], sep=';', skiprows=10, encoding='latin-1')
    col_si = 'Saldo Inicial'; col_ap = 'Actv Periodo'
    col_b = next(c for c in gl.columns if 'Saldo' in c and 'Final' in c)
    for c in [col_si, col_ap, col_b]:
        gl[c] = pd.to_numeric(gl[c].astype(str).str.replace(',','.'), errors='coerce')

    parts = gl['Combinacion Contable'].str.split('.', expand=True)
    gl['Cuenta_num'] = pd.to_numeric(gl.iloc[:,2], errors='coerce').astype('Int64')
    gl['_acct'] = gl['Cuenta_num'].astype(str)

    # Todo se lee del template maestro
    pl = pd.read_excel(rutas["rubros"], sheet_name='PL_Totalizador')
    pl['Cuenta'] = pl['Cuenta'].astype(str).str.strip()
    pl = pl.drop_duplicates(subset='Cuenta', keep='first')
    gl = gl.merge(pl[['Cuenta','PL_Totalizador']], left_on='_acct', right_on='Cuenta', how='left', suffixes=('','_pl'))

    tmpl_dpto = pd.read_excel(rutas["rubros"], sheet_name='Dpto_RUBRO')
    tmpl_exc = pd.read_excel(rutas["rubros"], sheet_name='Excepciones_Cuenta')
    tmpl_emp = pd.read_excel(rutas["rubros"], sheet_name='Futuro2_Empresa')
    tmpl_prod = pd.read_excel(rutas["rubros"], sheet_name='Productos')
    tmpl_di = pd.read_excel(rutas["rubros"], sheet_name='DOM_INT')

    dr_dict = dict(zip(tmpl_dpto['Dpto'].astype(int).astype(str).str.zfill(4), tmpl_dpto['RUBRO']))
    dd_dict = dict(zip(tmpl_dpto['Dpto'].astype(int).astype(str).str.zfill(4), tmpl_dpto['Depto']))
    exc_dict = dict(zip(tmpl_exc['Cuenta'].dropna().astype(int).astype(str), tmpl_exc['RUBRO'].dropna()))
    emp_dict = dict(zip(tmpl_emp['Codigo'].astype(int).astype(str), tmpl_emp['Empresa']))

    # PL_TO_RUBRO desde template (hoja PL_RUBRO)
    tmpl_plr = pd.read_excel(rutas["rubros"], sheet_name='PL_RUBRO')
    PL_TO_RUBRO = dict(zip(tmpl_plr['PL Totalizador'].dropna(), tmpl_plr['RUBRO'].dropna()))
    RUBRO_TO_PLT = {v: k for k, v in PL_TO_RUBRO.items()}

    # PROD_MAP desde template (hoja Productos)
    _tp = tmpl_prod.dropna(subset=['COD', 'SEGMENTACION'])
    PROD_MAP = dict(zip(_tp['COD'].astype(int).astype(str).str.zfill(4), _tp['SEGMENTACION']))

    # DI_MAP desde template (hoja DOM_INT)
    _td = tmpl_di.dropna(subset=['Negocio2', 'DOM/INT'])
    DI_MAP = dict(zip(_td['Negocio2'].astype(int), _td['DOM/INT']))

    # Derive RUBRO
    gl['RUBRO'] = gl['PL_Totalizador'].map(PL_TO_RUBRO).fillna('RESULTADO')
    m60 = (gl['Cuenta_num'] >= 60000) & (gl['Cuenta_num'] < 70000)
    dpto_col = parts[5]
    dr_mapped = dpto_col[m60].map(dr_dict)
    gl.loc[m60 & dr_mapped.notna(), 'RUBRO'] = dr_mapped[dr_mapped.notna()]
    for a, rb in exc_dict.items():
        gl.loc[gl['_acct'] == a, 'RUBRO'] = rb

    # Derive PL Totalizador (override for 60xxx)
    gl['MY_PLT'] = gl['PL_Totalizador']
    rubro_plt = gl.loc[m60, 'RUBRO'].map(RUBRO_TO_PLT)
    gl.loc[m60 & rubro_plt.notna(), 'MY_PLT'] = rubro_plt[rubro_plt.notna()]

    di_num = pd.to_numeric(parts[9], errors='coerce').astype('Int64')
    f2_num = pd.to_numeric(parts[10].fillna('0'), errors='coerce').fillna(0).astype(int).astype(str)

    output = pd.DataFrame({
        'Mayor22': gl['Mayor'], 'Moneda de Mayor22': gl['Moneda de Mayor'],
        'Cuenta22': gl.iloc[:,2], 'Descripcion22': gl['Descripcion'],
        'Combinacion Contable': gl['Combinacion Contable'],
        'Mapeo PL Nivel 1-': gl['Mapeo PL Nivel 1'], 'P&L2': gl['Mapeo PL Nivel 2'],
        'Clasificacion  Contable-': gl['Clasificacion  Contable'], 'Rubro-': gl['Rubro'],
        'Nota-': gl['Nota'], 'Subcuenta-': gl['Subcuenta'], 'Nombre subcuenta-': gl['Nombre subcuenta'],
        'Divisa-': gl['Divisa'], 'Saldo Inicial-': gl[col_si], 'Actv Periodo-': gl[col_ap],
        'Balance_Final': gl[col_b],
        'E.Legal Cta': parts[0], 'Cuenta': gl['Cuenta_num'],
        'Cuenta_Descripcion': gl['Descripcion'],
        'Mapeo PL Nivel 1': gl['Mapeo PL Nivel 1'],
        'RUBRO': gl['RUBRO'], 'Subcuenta': parts[2], 'Resp_Cargo': parts[3],
        'Prod': parts[4], 'PRODUCTO': parts[4].map(PROD_MAP).fillna(''),
        'Dpto': parts[5], 'Depto': parts[5].map(dd_dict).fillna(''),
        'Canal': parts[6], 'Tipo Cliente': parts[7], 'Tipo Cobro': parts[8],
        'Negocio': parts[9], 'DOM/INT': di_num.map(DI_MAP),
        'Futuro2': parts[10] if 10 in parts.columns else '',
        'Empresa': f2_num.map(emp_dict).fillna(''),
        'PL Totalizador': gl['MY_PLT'],
    })
    return output


def generar_excel_ordenes(rutas, r):
    """Genera el Excel de Órdenes con 3 hojas: Ordenes FY 2026, HT, ONA"""
    try:
        _, df_si = procesar_ordenes(rutas["ordenes"])
        df_use = df_si[df_si['Usar'] == 'SI']
    except:
        return None

    dom_vals = {'Dom', 'DOM'}
    int_vals = {'Int'}

    # Pivot by site × Producto_1 × DOM_INT (para Ordenes FY y HT)
    pivot1 = df_use.groupby(['site', 'Producto_1', 'DOM_INT'])['Bookings Emitidos'].sum()
    # Pivot by site × Producto_2 × trip_type_code (para ONA)
    pivot2 = df_use.groupby(['site', 'Producto_2', 'trip_type_code'])['Bookings Emitidos'].sum()

    def gv1(s, prod, di_set):
        total = 0
        for di in di_set:
            try: total += int(pivot1.get((s, prod, di), 0))
            except: pass
        return total

    def gv2(s, prod, trip):
        try: return int(pivot2.get((s, prod, trip), 0))
        except: return 0

    countries = ['Argentina', 'Bolivia', 'Brasil', 'Chile', 'Colombia', 'Costa Rica',
                 'Ecuador', 'El Salvador', 'España', 'Usa', 'Guatemala', 'Honduras',
                 'International', 'Mexico', 'Nicaragua', 'Panama', 'Paraguay', 'Peru',
                 'Puerto Rico', 'Republica Dominicana', 'Uruguay', 'Venezuela']
    display = {c: c for c in countries}
    display['Usa'] = 'USA'

    mes = r.get("mes_base", 2)
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        # ── HOJA 1: Ordenes FY 2026 ──
        rows_fy = []
        for s in countries:
            v = gv1(s, 'Vuelos', dom_vals) + gv1(s, 'Vuelos', int_vals)
            hd = gv1(s, 'Hoteles', dom_vals)
            hi = gv1(s, 'Hoteles', int_vals)
            od = gv1(s, 'ONA', dom_vals) + gv1(s, 'Vuelos Paquetes', dom_vals) + gv1(s, 'Vuelos Paquetes', int_vals)
            oi = gv1(s, 'ONA', int_vals)
            tot = v + hd + hi + od + oi
            pv = round(v/tot*100, 2) if tot else 0
            phd = round(hd/tot*100, 2) if tot else 0
            phi = round(hi/tot*100, 2) if tot else 0
            pod = round(od/tot*100, 2) if tot else 0
            poi = round(oi/tot*100, 2) if tot else 0
            rows_fy.append({
                'País': display[s], 'Vuelos': v, 'HT DOM': hd, 'HT INT': hi,
                'ONA DOM': od, 'ONA INT': oi, 'Total': tot,
                '': '', 'Vuelos%': f'{pv:.2f}%', 'HT DOM%': f'{phd:.2f}%', 'HT INT%': f'{phi:.2f}%',
                'ONA DOM%': f'{pod:.2f}%', 'ONA INT%': f'{poi:.2f}%', 'Total%': f'{pv+phd+phi+pod+poi:.2f}%',
                '  ': '', 'ONA+VP': od + oi,
            })
        df_fy = pd.DataFrame(rows_fy)
        df_fy.to_excel(writer, sheet_name='Ordenes FY 2026', index=False, startrow=1)

        # ── HOJA 2: HT ──
        rows_ht = []
        for s in countries:
            hd = gv1(s, 'Hoteles', dom_vals)
            hi = gv1(s, 'Hoteles', int_vals)
            tot = hd + hi
            pd_pct = round(hd/tot*100, 2) if tot else 0
            pi_pct = round(hi/tot*100, 2) if tot else 0
            rows_ht.append({
                'País': display[s], 'DOM': hd, 'INT': hi,
                '': '', 'DOM%': f'{pd_pct:.2f}%', 'INT%': f'{pi_pct:.2f}%',
            })
        df_ht = pd.DataFrame(rows_ht)
        df_ht.to_excel(writer, sheet_name='HT', index=False, startrow=1)

        # ── HOJA 3: ONA ──
        rows_ona = []
        for s in countries:
            ona_nac = gv2(s, 'ONA', 'Nac')
            ona_int = gv2(s, 'ONA', 'Int')
            vp_nac = gv2(s, 'Vuelos Paquetes', 'Nac')
            vp_int = gv2(s, 'Vuelos Paquetes', 'Int')
            ona_tot = ona_nac + ona_int
            # ONA+VP combined: DOM = ONA_Nac + VP_Nac + VP_Int, INT = ONA_Int
            comb_d = ona_nac + vp_nac + vp_int
            comb_i = ona_int
            comb_tot = comb_d + comb_i
            od_pct = round((ona_nac + vp_nac) / ona_tot * 100, 2) if ona_tot else 0
            oi_pct = round((ona_int + vp_int) / ona_tot * 100, 2) if ona_tot else 0
            cd_pct = round(comb_d / comb_tot * 100, 2) if comb_tot else 0
            ci_pct = round(comb_i / comb_tot * 100, 2) if comb_tot else 0
            rows_ona.append({
                'País': display[s], 'DOM': ona_nac + vp_nac, 'INT': ona_int + vp_int,
                '': '', 'DOM%': f'{od_pct:.2f}%', 'INT%': f'{oi_pct:.2f}%',
                '  ': '', 'VP NAC': vp_nac, 'VP INT': vp_int,
                '   ': '', 'ONA+VP DOM': comb_d, 'ONA+VP INT': comb_i,
                '    ': '', 'ONA+VP DOM%': f'{cd_pct:.2f}%', 'ONA+VP INT%': f'{ci_pct:.2f}%',
            })
        df_ona = pd.DataFrame(rows_ona)
        df_ona.to_excel(writer, sheet_name='ONA', index=False, startrow=1)

    return buffer.getvalue()


def guardar_temporal(uploaded, key):
    if uploaded is None:
        return None
    suffix = "." + uploaded.name.split(".")[-1]
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    tmp.write(uploaded.getbuffer())
    tmp.close()
    st.session_state.archivos_tmp[key] = tmp.name
    return tmp.name


def ejecutar_calculo(rutas):
    import motor
    motor.RUTA_RUBROS = rutas["rubros"]
    motor.RUTA_BALANCE = rutas["gl"]
    motor.RUTA_ORDENES = rutas["ordenes"]
    motor.RUTA_TC = rutas["tc"]
    motor.RUTA_ADI = rutas["adi"]

    # Inicializar variables globales del motor desde template
    try:
        _tp = pd.read_excel(rutas["rubros"], sheet_name='Productos')
        _tp = _tp.dropna(subset=['COD', 'SEGMENTACION'])
        motor.PROD_MAP = dict(zip(_tp['COD'].astype(int), _tp['SEGMENTACION'].astype(str)))
    except:
        motor.PROD_MAP = {}

    try:
        _td = pd.read_excel(rutas["rubros"], sheet_name='DOM_INT')
        _td = _td.dropna(subset=['Negocio2', 'DOM/INT'])
        motor.DI_MAP = dict(zip(_td['Negocio2'].astype(int), _td['DOM/INT'].astype(str)))
        motor.DI_INT = {str(k).zfill(4) for k, v in motor.DI_MAP.items() if v == 'INT'}
    except:
        motor.DI_MAP = {1: 'DOM', 2: 'INT', 101: 'DOM', 102: 'INT', 111: 'DOM', 112: 'INT', 121: 'DOM', 300: 'Iniciativas'}
        motor.DI_INT = {'0002', '0102', '0112'}

    try:
        _tr = pd.read_excel(rutas["rubros"], sheet_name='RFC_Cost_Rubros')
        motor.RFC_COST_RUBROS = set(_tr['PL Totalizador'].dropna().unique())
    except:
        motor.RFC_COST_RUBROS = {'Total Cost of Revenue', 'Total Technology and content', 'Total Sales & Marketing', 'Total General and Administrative'}

    r = {}
    mes_base = leer_periodo_balance(rutas["gl"]) or 10
    r["mes_base"] = mes_base

    df_glosario = cargar_glosario(rutas["rubros"])
    df_balance = procesar_balance(rutas["gl"])
    _, df_si = procesar_ordenes(rutas["ordenes"])
    df_adi = procesar_adi(rutas["adi"])
    tc = cargar_tc(rutas["tc"])

    r["tc"] = tc
    tc_clp = tc.get('CLP', 946.38)
    tc_pen = tc.get('PEN', 3.385)
    tc_eur = tc.get('EUR', 0.8613)

    llaves_ordenes = calcular_llaves_ordenes(df_si) if df_si is not None else {}
    llaves_adi = calcular_llaves_adi(df_adi) if df_adi is not None else {}
    r["llaves_ordenes"] = llaves_ordenes
    r["llaves_adi"] = llaves_adi

    r["ecuador"] = calcular_ecuador_hoja_llave(df_balance, llaves_ordenes, mes_base=mes_base) \
                   if (df_balance is not None and llaves_ordenes) else None
    r["chile"] = calcular_chile_hoja_llave(df_balance, llaves_ordenes, tc_clp=tc_clp, mes_base=mes_base) \
                 if (df_balance is not None and llaves_ordenes) else None
    r["peru"] = calcular_peru_hoja_llave(df_balance, llaves_ordenes, tc_pen=tc_pen, mes_base=mes_base) \
                if (df_balance is not None and llaves_ordenes) else None
    r["usa"] = calcular_usa_hoja_llave(df_balance, llaves_ordenes, mes_base=mes_base) \
               if (df_balance is not None and llaves_ordenes) else None
    r["tr"] = calcular_tr_hoja_llave(df_balance, mes_base=mes_base) \
              if df_balance is not None else None
    r["espana"] = calcular_espana_hoja_llave(df_balance, tc_eur=tc_eur, mes_base=mes_base) \
                  if df_balance is not None else None
    return r


def fmt_usd(v):
    """Formatea USD con color"""
    if v < 0:
        return f":red[**${v:,.2f}**]"
    return f":green[**${v:,.2f}**]"


def render_steps(current, total=5):
    labels = ["GL", "Órdenes", "TC", "ADI", "Template"]
    html = '<div class="steps-container">'
    for i in range(1, total + 1):
        if i < current:
            cls, icon = "done", "✓"
        elif i == current:
            cls, icon = "active", str(i)
        else:
            cls, icon = "pending", str(i)
        html += f'<div class="step-item"><div class="step-circle {cls}">{icon}</div><div class="step-label {cls}">{labels[i-1]}</div></div>'
        if i < total:
            line_cls = "done" if i < current else ("active" if i == current else "pending")
            html += f'<div class="step-line {line_cls}"></div>'
    html += '</div>'
    st.markdown(html, unsafe_allow_html=True)


def render_pais_card(flag, nombre, metricas, alert=None, sub=None):
    """Renderiza card de país usando componentes nativos de Streamlit"""
    with st.container(border=True):
        st.markdown(f"### {flag} {nombre}")
        if sub:
            st.caption(sub)

        cols = st.columns(len(metricas))
        for i, (label, value) in enumerate(metricas):
            with cols[i]:
                color = "normal" if value >= 0 else "inverse"
                st.metric(label=label, value=f"${value:,.2f}", delta=None)

        if alert:
            st.warning(alert, icon="⚠️")


def mostrar_llaves(r):
    """Muestra llaves K1-K12 una debajo de otra"""
    llaves = r.get("llaves_ordenes", {})
    if not llaves:
        return

    nombres = {
        'K1': 'Producto Global', 'K2': 'Por Entidad (Base)', 'K3': 'Sin Brasil',
        'K4': '150% (Inc Decolar + 601)', 'K5': '100% Decolar Inc',
        'K6': 'Matriz Vuelos/HT/ONA × DOM/INT', 'K7': 'DOM-INT Total por País',
        'K8': 'Productos por País (sin AR)', 'K9': 'DOM / INT por POS',
        'K10': 'Producto Detallado por POS', 'K11': 'Hoteles DOM / INT por POS',
        'K12': 'ONA DOM / INT por POS',
    }

    with st.expander("🔑 Llaves de distribución K1 — K12", expanded=False):
        keys_list = [k for k in ['K1','K2','K3','K4','K5','K6','K7','K8','K9','K10','K11','K12'] if k in llaves]

        for k in keys_list:
            data = llaves[k]
            st.markdown(f"**{k}: {nombres.get(k, k)}**")

            if isinstance(data, pd.Series):
                df_l = data.reset_index()
                df_l.columns = ['Entidad', '%']
                # Valores ya vienen como porcentaje (23.39, no 0.2339)
                df_l['%'] = df_l['%'].apply(lambda x: f"{x:.2f}%" if isinstance(x, (int, float)) else x)
                st.dataframe(df_l, use_container_width=True, hide_index=True, height=min(len(df_l)*38+38, 400))
            elif isinstance(data, pd.DataFrame):
                df_fmt = data.copy()
                for c in df_fmt.select_dtypes(include='number').columns:
                    df_fmt[c] = df_fmt[c].apply(lambda x: f"{x:.2f}%")
                st.dataframe(df_fmt, use_container_width=True, height=min(len(df_fmt)*38+38, 400))
            st.markdown("---")


def mostrar_resultados(r):
    tc = r.get("tc", {})
    tc_ars = tc.get('ARS', 1399.52)
    tc_eur = tc.get('EUR', 0.8613)
    mes = r.get("mes_base", 2)

    st.markdown(f"""
    <div class="app-header">
        <h1>📊 Resultados</h1>
        <div class="subtitle">Facturación Intercompany — Período mes {mes} (base YTD)</div>
        <div class="badge">6 mercados calculados</div>
    </div>
    """, unsafe_allow_html=True)

    # ── LLAVES ──
    mostrar_llaves(r)

    # ── PAÍSES ──
    st.markdown("---")

    paises = []

    ec = r.get("ecuador")
    if ec:
        paises.append(("🇪🇨", "Ecuador", [
            ("Marca", ec["Licencia_Marca"]),
            ("IT", ec["Licencia_IT"]),
            ("RFC", ec["RFC"]),
        ], None, None))

    cl = r.get("chile")
    if cl:
        paises.append(("🇨🇱", "Chile", [
            ("Marca", cl["Licencia_Marca"]),
            ("IT", cl["Licencia_IT"]),
            ("RFC", cl["RFC"]),
        ], None, None))

    pe = r.get("peru")
    if pe:
        paises.append(("🇵🇪", "Perú", [
            ("Marca", pe["Licencia_Marca"]),
            ("IT", pe["Licencia_IT"]),
            ("RFC", pe["RFC"]),
        ], None, None))

    us = r.get("usa")
    if us:
        paises.append(("🇺🇸", "USA", [
            ("Marca", us["Licencia_Marca"]),
            ("IT", us["Licencia_IT"]),
            ("RFC", us["RFC"]),
            ("Hosting", us["Hosting"]),
        ], None, None))

    tr_d = r.get("tr")
    if tr_d:
        fc_usd = tr_d["Licencia_IT"]
        fc_ars = fc_usd * tc_ars
        paises.append(("🇺🇾", "Travel Reservation", [
            ("IT (USD)", fc_usd),
            ("IT (ARS)", fc_ars),
        ], None, f"Despegar.com.ar → TR S.R.L | TC {tc_ars:,.2f}"))

    es = r.get("espana")
    if es:
        fc_eur_mk = es.get("Licencia_Marca_EUR", es["Licencia_Marca"] * tc_eur)
        paises.append(("🇪🇸", "España", [
            ("IT (USD)", es["Licencia_IT"]),
            ("Marca (USD)", es["Licencia_Marca"]),
            ("Marca (EUR)", fc_eur_mk),
        ], None, f"IT: Despegar.com.ar → España | Marca: TR → España | EUR/USD {tc_eur:.4f}"))

    with st.expander("🌎 Hojas Llave — 6 mercados", expanded=False):
        for flag, nombre, metricas, alert, sub in paises:
            render_pais_card(flag, nombre, metricas, alert=alert, sub=sub)

    # ── DESCARGA ──
    st.markdown("---")
    st.markdown("### 📥 Descargar resultados")

    rows = []
    for pais, key, conceptos in [
        ("Ecuador", "ecuador", [("Licencia de Marca","Licencia_Marca"),("Licencia de IT","Licencia_IT"),("Referencia de Clientes","RFC")]),
        ("Chile", "chile", [("Licencia de Marca","Licencia_Marca"),("Licencia de IT","Licencia_IT"),("Referencia de Clientes","RFC")]),
        ("Perú", "peru", [("Licencia de Marca","Licencia_Marca"),("Licencia de IT","Licencia_IT"),("Referencia de Clientes","RFC")]),
        ("USA", "usa", [("Licencia de Marca","Licencia_Marca"),("Licencia de IT","Licencia_IT"),("Referencia de Clientes","RFC"),("Servicios de Hosting","Hosting")]),
        ("TR", "tr", [("Licencia de IT","Licencia_IT")]),
        ("España", "espana", [("Licencia de IT","Licencia_IT"),("Licencia de Marca","Licencia_Marca")]),
    ]:
        d = r.get(key)
        if d:
            for label, field in conceptos:
                rows.append({"País": pais, "Concepto": label, "USD": d[field]})

    if rows:
        df_export = pd.DataFrame(rows)
        periodo = r.get("mes_base", 0)

        col_d1, col_d2 = st.columns(2)
        with col_d1:
            st.download_button(
                "📄 Descargar CSV",
                data=df_export.to_csv(index=False),
                file_name=f"facturacion_mes{periodo}.csv",
                mime="text/csv",
                use_container_width=True,
            )
        with col_d2:
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                df_export.to_excel(writer, index=False, sheet_name='Resultados')
                llaves_exp = r.get("llaves_ordenes", {})
                for k in ['K1','K2','K3','K4','K5','K6','K7','K8','K9','K10','K11','K12']:
                    if k not in llaves_exp:
                        continue
                    data = llaves_exp[k]
                    if isinstance(data, pd.Series):
                        data.reset_index().to_excel(writer, index=False, sheet_name=k)
                    elif isinstance(data, pd.DataFrame):
                        data.to_excel(writer, sheet_name=k)
            st.download_button(
                "📊 Descargar Excel",
                data=buffer.getvalue(),
                file_name=f"facturacion_mes{periodo}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        # ── Órdenes y Bajada Trabajada ──
        st.markdown("### 📥 Descargar outputs")
        col_o1, col_o2 = st.columns(2)
        with col_o1:
            ordenes_data = generar_excel_ordenes(st.session_state.archivos_tmp, r)
            if ordenes_data:
                st.download_button(
                    "📦 Descargar Órdenes",
                    data=ordenes_data,
                    file_name=f"Ordenes_GB_trabajado_mes{periodo}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
        with col_o2:
            if st.button("📋 Generar Bajada Trabajada", use_container_width=True):
                with st.spinner("Generando Bajada Trabajada (~80K filas)..."):
                    try:
                        bajada = generar_bajada_trabajada(st.session_state.archivos_tmp, r)
                        buf_baj = io.BytesIO()
                        with pd.ExcelWriter(buf_baj, engine='xlsxwriter') as writer:
                            bajada.to_excel(writer, index=False, sheet_name='Bajada Trabajada')
                            entities = sorted(bajada['E.Legal Cta'].unique())
                            for ent in entities:
                                df_ent = bajada[bajada['E.Legal Cta'] == ent]
                                sheet_name = str(ent)[:31]
                                df_ent.to_excel(writer, index=False, sheet_name=sheet_name)
                        st.session_state.bajada_data = buf_baj.getvalue()
                        st.rerun()
                    except Exception as e:
                        st.error(f"Error: {e}")
                        import traceback
                        st.code(traceback.format_exc())

        if st.session_state.get("bajada_data"):
            st.download_button(
                "⬇️ Descargar Bajada Trabajada",
                data=st.session_state.bajada_data,
                file_name=f"Bajada_Trabajada_mes{periodo}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
            st.success("Bajada Trabajada generada — lista para descargar")


# ============================================================================
# ESTADO
# ============================================================================
ARCHIVOS = [
    {"key": "gl",       "label": "Balance de Comprobación",  "step": 1, "ext": ["txt"],  "desc": "GL Oracle (.txt separado por ;)", "icon": "📄"},
    {"key": "ordenes",  "label": "Órdenes (Bookings)",       "step": 2, "ext": ["xlsx"], "desc": "Órdenes completo (.xlsx)", "icon": "📦"},
    {"key": "tc",       "label": "Tipo de Cambio EPM",       "step": 3, "ext": ["xlsx"], "desc": "TC EPM (.xlsx)", "icon": "💱"},
    {"key": "adi",      "label": "ADI Headcount",            "step": 4, "ext": ["xlsx"], "desc": "ADI Definitivo (.xlsx)", "icon": "👥"},
    {"key": "rubros",   "label": "Template Maestro",          "step": 5, "ext": ["xlsx"], "desc": "Template con PL, Productos, DOM/INT, Dptos, Empresa (.xlsx)", "icon": "🏷️"},
]

if "step" not in st.session_state:
    st.session_state.step = 1
if "archivos_tmp" not in st.session_state:
    st.session_state.archivos_tmp = {}
if "resultados" not in st.session_state:
    st.session_state.resultados = None

# ============================================================================
# PANTALLA DE CARGA
# ============================================================================
if st.session_state.resultados is None:
    st.markdown("""
    <div class="app-header">
        <h1>📊 Facturación Intercompany</h1>
        <div class="subtitle">Cargá los archivos de input para calcular llaves + hojas llave</div>
        <div class="badge">Ecuador · Chile · Perú · USA · TR · España</div>
    </div>
    """, unsafe_allow_html=True)

    current_step = st.session_state.step
    render_steps(min(current_step, 6))

    for arch in ARCHIVOS:
        s = arch["step"]
        key = arch["key"]

        if s != current_step:
            continue

        st.markdown(f"""
        <div class="upload-card">
            <h3>{arch['icon']} Step {s}: {arch['label']}</h3>
            <p>{arch['desc']}</p>
        </div>
        """, unsafe_allow_html=True)

        uploaded = st.file_uploader(
            "Subí el archivo",
            type=arch["ext"],
            key=f"upload_{key}",
            label_visibility="collapsed",
        )

        if uploaded:
            st.success(f"**{uploaded.name}** — {uploaded.size/1024:.0f} KB")

            # ── VALIDACIÓN DEL ARCHIVO ──
            validation_ok = True
            try:
                if key == "gl":
                    _df = pd.read_csv(uploaded, sep=';', skiprows=10, encoding='latin-1', nrows=5)
                    uploaded.seek(0)
                    required_gl = ['Combinacion Contable', 'Mapeo PL Nivel 1', 'Mapeo PL Nivel 2',
                                   'Mayor', 'Moneda de Mayor', 'Descripcion', 'Clasificacion  Contable',
                                   'Rubro', 'Nota', 'Subcuenta', 'Nombre subcuenta', 'Divisa']
                    saldo_cols = [c for c in _df.columns if 'Saldo' in c and 'Final' in c]
                    si_cols = [c for c in _df.columns if c == 'Saldo Inicial']
                    ap_cols = [c for c in _df.columns if c == 'Actv Periodo']
                    missing = [c for c in required_gl if c not in _df.columns]
                    if not saldo_cols: missing.append('Saldo Final')
                    if not si_cols: missing.append('Saldo Inicial')
                    if not ap_cols: missing.append('Actv Periodo')
                    if missing:
                        st.error(f"❌ GL incompleto — Faltan columnas: **{', '.join(missing)}**")
                        validation_ok = False
                    else:
                        st.success(f"GL correcto — {len(_df.columns)} columnas detectadas")

                elif key == "ordenes":
                    _df = pd.read_excel(uploaded, nrows=5)
                    uploaded.seek(0)
                    required_ord = ['site', 'brand', 'buy_type_code', 'product',
                                    'line_of_business_code', 'trip_type_code',
                                    'Bookings', 'Bookings Emitidos', 'Suma de gb']
                    missing = [c for c in required_ord if c not in _df.columns]
                    if missing:
                        st.error(f"❌ Órdenes incompleto — Faltan columnas: **{', '.join(missing)}**")
                        validation_ok = False
                    else:
                        st.success(f"Órdenes correcto — {len(_df.columns)} columnas detectadas")

                elif key == "tc":
                    _df = pd.read_excel(uploaded)
                    uploaded.seek(0)
                    content = _df.to_string().upper()
                    required_monedas = ['CLP', 'PEN', 'EUR']
                    found = [m for m in required_monedas if m in content]
                    missing = [m for m in required_monedas if m not in content]
                    if missing:
                        st.error(f"❌ TC incompleto — Faltan monedas: **{', '.join(missing)}** (necesarias para Chile, Perú, España)")
                        validation_ok = False
                    else:
                        all_monedas = [m for m in ['CLP', 'PEN', 'EUR', 'BRL', 'MXN', 'COP', 'ARS'] if m in content]
                        st.success(f"TC correcto — Monedas: {', '.join(all_monedas)}")

                elif key == "adi":
                    _df = pd.read_excel(uploaded, nrows=5)
                    uploaded.seek(0)
                    required_adi = ['COD. E.L.', 'ID', 'Estado']
                    # Flexible match for ADI columns
                    cols_str = ' '.join(str(c) for c in _df.columns)
                    missing = []
                    if not any('E.L' in str(c) for c in _df.columns): missing.append('COD. E.L.')
                    if not any('ID' in str(c).upper() for c in _df.columns): missing.append('ID')
                    if not any('Estado' in str(c) or 'STATUS' in str(c).upper() for c in _df.columns): missing.append('Estado')
                    if missing:
                        st.error(f"❌ ADI incompleto — No se detectaron: **{', '.join(missing)}**")
                        validation_ok = False
                    else:
                        st.success("ADI correcto — Columnas de entidad, ID y estado detectadas")

                elif key == "rubros":
                    _xls = pd.ExcelFile(uploaded)
                    uploaded.seek(0)
                    required_sheets = {
                        'Productos': ['COD', 'SEGMENTACION'],
                        'PL_RUBRO': ['PL Totalizador', 'RUBRO'],
                        'DOM_INT': ['Negocio2', 'DOM/INT'],
                        'Dpto_RUBRO': ['Dpto', 'RUBRO'],
                        'PL_Totalizador': ['Cuenta', 'PL_Totalizador'],
                        'Futuro2_Empresa': ['Codigo', 'Empresa'],
                    }
                    errors = []
                    for sheet, cols in required_sheets.items():
                        if sheet not in _xls.sheet_names:
                            errors.append(f"Falta hoja **{sheet}**")
                        else:
                            _sdf = pd.read_excel(uploaded, sheet_name=sheet, nrows=2)
                            uploaded.seek(0)
                            missing_cols = [c for c in cols if c not in _sdf.columns]
                            if missing_cols:
                                errors.append(f"Hoja {sheet}: faltan columnas **{', '.join(missing_cols)}**")
                    if errors:
                        st.error(f"❌ Template incompleto — {'; '.join(errors)}")
                        validation_ok = False
                    else:
                        st.success(f"Template correcto — {len(_xls.sheet_names)} hojas validadas")

            except Exception as e:
                st.warning(f"⚠️ No se pudo validar el archivo: {e}")

        c1, c2, c3 = st.columns([1, 1, 3])
        with c1:
            if st.button("Siguiente →", key=f"next_{key}", disabled=(uploaded is None or not validation_ok), type="primary"):
                guardar_temporal(uploaded, key)
                st.session_state.step = s + 1
                st.rerun()
        break

    # Ejecutar
    if current_step > 5:
        st.markdown("")

        # ── VALIDACIÓN PRE-CÁLCULO: cuentas no clasificadas ──
        pre_check_ok = True
        try:
            _gl_check = pd.read_csv(st.session_state.archivos_tmp["gl"], sep=';', skiprows=10, encoding='latin-1', usecols=[2])
            _gl_accts = set(pd.to_numeric(_gl_check.iloc[:,0], errors='coerce').dropna().astype(int).astype(str).unique())

            _pl_check = pd.read_excel(st.session_state.archivos_tmp["rubros"], sheet_name='PL_Totalizador', usecols=['Cuenta'])
            _pl_accts = set(_pl_check['Cuenta'].astype(str).str.strip().unique())

            try:
                _exc_check = pd.read_excel(st.session_state.archivos_tmp["rubros"], sheet_name='Cuentas_Excluidas', usecols=['Cuenta'])
                _exc_accts = set(_exc_check['Cuenta'].dropna().astype(int).astype(str).unique())
            except:
                _exc_accts = set()

            _sin_clasificar = _gl_accts - _pl_accts - _exc_accts
            if _sin_clasificar:
                pre_check_ok = False
                st.error(f"❌ **{len(_sin_clasificar)} cuenta(s) sin clasificar** — No están en PL ni en Cuentas Excluidas del template.")
                st.markdown("Agregá cada cuenta al **PL_Totalizador** (si afecta al cálculo) o a **Cuentas_Excluidas** (si no aplica):")
                for acct in sorted(_sin_clasificar):
                    st.markdown(f"- Cuenta **{acct}**")
                st.info("El cálculo no puede continuar hasta que todas las cuentas estén clasificadas.", icon="🚫")
            else:
                st.success(f"**Todos los archivos cargados** — {len(_gl_accts)} cuentas del GL clasificadas correctamente")
        except Exception as e:
            st.success("**Todos los archivos cargados** — Listo para calcular")

        c1, c2, c3 = st.columns([1, 1, 3])
        with c1:
            if st.button("▶️  Calcular", type="primary", use_container_width=True, disabled=(not pre_check_ok)):
                progress = st.progress(0, text="Inicializando...")
                try:
                    progress.progress(10, text="📄 Cargando Balance GL...")
                    import motor
                    motor.RUTA_RUBROS = st.session_state.archivos_tmp["rubros"]
                    motor.RUTA_BALANCE = st.session_state.archivos_tmp["gl"]
                    motor.RUTA_ORDENES = st.session_state.archivos_tmp["ordenes"]
                    motor.RUTA_TC = st.session_state.archivos_tmp["tc"]
                    motor.RUTA_ADI = st.session_state.archivos_tmp["adi"]

                    progress.progress(20, text="🏷️ Cargando Template Maestro...")
                    try:
                        _tp = pd.read_excel(st.session_state.archivos_tmp["rubros"], sheet_name='Productos')
                        _tp = _tp.dropna(subset=['COD', 'SEGMENTACION'])
                        motor.PROD_MAP = dict(zip(_tp['COD'].astype(int), _tp['SEGMENTACION'].astype(str)))
                    except: motor.PROD_MAP = {}
                    try:
                        _td = pd.read_excel(st.session_state.archivos_tmp["rubros"], sheet_name='DOM_INT')
                        _td = _td.dropna(subset=['Negocio2', 'DOM/INT'])
                        motor.DI_MAP = dict(zip(_td['Negocio2'].astype(int), _td['DOM/INT'].astype(str)))
                        motor.DI_INT = {str(k).zfill(4) for k, v in motor.DI_MAP.items() if v == 'INT'}
                    except: pass
                    try:
                        _tr = pd.read_excel(st.session_state.archivos_tmp["rubros"], sheet_name='RFC_Cost_Rubros')
                        motor.RFC_COST_RUBROS = set(_tr['PL Totalizador'].dropna().unique())
                    except: pass

                    progress.progress(30, text="📊 Procesando Balance...")
                    r = {}
                    mes_base = leer_periodo_balance(st.session_state.archivos_tmp["gl"]) or 10
                    r["mes_base"] = mes_base
                    df_glosario = cargar_glosario(st.session_state.archivos_tmp["rubros"])
                    df_balance = procesar_balance(st.session_state.archivos_tmp["gl"])

                    progress.progress(40, text="📦 Procesando Órdenes...")
                    _, df_si = procesar_ordenes(st.session_state.archivos_tmp["ordenes"])
                    df_adi = procesar_adi(st.session_state.archivos_tmp["adi"])
                    tc = cargar_tc(st.session_state.archivos_tmp["tc"])
                    r["tc"] = tc

                    progress.progress(50, text="🔑 Calculando Llaves K1-K12...")
                    llaves_ordenes = calcular_llaves_ordenes(df_si) if df_si is not None else {}
                    llaves_adi = calcular_llaves_adi(df_adi) if df_adi is not None else {}
                    r["llaves_ordenes"] = llaves_ordenes
                    r["llaves_adi"] = llaves_adi

                    tc_clp = tc.get('CLP', 946.38)
                    tc_pen = tc.get('PEN', 3.385)
                    tc_eur = tc.get('EUR', 0.8613)

                    progress.progress(60, text="🇪🇨 Calculando Ecuador...")
                    r["ecuador"] = calcular_ecuador_hoja_llave(df_balance, llaves_ordenes, mes_base=mes_base) if (df_balance is not None and llaves_ordenes) else None

                    progress.progress(68, text="🇨🇱 Calculando Chile...")
                    r["chile"] = calcular_chile_hoja_llave(df_balance, llaves_ordenes, tc_clp=tc_clp, mes_base=mes_base) if (df_balance is not None and llaves_ordenes) else None

                    progress.progress(76, text="🇵🇪 Calculando Perú...")
                    r["peru"] = calcular_peru_hoja_llave(df_balance, llaves_ordenes, tc_pen=tc_pen, mes_base=mes_base) if (df_balance is not None and llaves_ordenes) else None

                    progress.progress(84, text="🇺🇸 Calculando USA...")
                    r["usa"] = calcular_usa_hoja_llave(df_balance, llaves_ordenes, mes_base=mes_base) if (df_balance is not None and llaves_ordenes) else None

                    progress.progress(90, text="🇺🇾 Calculando TR...")
                    r["tr"] = calcular_tr_hoja_llave(df_balance, mes_base=mes_base) if df_balance is not None else None

                    progress.progress(96, text="🇪🇸 Calculando España...")
                    r["espana"] = calcular_espana_hoja_llave(df_balance, tc_eur=tc_eur, mes_base=mes_base) if df_balance is not None else None

                    progress.progress(100, text="✅ Cálculo completado")
                    st.session_state.resultados = r
                    st.rerun()
                except Exception as e:
                    st.error(f"Error: {e}")
                    import traceback
                    st.code(traceback.format_exc())
        with c2:
            if st.button("🔄 Reiniciar"):
                for p in st.session_state.archivos_tmp.values():
                    try:
                        os.unlink(p)
                    except:
                        pass
                st.session_state.step = 1
                st.session_state.archivos_tmp = {}
                st.rerun()

# ============================================================================
# PANTALLA DE RESULTADOS
# ============================================================================
if st.session_state.resultados is not None:
    mostrar_resultados(st.session_state.resultados)

    st.markdown("---")
    c1, c2, c3 = st.columns([1, 1, 3])
    with c1:
        if st.button("🔄 Nueva ejecución", type="primary", use_container_width=True):
            for p in st.session_state.archivos_tmp.values():
                try:
                    os.unlink(p)
                except:
                    pass
            st.session_state.step = 1
            st.session_state.archivos_tmp = {}
            st.session_state.resultados = None
            st.session_state.pop("bajada_data", None)
            st.rerun()

# Footer
st.markdown("""
<div class="app-footer">
    Facturación Intercompany v2.1 · Tax · Despegar
</div>
""", unsafe_allow_html=True)
