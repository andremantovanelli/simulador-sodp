import streamlit as st
import pandas as pd
import numpy as np
from datetime import date, timedelta
import io

st.set_page_config(page_title="Simulador ODP — Privalia", page_icon="📦", layout="wide")

# ═══════════════════════════════════════════════════════════════════════════════
# UTILITÁRIOS
# ═══════════════════════════════════════════════════════════════════════════════

def parse_br(s):
    if pd.api.types.is_numeric_dtype(s): return pd.to_numeric(s, errors='coerce')
    return pd.to_numeric(s.astype(str).str.replace('.','',regex=False).str.replace(',','.',regex=False), errors='coerce')

def proximos_dias_uteis(data_base, n_dias, feriados_set):
    """Retorna a data resultante de avançar/recuar n_dias úteis (sem fim de semana ou feriado)."""
    step = 1 if n_dias >= 0 else -1
    restante = abs(n_dias)
    d = data_base
    while restante > 0:
        d += timedelta(days=step)
        if d.weekday() < 5 and d not in feriados_set:
            restante -= 1
    return d

def calcular_datas_blocado(data_inicio, data_fim, feriados_set, dn_manual=None, do_manual=None):
    dn = dn_manual if dn_manual else proximos_dias_uteis(data_inicio, -2, feriados_set)
    do = do_manual if do_manual else proximos_dias_uteis(data_fim, 3, feriados_set)
    return dn, do

# ═══════════════════════════════════════════════════════════════════════════════
# CARREGAMENTO
# ═══════════════════════════════════════════════════════════════════════════════

@st.cache_data
def carregar_calendario(f):
    df = pd.read_excel(f).rename(columns={
        'Id externo de Campanha Sacarino':'id_campanha','Nome da campanha':'campanha',
        'Data de início':'data_inicio','Data de término':'data_fim',
        'Webdays':'webdays','Status':'status','Centro de distibuição':'cd',
        'Categoria':'categoria','Sector Calendar':'setor',
        'Previsão de venda peças':'previsao_pecas','Estoque Total':'estoque_total',
        'Modelo de negócio':'modelo_negocio','Gerência':'gerencia',
    })
    df['data_inicio'] = pd.to_datetime(df['data_inicio'], dayfirst=True, errors='coerce')
    df['data_fim']    = pd.to_datetime(df['data_fim'],    dayfirst=True, errors='coerce')
    cols = ['id_campanha','campanha','data_inicio','data_fim','webdays','status',
            'cd','categoria','setor','previsao_pecas','estoque_total','modelo_negocio','gerencia']
    df = df[[c for c in cols if c in df.columns]].drop_duplicates(subset=['id_campanha'])
    df = df[df['modelo_negocio'].str.upper().str.strip() == 'ODP'].dropna(subset=['data_inicio','data_fim'])
    df['webdays'] = pd.to_numeric(df['webdays'], errors='coerce').fillna(1).clip(lower=1)
    df['previsao_pecas'] = pd.to_numeric(df['previsao_pecas'], errors='coerce').fillna(0)
    df['estoque_total']  = pd.to_numeric(df['estoque_total'],  errors='coerce').fillna(0)
    return df

@st.cache_data
def carregar_vendas(f):
    df = pd.read_csv(f, sep=';', dtype=str)
    df.columns = [c.strip() for c in df.columns]
    df = df.rename(columns={'ID':'id_campanha','Dia_Click':'data_venda',
                             'Items':'pecas_vendidas','Orders':'pedidos','Revenue':'receita'})
    if 'id_campanha' in df.columns:
        df['id_campanha'] = pd.to_numeric(df['id_campanha'].str.replace('.','',regex=False), errors='coerce')
    if 'data_venda' in df.columns:
        df['data_venda'] = pd.to_datetime(df['data_venda'], dayfirst=True, errors='coerce')
    for c in ['pecas_vendidas','pedidos','receita']:
        if c in df.columns: df[c] = parse_br(df[c])
    return df

@st.cache_data
def carregar_odp_vertical(f):
    df = pd.read_excel(f)
    cols = ['ID campanha','Nome','Categoria','Peças','Peças Convertidas','Status','Data']
    df = df[[c for c in cols if c in df.columns]].rename(columns={
        'ID campanha':'id_campanha','Nome':'campanha','Categoria':'categoria',
        'Peças':'pecas','Peças Convertidas':'pecas_convertidas',
        'Status':'status_entrada','Data':'data_entrada'})
    df['data_entrada'] = pd.to_datetime(df['data_entrada'], errors='coerce')
    return df.dropna(subset=['data_entrada'])

# ═══════════════════════════════════════════════════════════════════════════════
# PARÂMETROS PADRÃO (da aba Suporte)
# ═══════════════════════════════════════════════════════════════════════════════

PARAMS_DEFAULT = {
    'conversoes': pd.DataFrame([
        {'categoria':'Fashion',             'fator_in':1.0,  'fator_out':1.0},
        {'categoria':'Shoes',               'fator_in':2.0,  'fator_out':2.5},
        {'categoria':'Kids',                'fator_in':0.741,'fator_out':1.0},
        {'categoria':'Baby',                'fator_in':0.741,'fator_out':1.0},
        {'categoria':'Sport',               'fator_in':1.0,  'fator_out':1.0},
        {'categoria':'Underwear',           'fator_in':0.667,'fator_out':1.0},
        {'categoria':'Beachwear',           'fator_in':0.667,'fator_out':1.0},
        {'categoria':'Home & Decor',        'fator_in':2.5,  'fator_out':2.632},
        {'categoria':'Beauty and Wellness', 'fator_in':2.5,  'fator_out':2.632},
        {'categoria':'Eyewear',             'fator_in':1.667,'fator_out':2.632},
        {'categoria':'Bags',                'fator_in':2.0,  'fator_out':2.632},
        {'categoria':'Clearance',           'fator_in':1.0,  'fator_out':1.0},
    ]),
    'pallets': pd.DataFrame([
        {'categoria':'Acessories',      'pecas_palete':921},
        {'categoria':'Baby',            'pecas_palete':451},
        {'categoria':'Babycare',        'pecas_palete':451},
        {'categoria':'Beachwear',       'pecas_palete':401},
        {'categoria':'Beauty',          'pecas_palete':260},
        {'categoria':'Bodywear',        'pecas_palete':498},
        {'categoria':'Clearance',       'pecas_palete':150},
        {'categoria':'Fashion',         'pecas_palete':400},
        {'categoria':'Fitness',         'pecas_palete':401},
        {'categoria':'Kids I',          'pecas_palete':421},
        {'categoria':'Kids II',         'pecas_palete':405},
        {'categoria':'Kids Brands',     'pecas_palete':405},
        {'categoria':'Kids Shoes',      'pecas_palete':135},
        {'categoria':'Kids Trends',     'pecas_palete':421},
        {'categoria':'Shoes I',         'pecas_palete':97},
        {'categoria':'Shoes II',        'pecas_palete':109},
        {'categoria':'Shoes III',       'pecas_palete':107},
        {'categoria':'Shoes Brands',    'pecas_palete':97},
        {'categoria':'Shoes Comfort',   'pecas_palete':108},
        {'categoria':'Shoes Trends',    'pecas_palete':107},
        {'categoria':'Sports',          'pecas_palete':396},
        {'categoria':'Sul Brands',      'pecas_palete':408},
        {'categoria':'Sul Trends',      'pecas_palete':396},
        {'categoria':'Underwear',       'pecas_palete':596},
        {'categoria':'Varejo Feminino', 'pecas_palete':336},
        {'categoria':'Varejo Masculino','pecas_palete':377},
    ]),
    'feriados': pd.DataFrame([
        {'data':date(2024,1,1),  'nome':'Ano Novo',             'local':'Nacional'},
        {'data':date(2024,2,12), 'nome':'Carnaval',             'local':'Nacional'},
        {'data':date(2024,2,13), 'nome':'Carnaval',             'local':'Nacional'},
        {'data':date(2024,3,29), 'nome':'Sexta-feira Santa',    'local':'Nacional'},
        {'data':date(2024,4,21), 'nome':'Tiradentes',           'local':'Nacional'},
        {'data':date(2024,5,1),  'nome':'Dia do Trabalho',      'local':'Nacional'},
        {'data':date(2024,5,30), 'nome':'Corpus Christi',       'local':'Nacional'},
        {'data':date(2024,7,9),  'nome':'Revolução Constit.',   'local':'SP'},
        {'data':date(2024,9,7),  'nome':'Independência',        'local':'Nacional'},
        {'data':date(2024,10,12),'nome':'Nossa Senhora',        'local':'Nacional'},
        {'data':date(2024,11,2), 'nome':'Finados',              'local':'Nacional'},
        {'data':date(2024,11,15),'nome':'Proclamação República','local':'Nacional'},
        {'data':date(2024,11,20),'nome':'Consciência Negra',    'local':'SP'},
        {'data':date(2024,11,29),'nome':'Black Friday',         'local':'Nacional'},
        {'data':date(2024,12,25),'nome':'Natal',                'local':'Nacional'},
        {'data':date(2025,1,1),  'nome':'Ano Novo',             'local':'Nacional'},
        {'data':date(2025,3,3),  'nome':'Carnaval',             'local':'Nacional'},
        {'data':date(2025,3,4),  'nome':'Carnaval',             'local':'Nacional'},
        {'data':date(2025,4,18), 'nome':'Sexta-feira Santa',    'local':'Nacional'},
        {'data':date(2025,4,21), 'nome':'Tiradentes',           'local':'Nacional'},
        {'data':date(2025,5,1),  'nome':'Dia do Trabalho',      'local':'Nacional'},
        {'data':date(2025,6,19), 'nome':'Corpus Christi',       'local':'Nacional'},
        {'data':date(2025,9,7),  'nome':'Independência',        'local':'Nacional'},
        {'data':date(2025,10,12),'nome':'Nossa Senhora',        'local':'Nacional'},
        {'data':date(2025,11,2), 'nome':'Finados',              'local':'Nacional'},
        {'data':date(2025,11,15),'nome':'Proclamação República','local':'Nacional'},
        {'data':date(2025,11,20),'nome':'Consciência Negra',    'local':'SP'},
        {'data':date(2025,12,25),'nome':'Natal',                'local':'Nacional'},
        {'data':date(2026,1,1),  'nome':'Ano Novo',             'local':'Nacional'},
        {'data':date(2026,2,16), 'nome':'Carnaval',             'local':'Nacional'},
        {'data':date(2026,2,17), 'nome':'Carnaval',             'local':'Nacional'},
        {'data':date(2026,4,3),  'nome':'Sexta-feira Santa',    'local':'Nacional'},
        {'data':date(2026,4,21), 'nome':'Tiradentes',           'local':'Nacional'},
        {'data':date(2026,5,1),  'nome':'Dia do Trabalho',      'local':'Nacional'},
        {'data':date(2026,6,4),  'nome':'Corpus Christi',       'local':'Nacional'},
        {'data':date(2026,9,7),  'nome':'Independência',        'local':'Nacional'},
        {'data':date(2026,10,12),'nome':'Nossa Senhora',        'local':'Nacional'},
        {'data':date(2026,11,2), 'nome':'Finados',              'local':'Nacional'},
        {'data':date(2026,11,15),'nome':'Proclamação República','local':'Nacional'},
        {'data':date(2026,11,20),'nome':'Consciência Negra',    'local':'SP'},
        {'data':date(2026,12,25),'nome':'Natal',                'local':'Nacional'},
    ]),
    'lead_times': {'wlt': 1, 'dlt': 3},
    'curva_venda': pd.DataFrame([
        {'categoria':'Fashion', 'webdays':7,  'd1':0.156,'d2':0.290,'d3':0.155,'d4':0.118,'d5':0.096,'d6':0.087,'d7':0.098},
        {'categoria':'Fashion', 'webdays':8,  'd1':0.075,'d2':0.252,'d3':0.173,'d4':0.140,'d5':0.111,'d6':0.095,'d7':0.086,'d8':0.090},
        {'categoria':'Fashion', 'webdays':9,  'd1':0.042,'d2':0.225,'d3':0.121,'d4':0.134,'d5':0.086,'d6':0.123,'d7':0.104,'d8':0.083,'d9':0.081},
        {'categoria':'Shoes',   'webdays':7,  'd1':0.104,'d2':0.361,'d3':0.165,'d4':0.124,'d5':0.097,'d6':0.077,'d7':0.072},
        {'categoria':'Shoes',   'webdays':8,  'd1':0.043,'d2':0.272,'d3':0.161,'d4':0.134,'d5':0.120,'d6':0.090,'d7':0.087,'d8':0.093},
        {'categoria':'Shoes',   'webdays':9,  'd1':0.045,'d2':0.259,'d3':0.150,'d4':0.122,'d5':0.099,'d6':0.088,'d7':0.074,'d8':0.080,'d9':0.083},
        {'categoria':'Kids',    'webdays':7,  'd1':0.108,'d2':0.344,'d3':0.189,'d4':0.143,'d5':0.086,'d6':0.062,'d7':0.068},
        {'categoria':'Kids',    'webdays':8,  'd1':0.025,'d2':0.228,'d3':0.163,'d4':0.131,'d5':0.114,'d6':0.121,'d7':0.115,'d8':0.103},
        {'categoria':'Kids',    'webdays':9,  'd1':0.057,'d2':0.270,'d3':0.181,'d4':0.103,'d5':0.088,'d6':0.083,'d7':0.076,'d8':0.074,'d9':0.068},
        {'categoria':'Sports',  'webdays':7,  'd1':0.069,'d2':0.372,'d3':0.152,'d4':0.099,'d5':0.108,'d6':0.097,'d7':0.100},
        {'categoria':'Sports',  'webdays':8,  'd1':0.112,'d2':0.329,'d3':0.144,'d4':0.099,'d5':0.090,'d6':0.082,'d7':0.071,'d8':0.072},
        {'categoria':'Sports',  'webdays':9,  'd1':0.088,'d2':0.283,'d3':0.134,'d4':0.105,'d5':0.095,'d6':0.070,'d7':0.081,'d8':0.072,'d9':0.071},
    ]),
}

def get_params():
    if 'params' not in st.session_state:
        st.session_state['params'] = {
            'conversoes': PARAMS_DEFAULT['conversoes'].copy(),
            'pallets':    PARAMS_DEFAULT['pallets'].copy(),
            'feriados':   PARAMS_DEFAULT['feriados'].copy(),
            'lead_times': PARAMS_DEFAULT['lead_times'].copy(),
            'curva_venda':PARAMS_DEFAULT['curva_venda'].copy(),
        }
    return st.session_state['params']

def get_feriados_set(cd='Extrema'):
    params = get_params()
    df = params['feriados']
    locais_validos = ['Nacional']
    if 'extrema' in cd.lower():  locais_validos.append('Extrema'); locais_validos.append('SP')
    if 'jandira' in cd.lower():  locais_validos.append('Jandira'); locais_validos.append('SP')
    return set(df[df['local'].isin(locais_validos)]['data'].tolist())

def get_fator(categoria, tipo, params):
    """Busca fator de conversão IN ou OUT para a categoria."""
    df = params['conversoes']
    row = df[df['categoria'].str.lower() == str(categoria).lower()]
    if row.empty:
        # tenta match parcial
        for _, r in df.iterrows():
            if r['categoria'].lower() in str(categoria).lower():
                return r[f'fator_{tipo}']
        return 1.0
    return float(row.iloc[0][f'fator_{tipo}'])

def get_pecas_palete(categoria, params):
    df = params['pallets']
    row = df[df['categoria'].str.lower() == str(categoria).lower()]
    if row.empty:
        for _, r in df.iterrows():
            if r['categoria'].lower() in str(categoria).lower():
                return float(r['pecas_palete'])
        return 400.0
    return float(row.iloc[0]['pecas_palete'])

def get_curva(categoria, webdays, params):
    """Retorna lista de pesos por dia relativo."""
    df = params['curva_venda']
    wd = int(webdays)
    row = df[(df['categoria'].str.lower() == str(categoria).lower()) & (df['webdays'] == wd)]
    if row.empty:
        # fallback: distribuição uniforme
        return [1.0/wd]*wd
    cols_d = [f'd{i}' for i in range(1, wd+1) if f'd{i}' in row.columns]
    vals = [float(row.iloc[0][c]) for c in cols_d]
    total = sum(vals)
    return [v/total for v in vals] if total > 0 else [1.0/wd]*wd

# ═══════════════════════════════════════════════════════════════════════════════
# MOTOR DE CÁLCULO
# ═══════════════════════════════════════════════════════════════════════════════

def calcular_saidas(calendario, vendas, data_hoje, params, ajustes=None):
    """
    Para cada campanha ODP, gera volume de saída (picking) por dia.
    Usa curva de venda histórica para distribuir a previsão.
    """
    if ajustes is None: ajustes = {}
    linhas = []
    for _, camp in calendario.iterrows():
        cid = camp['id_campanha']
        aj  = ajustes.get(cid, {})

        inicio   = aj.get('data_inicio', camp['data_inicio'].date())
        fim      = aj.get('data_fim',    camp['data_fim'].date())
        webdays  = max(int((fim - inicio).days + 1), 1)
        prev_tot = aj.get('previsao_pecas', camp['previsao_pecas'])
        cat      = camp['categoria']

        curva = get_curva(cat, webdays, params)
        fator_out = get_fator(cat, 'out', params)

        vendas_camp = pd.DataFrame()
        if vendas is not None and 'id_campanha' in vendas.columns:
            vendas_camp = vendas[vendas['id_campanha'] == cid].copy()

        datas = [inicio + timedelta(days=i) for i in range(webdays)]
        for idx, d in enumerate(datas):
            peso = curva[idx] if idx < len(curva) else (1.0/webdays)
            prev_dia = prev_tot * peso

            if d < data_hoje:
                if not vendas_camp.empty and 'data_venda' in vendas_camp.columns:
                    vd = vendas_camp[vendas_camp['data_venda'].dt.date == d]
                    pecas_bruto = float(vd['pecas_vendidas'].sum()) if not vd.empty else prev_dia
                else:
                    pecas_bruto = prev_dia
                tipo = 'realizado'
            elif d == data_hoje:
                if not vendas_camp.empty and 'data_venda' in vendas_camp.columns:
                    vd = vendas_camp[vendas_camp['data_venda'].dt.date == d]
                    pecas_bruto = float(vd['pecas_vendidas'].sum()) if not vd.empty else prev_dia
                else:
                    pecas_bruto = prev_dia
                tipo = 'reforecast'
            else:
                pecas_bruto = prev_dia
                tipo = 'previsao'

            linhas.append({
                'id_campanha': cid,
                'campanha':    camp['campanha'],
                'data':        d,
                'gerencia':    camp.get('gerencia',''),
                'categoria':   cat,
                'cd':          camp.get('cd',''),
                'pecas_bruto': round(pecas_bruto, 1),
                'pecas_conv':  round(pecas_bruto * fator_out, 1),
                'tipo':        tipo,
            })
    return pd.DataFrame(linhas)

def calcular_blocado(calendario, data_hoje, params, ajustes=None):
    """
    Calcula o saldo do estoque blocado (horizontal/picking) por dia.
    Entra na DN, permanece integral, sai na DO.
    """
    if ajustes is None: ajustes = {}
    linhas = []
    for _, camp in calendario.iterrows():
        cid   = camp['id_campanha']
        aj    = ajustes.get(cid, {})
        cd    = camp.get('cd', 'Extrema')
        cat   = camp['categoria']
        feriados = get_feriados_set(cd)

        inicio = aj.get('data_inicio', camp['data_inicio'].date())
        fim    = aj.get('data_fim',    camp['data_fim'].date())
        est    = aj.get('estoque_total', camp['estoque_total'])

        dn = aj.get('dn', None)
        do = aj.get('do', None)
        dn, do = calcular_datas_blocado(inicio, fim, feriados, dn, do)

        fator_in  = get_fator(cat, 'in', params)
        ppp       = get_pecas_palete(cat, params)
        est_conv  = est * fator_in
        pallets   = est_conv / ppp if ppp > 0 else 0

        linhas.append({
            'id_campanha': cid,
            'campanha':    camp['campanha'],
            'categoria':   cat,
            'cd':          cd,
            'gerencia':    camp.get('gerencia',''),
            'data_inicio': inicio,
            'data_fim':    fim,
            'dn':          dn,
            'do':          do,
            'estoque_pecas': round(est, 0),
            'estoque_conv':  round(est_conv, 0),
            'pallets':       round(pallets, 2),
            'status':        'blocado' if dn <= data_hoje <= do else (
                             'aguardando' if data_hoje < dn else 'liberado'),
        })
    df = pd.DataFrame(linhas)
    if df.empty: return df, pd.DataFrame()

    # Série temporal: saldo dia a dia
    if df.empty: return df, pd.DataFrame()
    min_d = df['dn'].min()
    max_d = df['do'].max()
    datas = pd.date_range(min_d, max_d).date
    serie = []
    for d in datas:
        ativas = df[(df['dn'] <= d) & (df['do'] >= d)]
        serie.append({
            'data':       d,
            'pecas':      ativas['estoque_pecas'].sum(),
            'pecas_conv': ativas['estoque_conv'].sum(),
            'pallets':    ativas['pallets'].sum(),
            'campanhas':  len(ativas),
        })
    return df, pd.DataFrame(serie)

# ═══════════════════════════════════════════════════════════════════════════════
# INTERFACE
# ═══════════════════════════════════════════════════════════════════════════════

st.title("Simulador ODP — Privalia")

# Session state para ajustes de simulação
if 'ajustes' not in st.session_state: st.session_state['ajustes'] = {}

# ── Abas ──────────────────────────────────────────────────────────────────────
aba_dados, aba_saidas, aba_blocado, aba_reforecast, aba_sim, aba_params, aba_export = st.tabs([
    "Dados", "Saídas", "Blocado", "Reforecast", "Simulações", "Parâmetros", "Exportar"
])

# ═══════════════════════════════════════════════════════════════════════════════
# ABA DADOS
# ═══════════════════════════════════════════════════════════════════════════════
with aba_dados:
    st.subheader("Carregar dados")
    col1, col2, col3 = st.columns(3)
    with col1:
        arq_cal  = st.file_uploader("Calendário Salesforce (.xlsx)", type=["xlsx","xls"])
    with col2:
        arq_vend = st.file_uploader("Base de vendas DBeaver (.csv)", type=["csv"])
    with col3:
        arq_odp  = st.file_uploader("ODP Vertical (.xlsx)", type=["xlsx","xls"])

    st.divider()
    col_a, col_b = st.columns(2)
    with col_a:
        data_hoje = st.date_input("Data de referência (hoje)", value=date.today())
    with col_b:
        cds_opcoes = ["Extrema", "Jandira", "Extrema e Jandira"]
        sel_cd_global = st.selectbox("Centro de distribuição", cds_opcoes)

    if arq_cal:
        cal = carregar_calendario(arq_cal)
        st.success(f"Calendário: {len(cal)} campanhas ODP carregadas")
        with st.expander("Prévia"):
            st.dataframe(cal.head(10), use_container_width=True, hide_index=True)

    if arq_vend:
        try:
            vendas = carregar_vendas(arq_vend)
            st.success(f"Vendas: {len(vendas):,} registros")
        except Exception as e:
            st.error(f"Erro vendas: {e}")
            vendas = None

    if arq_odp:
        try:
            odp_vert = carregar_odp_vertical(arq_odp)
            st.success(f"ODP Vertical: {len(odp_vert):,} registros")
        except Exception as e:
            st.error(f"Erro ODP: {e}")

# Guarda no session state para usar em outras abas
if arq_cal is not None:
    st.session_state['cal']        = carregar_calendario(arq_cal)
    st.session_state['data_hoje']  = data_hoje
    st.session_state['sel_cd']     = sel_cd_global

cal_ok   = 'cal' in st.session_state
data_ref = st.session_state.get('data_hoje', date.today())
ajustes  = st.session_state.get('ajustes', {})
params   = get_params()

def filtrar_cd(df, col='cd'):
    sel = st.session_state.get('sel_cd','Extrema e Jandira')
    if sel == 'Extrema e Jandira': return df
    return df[df[col].str.lower().str.contains(sel.lower(), na=False)] if col in df.columns else df

# ═══════════════════════════════════════════════════════════════════════════════
# ABA SAÍDAS
# ═══════════════════════════════════════════════════════════════════════════════
with aba_saidas:
    if not cal_ok:
        st.info("Carregue o calendário na aba Dados.")
    else:
        cal = st.session_state['cal']
        vendas = None
        if arq_vend is not None:
            try: vendas = carregar_vendas(arq_vend)
            except: pass

        # Filtros
        col_f1, col_f2, col_f3, col_f4 = st.columns(4)
        with col_f1:
            conv_saida = st.radio("Unidade", ["Não convertido","Convertido"], horizontal=True,
                                  help="Convertido aplica fator OUT por categoria")
        with col_f2:
            gers  = sorted(cal['gerencia'].dropna().unique().tolist())
            sel_ger = st.multiselect("Gerência", gers, default=gers, key='sai_ger')
        with col_f3:
            cats  = sorted(cal['categoria'].dropna().unique().tolist())
            sel_cat = st.multiselect("Categoria", cats, default=cats, key='sai_cat')
        with col_f4:
            tipos_vis = st.multiselect("Tipo", ["realizado","reforecast","previsao"],
                                       default=["realizado","reforecast","previsao"], key='sai_tipo')

        cal_f = filtrar_cd(cal)
        cal_f = cal_f[cal_f['gerencia'].isin(sel_ger) & cal_f['categoria'].isin(sel_cat)]

        with st.spinner("Calculando saídas..."):
            df_sai = calcular_saidas(cal_f, vendas, data_ref, params, ajustes)

        if df_sai.empty:
            st.warning("Nenhum dado para os filtros selecionados.")
        else:
            col_pecas = 'pecas_conv' if conv_saida == 'Convertido' else 'pecas_bruto'
            df_vis = df_sai[df_sai['tipo'].isin(tipos_vis)]

            # Métricas
            m1,m2,m3,m4 = st.columns(4)
            total = df_vis[col_pecas].sum()
            real  = df_vis[df_vis['tipo']=='realizado'][col_pecas].sum()
            prev  = df_vis[df_vis['tipo'].isin(['reforecast','previsao'])][col_pecas].sum()
            m1.metric("Total peças", f"{total:,.0f}".replace(',','.'))
            m2.metric("Realizado",   f"{real:,.0f}".replace(',','.'))
            m3.metric("A realizar",  f"{prev:,.0f}".replace(',','.'))
            m4.metric("Campanhas",   str(df_vis['id_campanha'].nunique()))

            st.subheader("Volume de picking por dia")
            pivot = (df_vis.groupby(['data','tipo'])[col_pecas].sum().reset_index()
                     .pivot_table(index='data',columns='tipo',values=col_pecas,aggfunc='sum')
                     .fillna(0).sort_index())
            pivot.index = pd.to_datetime(pivot.index)
            pivot.columns.name = None
            pivot = pivot.rename(columns={'realizado':'Realizado','reforecast':'Reforecast','previsao':'Previsão'})
            st.bar_chart(pivot, use_container_width=True)

            st.subheader("Por gerência")
            por_ger = (df_vis.groupby(['gerencia','tipo'])[col_pecas].sum().reset_index()
                       .pivot_table(index='gerencia',columns='tipo',values=col_pecas,aggfunc='sum')
                       .fillna(0).reset_index())
            por_ger.columns.name = None
            for c in ['realizado','reforecast','previsao']:
                if c in por_ger.columns: por_ger[c] = por_ger[c].round(0).astype(int)
            por_ger = por_ger.rename(columns={'gerencia':'Gerência','realizado':'Realizado',
                                               'reforecast':'Reforecast','previsao':'Previsão'})
            st.dataframe(por_ger, use_container_width=True, hide_index=True)

            st.subheader("Por campanha")
            por_camp = (df_vis.groupby(['id_campanha','campanha','categoria','cd'])[col_pecas].sum()
                        .reset_index().sort_values(col_pecas, ascending=False))
            por_camp[col_pecas] = por_camp[col_pecas].round(0).astype(int)
            por_camp = por_camp.rename(columns={'id_campanha':'ID','campanha':'Campanha',
                                                'categoria':'Categoria','cd':'CD',
                                                col_pecas:'Peças'})
            st.dataframe(por_camp, use_container_width=True, hide_index=True)

            with st.expander("Detalhe diário por campanha"):
                det = df_vis[['data','campanha','tipo',col_pecas]].copy()
                det['data'] = det['data'].astype(str)
                det[col_pecas] = det[col_pecas].round(1)
                det = det.rename(columns={'data':'Data','campanha':'Campanha',
                                          'tipo':'Tipo',col_pecas:'Peças'})
                st.dataframe(det, use_container_width=True, hide_index=True)

# ═══════════════════════════════════════════════════════════════════════════════
# ABA BLOCADO
# ═══════════════════════════════════════════════════════════════════════════════
with aba_blocado:
    if not cal_ok:
        st.info("Carregue o calendário na aba Dados.")
    else:
        cal = st.session_state['cal']
        cal_f = filtrar_cd(cal)

        col_f1, col_f2 = st.columns(2)
        with col_f1:
            conv_bloc = st.radio("Unidade", ["Peças","Peças convertidas","Pallets"],
                                 horizontal=True, key='bloc_conv')
        with col_f2:
            gers_b = sorted(cal_f['gerencia'].dropna().unique().tolist())
            sel_ger_b = st.multiselect("Gerência", gers_b, default=gers_b, key='bloc_ger')

        cal_fb = cal_f[cal_f['gerencia'].isin(sel_ger_b)]

        with st.spinner("Calculando blocado..."):
            df_bloc, serie_bloc = calcular_blocado(cal_fb, data_ref, params, ajustes)

        if df_bloc.empty:
            st.warning("Nenhum dado de blocado.")
        else:
            col_serie = {'Peças':'pecas','Peças convertidas':'pecas_conv','Pallets':'pallets'}[conv_bloc]
            unidade   = conv_bloc

            # Métricas
            m1,m2,m3 = st.columns(3)
            hoje_bloc = df_bloc[df_bloc['status']=='blocado']
            m1.metric("Campanhas blocadas hoje", str(len(hoje_bloc)))
            m2.metric(f"{unidade} blocadas hoje",
                      f"{hoje_bloc[col_serie.replace('pecas','estoque_pecas').replace('pecas_conv','estoque_conv').replace('pallets','pallets')].sum():,.0f}".replace(',','.'))
            m3.metric("Campanhas aguardando", str(len(df_bloc[df_bloc['status']=='aguardando'])))

            st.subheader(f"Saldo blocado por dia — {unidade}")
            if not serie_bloc.empty:
                serie_vis = serie_bloc[['data', col_serie]].set_index('data').rename(columns={col_serie: unidade})
                serie_vis.index = pd.to_datetime(serie_vis.index)
                st.line_chart(serie_vis, use_container_width=True)

            st.subheader("Campanhas no picking")
            exib = df_bloc[['id_campanha','campanha','categoria','cd','gerencia',
                             'data_inicio','data_fim','dn','do',
                             'estoque_pecas','estoque_conv','pallets','status']].copy()
            for c in ['data_inicio','data_fim','dn','do']:
                exib[c] = exib[c].astype(str)
            exib['estoque_pecas'] = exib['estoque_pecas'].astype(int)
            exib['estoque_conv']  = exib['estoque_conv'].astype(int)
            exib['pallets']       = exib['pallets'].round(2)
            exib = exib.rename(columns={
                'id_campanha':'ID','campanha':'Campanha','categoria':'Categoria',
                'cd':'CD','gerencia':'Gerência','data_inicio':'Início','data_fim':'Fim',
                'dn':'Data descida','do':'Data subida',
                'estoque_pecas':'Peças','estoque_conv':'Peças conv.','pallets':'Pallets','status':'Status'})
            st.dataframe(exib, use_container_width=True, hide_index=True)

# ═══════════════════════════════════════════════════════════════════════════════
# ABA REFORECAST
# ═══════════════════════════════════════════════════════════════════════════════
with aba_reforecast:
    if not cal_ok:
        st.info("Carregue o calendário na aba Dados.")
    else:
        cal = st.session_state['cal']
        vendas = None
        if arq_vend is not None:
            try: vendas = carregar_vendas(arq_vend)
            except: pass

        cal_f = filtrar_cd(cal)
        df_sai_rf = calcular_saidas(cal_f, vendas, data_ref, params, ajustes)

        if df_sai_rf.empty:
            st.warning("Sem dados.")
        else:
            st.subheader("Previsão original vs reforecast vs realizado")
            resumo_rf = (df_sai_rf.groupby(['id_campanha','campanha','categoria'])
                         .agg(
                             previsao=('pecas_bruto', lambda x: x[df_sai_rf.loc[x.index,'tipo'].isin(['previsao','reforecast'])].sum() + x[df_sai_rf.loc[x.index,'tipo']=='realizado'].sum()),
                             realizado=('pecas_bruto', lambda x: x[df_sai_rf.loc[x.index,'tipo']=='realizado'].sum()),
                             reforecast=('pecas_bruto', lambda x: x[df_sai_rf.loc[x.index,'tipo'].isin(['reforecast','previsao'])].sum()),
                         ).reset_index())

            # Junta previsão original do calendário
            prev_orig = cal_f[['id_campanha','previsao_pecas']].copy()
            resumo_rf = resumo_rf.merge(prev_orig, on='id_campanha', how='left')
            resumo_rf['desvio_pct'] = ((resumo_rf['realizado'] - resumo_rf['previsao_pecas'])
                                        / resumo_rf['previsao_pecas'].replace(0, np.nan) * 100).round(1)
            for c in ['previsao_pecas','realizado','reforecast']:
                resumo_rf[c] = resumo_rf[c].round(0).astype(int)
            resumo_rf = resumo_rf.rename(columns={
                'id_campanha':'ID','campanha':'Campanha','categoria':'Categoria',
                'previsao_pecas':'Previsão original','realizado':'Realizado',
                'reforecast':'A realizar','desvio_pct':'Desvio % (real vs prev)'})
            st.dataframe(resumo_rf, use_container_width=True, hide_index=True)

            st.subheader("Evolução acumulada")
            acum = (df_sai_rf.sort_values('data')
                    .assign(data=lambda x: pd.to_datetime(x['data']))
                    .groupby(['data','tipo'])['pecas_bruto'].sum().reset_index()
                    .pivot_table(index='data',columns='tipo',values='pecas_bruto',aggfunc='sum')
                    .fillna(0).cumsum())
            acum.columns.name = None
            acum = acum.rename(columns={'realizado':'Realizado','reforecast':'Reforecast','previsao':'Previsão'})
            st.line_chart(acum, use_container_width=True)

# ═══════════════════════════════════════════════════════════════════════════════
# ABA SIMULAÇÕES
# ═══════════════════════════════════════════════════════════════════════════════
with aba_sim:
    if not cal_ok:
        st.info("Carregue o calendário na aba Dados.")
    else:
        cal = st.session_state['cal']
        st.caption("Ajustes aqui são efêmeros — não afetam as bases originais. Use 'Resetar' para voltar ao estado original.")

        if st.button("Resetar todas as simulações"):
            st.session_state['ajustes'] = {}
            st.rerun()

        camp_ids  = cal['id_campanha'].tolist()
        camp_nomes = cal.set_index('id_campanha')['campanha'].to_dict()
        opcoes = [f"{cid} — {camp_nomes.get(cid,'')}" for cid in camp_ids]
        sel_str = st.selectbox("Selecionar campanha para ajustar", opcoes)
        cid_sel = int(sel_str.split(' — ')[0])
        camp_row = cal[cal['id_campanha'] == cid_sel].iloc[0]
        aj_atual = st.session_state['ajustes'].get(cid_sel, {})

        st.markdown(f"**{camp_row['campanha']}** — {camp_row['categoria']} | CD: {camp_row.get('cd','')}")

        col1, col2 = st.columns(2)
        with col1:
            new_ini = st.date_input("Data início",
                value=aj_atual.get('data_inicio', camp_row['data_inicio'].date()), key=f'ini_{cid_sel}')
            new_fim = st.date_input("Data fim",
                value=aj_atual.get('data_fim', camp_row['data_fim'].date()), key=f'fim_{cid_sel}')
            new_est = st.number_input("Estoque total (peças)",
                value=float(aj_atual.get('estoque_total', camp_row['estoque_total'])),
                min_value=0.0, step=100.0, key=f'est_{cid_sel}')
        with col2:
            feriados_sim = get_feriados_set(camp_row.get('cd','Extrema'))
            dn_calc, do_calc = calcular_datas_blocado(new_ini, new_fim, feriados_sim)
            new_dn = st.date_input("Data descida (blocado)",
                value=aj_atual.get('dn', dn_calc), key=f'dn_{cid_sel}')
            new_do = st.date_input("Data subida (blocado)",
                value=aj_atual.get('do', do_calc), key=f'do_{cid_sel}')
            new_prev = st.number_input("Previsão de vendas (peças)",
                value=float(aj_atual.get('previsao_pecas', camp_row['previsao_pecas'])),
                min_value=0.0, step=100.0, key=f'prev_{cid_sel}')

        if st.button("Aplicar ajuste"):
            st.session_state['ajustes'][cid_sel] = {
                'data_inicio':   new_ini,
                'data_fim':      new_fim,
                'estoque_total': new_est,
                'dn':            new_dn,
                'do':            new_do,
                'previsao_pecas':new_prev,
            }
            st.success(f"Ajuste aplicado para {camp_row['campanha']}. Verifique as abas Saídas e Blocado.")

        if st.session_state['ajustes']:
            st.subheader("Ajustes ativos")
            rows = []
            for cid, aj in st.session_state['ajustes'].items():
                rows.append({'ID': cid, 'Campanha': camp_nomes.get(cid,''),
                             'Início': str(aj.get('data_inicio','')),
                             'Fim': str(aj.get('data_fim','')),
                             'Estoque': aj.get('estoque_total',''),
                             'DN': str(aj.get('dn','')),
                             'DO': str(aj.get('do',''))})
            st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

# ═══════════════════════════════════════════════════════════════════════════════
# ABA PARÂMETROS
# ═══════════════════════════════════════════════════════════════════════════════
with aba_params:
    senha = st.text_input("Senha de acesso", type="password", key="senha_params")
    if senha == "privalia2024":
        st.success("Acesso liberado.")
        params = get_params()

        tab_p1, tab_p2, tab_p3, tab_p4 = st.tabs(["Conversões","Pallets","Lead times","Curva de venda"])

        with tab_p1:
            st.caption("Fatores IN e OUT por categoria")
            df_conv = st.data_editor(params['conversoes'], use_container_width=True,
                                     hide_index=True, num_rows="dynamic")
            params['conversoes'] = df_conv

        with tab_p2:
            st.caption("Peças por palete por categoria")
            df_pal = st.data_editor(params['pallets'], use_container_width=True,
                                    hide_index=True, num_rows="dynamic")
            params['pallets'] = df_pal

        with tab_p3:
            st.caption("Lead times globais (dias úteis)")
            wlt = st.number_input("WLT (Warehouse Lead Time)", value=params['lead_times']['wlt'],
                                   min_value=0, max_value=10, step=1)
            dlt = st.number_input("DLT (Delivery Lead Time)",  value=params['lead_times']['dlt'],
                                   min_value=0, max_value=10, step=1)
            params['lead_times'] = {'wlt': wlt, 'dlt': dlt}

            st.caption("Feriados")
            df_fer = st.data_editor(params['feriados'], use_container_width=True,
                                    hide_index=True, num_rows="dynamic")
            params['feriados'] = df_fer

        with tab_p4:
            st.caption("Curva de venda histórica por categoria e webdays")
            df_curva = st.data_editor(params['curva_venda'], use_container_width=True,
                                      hide_index=True, num_rows="dynamic")
            params['curva_venda'] = df_curva

        st.session_state['params'] = params

    elif senha:
        st.error("Senha incorreta.")

# ═══════════════════════════════════════════════════════════════════════════════
# ABA EXPORTAR
# ═══════════════════════════════════════════════════════════════════════════════
with aba_export:
    st.subheader("Exportar resultados")
    st.caption("Exporta o estado atual da simulação. Nada é salvo no sistema.")

    if not cal_ok:
        st.info("Carregue os dados primeiro.")
    else:
        cal = st.session_state['cal']
        vendas = None
        if arq_vend is not None:
            try: vendas = carregar_vendas(arq_vend)
            except: pass

        cal_f = filtrar_cd(cal)
        df_sai_exp = calcular_saidas(cal_f, vendas, data_ref, params, ajustes)
        df_bloc_exp, serie_bloc_exp = calcular_blocado(cal_f, data_ref, params, ajustes)

        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine='openpyxl') as writer:
            if not df_sai_exp.empty:
                pivot_exp = (df_sai_exp.groupby(['data','tipo'])['pecas_bruto'].sum().reset_index()
                             .pivot_table(index='data',columns='tipo',values='pecas_bruto',aggfunc='sum')
                             .fillna(0).sort_index())
                pivot_exp.columns.name = None
                pivot_exp.to_excel(writer, sheet_name='Saídas — dia a dia')
                df_sai_exp.to_excel(writer, sheet_name='Saídas — detalhe', index=False)
            if not df_bloc_exp.empty:
                df_bloc_exp.to_excel(writer, sheet_name='Blocado — campanhas', index=False)
            if not serie_bloc_exp.empty:
                serie_bloc_exp.to_excel(writer, sheet_name='Blocado — série', index=False)
            cal_f.to_excel(writer, sheet_name='Calendário ODP', index=False)

        st.download_button(
            "Baixar Excel",
            data=buf.getvalue(),
            file_name=f"simulador_odp_{date.today().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
