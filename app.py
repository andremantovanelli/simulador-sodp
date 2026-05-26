import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
from datetime import date, timedelta
import io, sys, os
sys.path.insert(0, '/home/claude')
from odp_parser import (carregar_odp, gerar_serie_blocado,
                         gerar_serie_saidas, carregar_curvas,
                         get_curva, FERIADOS)
from openpyxl import load_workbook

st.set_page_config(
    page_title="ODP · Privalia",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ── Tema ──────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600&family=DM+Mono:wght@400;500&display=swap');

html, body, [class*="css"] { font-family: 'DM Sans', sans-serif; }
.block-container { padding: 1.5rem 2rem 2rem; }

/* tabs */
[data-baseweb="tab-list"] { gap: 0; border-bottom: 1px solid #e5e7eb; }
[data-baseweb="tab"] { 
    font-family: 'DM Sans', sans-serif; font-size: 13px; font-weight: 500;
    padding: 10px 18px; color: #6b7280;
}
[aria-selected="true"] { color: #111827 !important; border-bottom: 2px solid #111827 !important; }

/* métricas */
[data-testid="metric-container"] { 
    background: #f9fafb; border-radius: 10px; padding: 16px 20px;
    border: 1px solid #f3f4f6;
}
[data-testid="stMetricValue"] { font-size: 26px !important; font-weight: 600; }
[data-testid="stMetricLabel"] { font-size: 12px !important; color: #6b7280; }

/* tabela */
[data-testid="stDataFrame"] { font-size: 12px; }

/* expander */
[data-testid="stExpander"] { border: 1px solid #f3f4f6 !important; border-radius: 10px; }

/* badges */
.badge-on  { background:#dcfce7; color:#166534; padding:2px 8px; border-radius:4px; font-size:11px; font-weight:600; }
.badge-x   { background:#f3f4f6; color:#6b7280; padding:2px 8px; border-radius:4px; font-size:11px; }
.badge-fut { background:#eff6ff; color:#1d4ed8; padding:2px 8px; border-radius:4px; font-size:11px; }
.alerta    { background:#fef2f2; color:#991b1b; padding:2px 8px; border-radius:4px; font-size:11px; font-weight:600; }

.titulo-sec { font-size: 11px; font-weight: 600; color: #9ca3af; 
              text-transform: uppercase; letter-spacing: .06em; margin: 20px 0 10px; }
</style>
""", unsafe_allow_html=True)

# ── Helpers ───────────────────────────────────────────────────────────────────
def fmt_n(v, dec=0):
    if v is None or (isinstance(v, float) and np.isnan(v)): return "—"
    return f"{v:,.{dec}f}".replace(",", "X").replace(".", ",").replace("X", ".")

def fmt_pct(v, dec=1):
    if v is None or (isinstance(v, float) and np.isnan(v)): return "—"
    return f"{v*100:.{dec}f}%"

def cor_desvio(v):
    if abs(v) < 0.05: return "#16a34a"
    if abs(v) < 0.15: return "#d97706"
    return "#dc2626"

CAP_BU = 4200        # pallets capacidade armazenagem BU
CAP_CHECKIN = 50000  # peças/dia check-in Extrema
CAP_CHECKOUT = 45000 # peças/dia check-out Extrema

CORES_TIPO = {
    'realizado': '#1e3a5f',
    'reforecast': '#3b82f6',
    'previsto':   '#bfdbfe',
}

# ── Carregamento de dados ─────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def carregar(arquivo_bytes, nome_arquivo):
    import tempfile
    with tempfile.NamedTemporaryFile(suffix='.xlsm', delete=False) as tmp:
        tmp.write(arquivo_bytes)
        tmp_path = tmp.name
    df = carregar_odp(tmp_path)
    os.unlink(tmp_path)
    return df

# ── Estado da sessão ──────────────────────────────────────────────────────────
if 'df' not in st.session_state:
    st.session_state.df = None
if 'data_ref' not in st.session_state:
    st.session_state.data_ref = date.today()

# ── Header ────────────────────────────────────────────────────────────────────
col_logo, col_ref, col_up = st.columns([3, 2, 2])
with col_logo:
    st.markdown("### ODP &nbsp;·&nbsp; Privalia")
with col_ref:
    if st.session_state.df is not None:
        dr = st.session_state.data_ref
        st.markdown(f"<span style='font-size:13px;color:#6b7280'>ref. {dr.strftime('%d/%m/%Y')}</span>",
                    unsafe_allow_html=True)
with col_up:
    uploaded = st.file_uploader("", type=['xlsm','xlsx'], label_visibility='collapsed',
                                 key='uploader')
    if uploaded:
        with st.spinner("Lendo simulador..."):
            df_novo = carregar(uploaded.read(), uploaded.name)
            st.session_state.df = df_novo
            if not df_novo.empty:
                st.session_state.data_ref = df_novo['data_ref'].iloc[0]
        st.rerun()

st.markdown("<hr style='margin:0 0 16px;border:none;border-top:1px solid #f3f4f6'>",
            unsafe_allow_html=True)

# ── Sem dados ─────────────────────────────────────────────────────────────────
df = st.session_state.df
if df is None or df.empty:
    st.info("Faça upload do Simulador ODP (.xlsm) para começar.")
    st.stop()

data_ref  = st.session_state.data_ref
ontem     = data_ref - timedelta(days=1)

# partições principais
df_on    = df[df['status'] == 'ON'].copy()
df_ontem = df[(df['status'] == 'x') & (df['ed'].dt.date == ontem)].copy()
df_fut   = df[df['status'] == '-'].copy()

# ── ABAS ─────────────────────────────────────────────────────────────────────
aba_dash, aba_saidas, aba_blocado, aba_rfcst, aba_dados = st.tabs([
    "Dashboard", "Saídas", "Blocado", "Reforecast", "Dados"
])

# ═════════════════════════════════════════════════════════════════════════════
# ABA 1 — DASHBOARD
# ═════════════════════════════════════════════════════════════════════════════
with aba_dash:

    # ── Métricas de topo ──────────────────────────────────────────────────────
    pal_hoje = df_on['pallets'].sum() + df_ontem['pallets'].sum()
    alertas  = df_on[df_on['desvio_abs'].abs() > 0.15]

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Campanhas ON agora",   len(df_on))
    m2.metric("Encerradas ontem",     len(df_ontem))
    m3.metric("Pallets blocados hoje",fmt_n(pal_hoje),
              delta=f"{pal_hoje/CAP_BU*100:.0f}% da cap. BU",
              delta_color="off")
    m4.metric("Com desvio > ±15 p.p.",len(alertas),
              delta="requerem atenção" if len(alertas) else "tudo dentro do esperado",
              delta_color="inverse" if len(alertas) else "off")

    # ── Campanhas ativas ──────────────────────────────────────────────────────
    st.markdown('<div class="titulo-sec">Campanhas ativas agora</div>', unsafe_allow_html=True)
    _mostrar_encerradas = st.checkbox("Incluir encerradas ontem", value=False, key='dash_enc')
    df_view = pd.concat([df_on, df_ontem]) if _mostrar_encerradas else df_on

    for _, camp in df_view.sort_values('ed').iterrows():
        badge = ('badge-on'  if camp['status'] == 'ON'
                 else 'badge-x')
        alerta_str = ""
        if abs(camp['desvio_abs']) > 0.15:
            sinal = "+" if camp['desvio_abs'] > 0 else ""
            alerta_str = f'<span class="alerta">⚑ {sinal}{camp["desvio_abs"]*100:.1f} p.p.</span>'

        header_html = (
            f'<span class="{badge}">{camp["status"]}</span> &nbsp;'
            f'<b>{camp["campanha"]}</b> &nbsp;'
            f'<span style="color:#9ca3af;font-size:12px">'
            f'{camp["cat"]} · {camp["cd"]} · '
            f'{camp["sd"].strftime("%d/%m")} → {camp["ed"].strftime("%d/%m")} · '
            f'dia {camp["dia_atual"]} de {camp["wd"]}'
            f'</span> &nbsp; {alerta_str}'
        )

        with st.expander(camp['campanha'], expanded=False):
            st.markdown(header_html, unsafe_allow_html=True)
            c1,c2,c3,c4,c5,c6 = st.columns(6)
            c1.metric("Estoque",    fmt_n(camp['estoque']))
            c2.metric("Pallets",    fmt_n(camp['pallets']))
            c3.metric("Forecast",   fmt_n(camp['forecast']))
            c4.metric("SO congel.", fmt_pct(camp['so_cong']))
            c5.metric("Venda real", fmt_n(camp['venda_real']))
            c6.metric("SO real",    fmt_pct(camp['so_real_acum']))

            c7,c8,c9,c10 = st.columns(4)
            c7.metric("Reforecast", fmt_n(camp['reforecast']))
            c8.metric("Novo SO",    fmt_pct(camp['novo_so']))
            c9.metric("Descida DN", camp['dn'].strftime('%d/%m/%Y'))
            c10.metric("Subida DO", camp['do_'].strftime('%d/%m/%Y'))

            # barra de progresso
            prog = camp['dia_atual'] / camp['wd'] if camp['wd'] > 0 else 0
            st.progress(float(prog), text=f"Progresso: dia {camp['dia_atual']} / {camp['wd']}")

# ═════════════════════════════════════════════════════════════════════════════
# ABA 2 — SAÍDAS
# ═════════════════════════════════════════════════════════════════════════════
with aba_saidas:
    # ── Filtros ───────────────────────────────────────────────────────────────
    fc1, fc2, fc3, fc4 = st.columns([2,2,2,2])
    with fc1:
        opcoes_cd  = ['Todos'] + sorted(df['cd'].unique().tolist())
        sel_cd     = st.selectbox("CD", opcoes_cd, key='sai_cd')
    with fc2:
        opcoes_cat = ['Todas'] + sorted(df['cat'].unique().tolist())
        sel_cat    = st.selectbox("Categoria", opcoes_cat, key='sai_cat')
    with fc3:
        unidade    = st.radio("Unidade", ["Convertido","Bruto"], horizontal=True, key='sai_uni')
    with fc4:
        dt_range   = st.date_input("Período",
                     value=(data_ref - timedelta(days=7), data_ref + timedelta(days=30)),
                     key='sai_per')

    df_sai_base = df.copy()
    if sel_cd  != 'Todos':  df_sai_base = df_sai_base[df_sai_base['cd'] == sel_cd]
    if sel_cat != 'Todas':  df_sai_base = df_sai_base[df_sai_base['cat'] == sel_cat]

    dt_ini_s = dt_range[0] if len(dt_range) == 2 else data_ref - timedelta(days=7)
    dt_fim_s = dt_range[1] if len(dt_range) == 2 else data_ref + timedelta(days=30)

    # gera série — usa curva simplificada (uniforme)
    with st.spinner("Calculando saídas..."):
        uni_str = 'convertido' if unidade == 'Convertido' else 'bruto'
        serie_sai = gerar_serie_saidas(df_sai_base, dt_ini_s, dt_fim_s, uni_str)

    if not serie_sai.empty:
        # agrega por data e tipo
        pivot = (serie_sai.groupby(['data','tipo'])['volume']
                 .sum().unstack(fill_value=0).reset_index())
        for t in ['realizado','reforecast','previsto']:
            if t not in pivot.columns:
                pivot[t] = 0

        fig_sai = go.Figure()
        for tipo, cor, label in [
            ('realizado',  CORES_TIPO['realizado'],  'Realizado'),
            ('reforecast', CORES_TIPO['reforecast'], 'Reforecast'),
            ('previsto',   CORES_TIPO['previsto'],   'Previsto'),
        ]:
            fig_sai.add_bar(x=pivot['data'], y=pivot[tipo],
                            name=label, marker_color=cor)

        fig_sai.update_layout(
            barmode='stack', height=320,
            margin=dict(l=0,r=0,t=10,b=0),
            legend=dict(orientation='h', y=1.08),
            xaxis=dict(showgrid=False),
            yaxis=dict(title=f"Peças ({unidade.lower()})", gridcolor='#f3f4f6'),
            plot_bgcolor='white', paper_bgcolor='white',
            font=dict(family='DM Sans', size=12),
        )
        # linha de capacidade check-out
        cap_val = CAP_CHECKOUT if uni_str == 'convertido' else CAP_CHECKIN
        fig_sai.add_hline(y=cap_val, line_dash='dot', line_color='#dc2626',
                          annotation_text=f"Cap. {cap_val:,.0f}", annotation_font_size=11)
        st.plotly_chart(fig_sai, use_container_width=True)

    # ── Tabela por campanha ───────────────────────────────────────────────────
    st.markdown('<div class="titulo-sec">Detalhe por campanha</div>', unsafe_allow_html=True)
    for _, camp in df_sai_base.sort_values('sd').iterrows():
        with st.expander(f"{camp['campanha']} · {camp['cat']} · {camp['sd'].strftime('%d/%m')}→{camp['ed'].strftime('%d/%m')}"):
            c1,c2,c3,c4 = st.columns(4)
            c1.metric("Reforecast total", fmt_n(camp['reforecast']))
            c2.metric("Fat. OUT",         f"{camp['fat_out']:.2f}×")
            c3.metric("Curva usada",      f"{camp['cat']} {camp['wd']}wd")
            c4.metric("Dias restantes",   fmt_n(camp['dias_rest']))

            # mini série desta campanha
            s_camp = serie_sai[serie_sai['id'] == camp['id']]
            if not s_camp.empty:
                s_p = s_camp.groupby(['data','tipo'])['volume'].sum().unstack(fill_value=0).reset_index()
                for t in ['realizado','reforecast','previsto']:
                    if t not in s_p.columns: s_p[t] = 0
                fig_c = go.Figure()
                for tipo, cor in [('realizado','#1e3a5f'),('reforecast','#3b82f6'),('previsto','#bfdbfe')]:
                    fig_c.add_bar(x=s_p['data'], y=s_p[tipo], name=tipo, marker_color=cor)
                fig_c.update_layout(barmode='stack', height=180,
                                    margin=dict(l=0,r=0,t=5,b=0),
                                    showlegend=False, plot_bgcolor='white',
                                    paper_bgcolor='white',
                                    xaxis=dict(showgrid=False),
                                    yaxis=dict(gridcolor='#f3f4f6'))
                st.plotly_chart(fig_c, use_container_width=True)

    # ── Exportar ──────────────────────────────────────────────────────────────
    if not serie_sai.empty:
        csv = (serie_sai.groupby(['data','campanha','cat','cd','tipo'])['volume']
               .sum().reset_index().to_csv(index=False, sep=';').encode('utf-8-sig'))
        st.download_button("⬇ Exportar tabela de saídas (.csv)", csv,
                           "saidas_odp.csv", "text/csv")

# ═════════════════════════════════════════════════════════════════════════════
# ABA 3 — BLOCADO
# ═════════════════════════════════════════════════════════════════════════════
with aba_blocado:
    # ── Filtros ───────────────────────────────────────────────────────────────
    fb1, fb2, fb3 = st.columns([2,2,3])
    with fb1:
        sel_cd_b = st.selectbox("CD", ['Todos'] + sorted(df['cd'].unique().tolist()), key='bl_cd')
    with fb2:
        uni_b = st.radio("Unidade", ["Pallets","Peças brutas","Peças conv."], horizontal=True, key='bl_uni')
    with fb3:
        dt_range_b = st.date_input("Período",
                     value=(data_ref - timedelta(days=7), data_ref + timedelta(days=45)),
                     key='bl_per')

    df_bl = df.copy()
    if sel_cd_b != 'Todos':
        df_bl = df_bl[df_bl['cd'] == sel_cd_b]

    dt_ini_b = dt_range_b[0] if len(dt_range_b) == 2 else data_ref - timedelta(days=7)
    dt_fim_b = dt_range_b[1] if len(dt_range_b) == 2 else data_ref + timedelta(days=45)

    # série de blocado
    with st.spinner("Calculando blocado..."):
        serie_bl = gerar_serie_blocado(df_bl, dt_ini_b, dt_fim_b)

    if uni_b == "Peças brutas":
        col_vol = 'pcs_brutas'
        # recalcular com peças brutas (pallets * pcs_pal por campanha)
        rows_bl2 = []
        datas_bl = pd.date_range(dt_ini_b, dt_fim_b, freq='D')
        for dt in datas_bl:
            dt_d = dt.date()
            ativas = df_bl[(df_bl['dn'].dt.date <= dt_d) & (df_bl['do_'].dt.date >= dt_d)]
            rows_bl2.append({'data': dt, 'vol': ativas['estoque'].sum()})
        serie_plot = pd.DataFrame(rows_bl2)
        cap_linha  = None
        ytitle     = "Peças brutas"
    elif uni_b == "Peças conv.":
        rows_bl3 = []
        datas_bl = pd.date_range(dt_ini_b, dt_fim_b, freq='D')
        for dt in datas_bl:
            dt_d = dt.date()
            ativas = df_bl[(df_bl['dn'].dt.date <= dt_d) & (df_bl['do_'].dt.date >= dt_d)]
            vol = (ativas['estoque'] / ativas['fat_in'].replace(0,1)).sum()
            rows_bl3.append({'data': dt, 'vol': vol})
        serie_plot = pd.DataFrame(rows_bl3)
        cap_linha  = CAP_CHECKOUT
        ytitle     = "Peças convertidas"
    else:
        serie_plot = serie_bl.rename(columns={'pallets':'vol'})
        cap_linha  = CAP_BU
        ytitle     = "Pallets"

    if not serie_plot.empty:
        fig_bl = go.Figure()
        fig_bl.add_scatter(x=serie_plot['data'], y=serie_plot['vol'],
                           fill='tozeroy', fillcolor='rgba(59,130,246,0.15)',
                           line=dict(color='#3b82f6', width=2),
                           name=ytitle, hovertemplate='%{x|%d/%m}<br>%{y:,.0f}<extra></extra>')
        if cap_linha:
            fig_bl.add_hline(y=cap_linha, line_dash='dot', line_color='#dc2626',
                             annotation_text=f"Cap. {cap_linha:,}", annotation_font_size=11)
        fig_bl.update_layout(
            height=320, margin=dict(l=0,r=0,t=10,b=0),
            xaxis=dict(showgrid=False),
            yaxis=dict(title=ytitle, gridcolor='#f3f4f6'),
            plot_bgcolor='white', paper_bgcolor='white',
            font=dict(family='DM Sans', size=12),
        )
        st.plotly_chart(fig_bl, use_container_width=True)

    # ── Métricas ──────────────────────────────────────────────────────────────
    pico_pal = serie_bl['pallets'].max() if not serie_bl.empty else 0
    desc_sem = df_bl[(df_bl['dn'].dt.date >= data_ref) &
                     (df_bl['dn'].dt.date <= data_ref + timedelta(days=7))]
    sub_sem  = df_bl[(df_bl['do_'].dt.date >= data_ref) &
                     (df_bl['do_'].dt.date <= data_ref + timedelta(days=7))]
    mb1,mb2,mb3 = st.columns(3)
    mb1.metric("Pico próx. 45 dias", fmt_n(pico_pal) + " pal",
               delta=f"{pico_pal/CAP_BU*100:.0f}% da cap. BU", delta_color="off")
    mb2.metric("Descidas esta semana", len(desc_sem),
               delta="campanhas entram no picking")
    mb3.metric("Subidas esta semana",  len(sub_sem),
               delta="campanhas saem do picking")

    # ── Tabela por campanha ───────────────────────────────────────────────────
    st.markdown('<div class="titulo-sec">Detalhe por campanha</div>', unsafe_allow_html=True)
    df_bl_sort = df_bl.sort_values('dn')
    for _, camp in df_bl_sort.iterrows():
        status_badge = {'ON':'🟢','x':'⚪','-':'🔵'}.get(camp['status'],'')
        with st.expander(f"{status_badge} {camp['campanha']} · {camp['cat']} · DN {camp['dn'].strftime('%d/%m')} → DO {camp['do_'].strftime('%d/%m')}"):
            c1,c2,c3,c4,c5 = st.columns(5)
            c1.metric("Pallets",    fmt_n(camp['pallets']))
            c2.metric("Pç brutas",  fmt_n(camp['estoque']))
            c3.metric("Pç/palete",  fmt_n(camp['pcs_pal']))
            c4.metric("Fat. IN",    f"{camp['fat_in']:.2f}×")
            c5.metric("CD",         camp['cd'])

    # ── Exportar ──────────────────────────────────────────────────────────────
    exp_bl = df_bl[['id','campanha','cat','setor','cd','status','sd','ed','dn','do_',
                     'estoque','pallets','pcs_pal','fat_in','fat_out']].copy()
    exp_bl.columns = ['ID','Campanha','Categoria','Setor','CD','Status',
                       'SD','ED','DN','DO','Estoque','Pallets','Pcs/Pal','Fat IN','Fat OUT']
    csv_bl = exp_bl.to_csv(index=False, sep=';').encode('utf-8-sig')
    st.download_button("⬇ Exportar tabela de blocado (.csv)", csv_bl,
                       "blocado_odp.csv", "text/csv")

# ═════════════════════════════════════════════════════════════════════════════
# ABA 4 — REFORECAST
# ═════════════════════════════════════════════════════════════════════════════
with aba_rfcst:
    # ── Filtros ───────────────────────────────────────────────────────────────
    fr1, fr2, fr3 = st.columns([2,2,2])
    with fr1:
        mostrar_rf = st.radio("Campanhas", ["Só ativas (ON)","Todas"], horizontal=True, key='rf_show')
    with fr2:
        so_alerta_rf = st.radio("Filtro", ["Todas","Só com desvio >15 p.p."], horizontal=True, key='rf_fil')
    with fr3:
        sel_cd_rf = st.selectbox("CD", ['Todos'] + sorted(df['cd'].unique().tolist()), key='rf_cd')

    df_rf = df.copy()
    if mostrar_rf == "Só ativas (ON)":
        df_rf = df_rf[df_rf['status'] == 'ON']
    if so_alerta_rf == "Só com desvio >15 p.p.":
        df_rf = df_rf[df_rf['desvio_abs'].abs() > 0.15]
    if sel_cd_rf != 'Todos':
        df_rf = df_rf[df_rf['cd'] == sel_cd_rf]

    df_rf = df_rf.sort_values('desvio_abs', key=abs, ascending=False)

    st.markdown('<div class="titulo-sec">Racional de reforecast por campanha</div>',
                unsafe_allow_html=True)

    for _, camp in df_rf.iterrows():
        desvio_cor = cor_desvio(camp['desvio_abs'])
        sinal = "+" if camp['desvio_abs'] >= 0 else ""
        status_str = {'ON':'🟢','x':'⚪','-':'🔵'}.get(camp['status'],'')
        badge_dev = (f'<span style="color:{desvio_cor};font-weight:600">'
                     f'{sinal}{camp["desvio_abs"]*100:.1f} p.p.</span>')

        with st.expander(
            f"{status_str} {camp['campanha']}  ·  "
            f"dia {camp['dia_atual']}/{camp['wd']}  ·  "
            f"desvio {sinal}{camp['desvio_abs']*100:.1f} p.p."
        ):
            # barra dupla: previsto vs real
            prog_prev = float(min(1, camp['so_prev_acum']))
            prog_real = float(min(1, camp['so_real_acum']))

            st.markdown(f"""
            <div style='margin-bottom:12px'>
              <div style='display:flex;justify-content:space-between;font-size:12px;color:#6b7280;margin-bottom:4px'>
                <span>SO previsto acum.</span><span>{prog_prev*100:.1f}%</span>
              </div>
              <div style='background:#f3f4f6;border-radius:4px;height:8px'>
                <div style='width:{prog_prev*100:.1f}%;background:#93c5fd;border-radius:4px;height:100%'></div>
              </div>
              <div style='display:flex;justify-content:space-between;font-size:12px;color:#6b7280;margin:8px 0 4px'>
                <span>SO real acum.</span>
                <span style='color:{desvio_cor};font-weight:600'>{prog_real*100:.1f}% ({sinal}{camp["desvio_abs"]*100:.1f} p.p.)</span>
              </div>
              <div style='background:#f3f4f6;border-radius:4px;height:8px'>
                <div style='width:{prog_real*100:.1f}%;background:{desvio_cor};border-radius:4px;height:100%'></div>
              </div>
            </div>
            """, unsafe_allow_html=True)

            c1,c2,c3,c4,c5,c6 = st.columns(6)
            c1.metric("Venda real",   fmt_n(camp['venda_real']))
            c2.metric("Reforecast",   fmt_n(camp['reforecast']))
            c3.metric("Var vs FC",    fmt_pct(camp['mudanca_vs_fc']))
            c4.metric("Novo SO",      fmt_pct(camp['novo_so']))
            c5.metric("SO congel.",   fmt_pct(camp['so_cong']))
            c6.metric("Fator ajuste", f"{camp['fator_ajuste']:.2f}×")

            # funil QB→QR
            st.markdown("**Racional completo (QB → QR)**")
            tb = {
                "QB · Dias de venda":                 camp['dias_venda'],
                "QC · Dias restantes":                camp['dias_rest'],
                "QD · Proj. acumulada até D-1":       fmt_pct(camp['proj_acum']),
                "QE · SO previsto acum.":             fmt_pct(camp['so_prev_acum']),
                "QF · Venda real acum.":              fmt_n(camp['venda_real']),
                "QG · SO real acum.":                 fmt_pct(camp['so_real_acum']),
                "QH · Desvio abs (SO real–prev)":     f"{camp['desvio_abs']*100:+.2f} p.p.",
                "QI · Desvio rel (SO real/prev – 1)": fmt_pct(camp['desvio_rel']),
                "QJ · Desvio médio":                  fmt_pct((camp['desvio_abs']+camp['desvio_rel'])/2),
                "QK · Fator de ajuste (1+QJ)":        f"{camp['fator_ajuste']:.3f}×",
                "QL · % restante da curva":           fmt_pct(1-camp['proj_acum']),
                "QO · Reforecast final":              fmt_n(camp['reforecast']),
                "QP · Var vs Forecast congelado":     fmt_pct(camp['mudanca_vs_fc']),
                "QQ · Novo Sell-out (estoque)":       fmt_pct(camp['novo_so']),
                "QR · Var vs SO original":            f"{camp['desvio_vs_so']*100:+.2f} p.p.",
            }
            st.table(pd.DataFrame.from_dict(tb, orient='index', columns=['Valor']))

    # ── Exportar ──────────────────────────────────────────────────────────────
    cols_exp = ['id','campanha','cat','cd','status','sd','ed','wd',
                'estoque','forecast','so_cong','venda_real','so_real_acum',
                'desvio_abs','fator_ajuste','reforecast','novo_so','mudanca_vs_fc']
    exp_rf = df_rf[cols_exp].copy()
    exp_rf.columns = ['ID','Campanha','Categoria','CD','Status','SD','ED','WD',
                       'Estoque','Forecast','SO Congel.','Venda Real','SO Real',
                       'Desvio (p.p.)','Fator Ajuste','Reforecast','Novo SO','Var vs FC']
    csv_rf = exp_rf.to_csv(index=False, sep=';').encode('utf-8-sig')
    st.download_button("⬇ Exportar tabela de reforecast (.csv)", csv_rf,
                       "reforecast_odp.csv", "text/csv")

# ═════════════════════════════════════════════════════════════════════════════
# ABA 5 — DADOS
# ═════════════════════════════════════════════════════════════════════════════
with aba_dados:
    st.markdown("#### Upload do Simulador ODP")
    st.markdown("""
    <div style='border:1.5px dashed #d1d5db;border-radius:12px;padding:24px;
                text-align:center;color:#6b7280;font-size:13px;margin-bottom:20px'>
        Use o botão de upload no topo da página para carregar o arquivo <b>.xlsm</b>.<br>
        <span style='font-size:11px'>O arquivo permanece apenas na sessão — não é armazenado.</span>
    </div>
    """, unsafe_allow_html=True)

    if df is not None and not df.empty:
        st.markdown("#### Arquivo carregado")
        cdi1,cdi2,cdi3 = st.columns(3)
        cdi1.metric("Total campanhas ODP", len(df))
        cdi2.metric("Campanhas ON",  len(df_on))
        cdi3.metric("Data referência", data_ref.strftime('%d/%m/%Y'))

        st.markdown('<div class="titulo-sec">Resumo por status</div>', unsafe_allow_html=True)
        resumo = (df.groupby('status')
                  .agg(campanhas=('id','count'),
                       estoque_total=('estoque','sum'),
                       pallets_total=('pallets','sum'))
                  .reset_index()
                  .rename(columns={'status':'Status','campanhas':'Campanhas',
                                   'estoque_total':'Estoque Total','pallets_total':'Pallets Total'}))
        st.dataframe(resumo, use_container_width=True, hide_index=True)

        st.markdown('<div class="titulo-sec">Tabela completa de campanhas</div>', unsafe_allow_html=True)
        cols_vis = ['id','campanha','status','cat','cd','sd','ed','wd',
                    'estoque','pallets','forecast','so_cong','reforecast','novo_so']
        df_vis = df[cols_vis].copy()
        df_vis.columns = ['ID','Campanha','Status','Categoria','CD','SD','ED','WD',
                           'Estoque','Pallets','Forecast','SO Congel.','Reforecast','Novo SO']
        st.dataframe(df_vis, use_container_width=True, hide_index=True)

        csv_all = df_vis.to_csv(index=False, sep=';').encode('utf-8-sig')
        st.download_button("⬇ Exportar todas as campanhas (.csv)", csv_all,
                           "campanhas_odp.csv", "text/csv")
