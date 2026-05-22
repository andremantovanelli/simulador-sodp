import streamlit as st
import pandas as pd
from datetime import date, timedelta
import io

st.set_page_config(
    page_title="Simulador S&OP — Privalia",
    page_icon="📦",
    layout="wide"
)

# ── Funções de carregamento ──────────────────────────────────────────────────

@st.cache_data
def carregar_calendario(arquivo):
    df = pd.read_excel(arquivo)
    df = df.rename(columns={
        'Id externo de Campanha Sacarino': 'id_campanha',
        'Nome da campanha':                'campanha',
        'Data de início':                  'data_inicio',
        'Data de término':                 'data_fim',
        'Webdays':                         'webdays',
        'Status':                          'status',
        'Centro de distibuição':           'cd',
        'Categoria':                       'categoria',
        'Sector Calendar':                 'setor',
        'Previsão de venda peças':         'previsao_pecas',
        'Forecast':                        'forecast_receita',
        'Estoque Total':                   'estoque_total',
        'Modelo de negócio':               'modelo_negocio',
        'Gerência':                        'gerencia',
        'PVS médio':                       'pvs_medio',
        'Tipologia':                       'tipologia',
    })
    df['data_inicio'] = pd.to_datetime(df['data_inicio'], dayfirst=True, errors='coerce')
    df['data_fim']    = pd.to_datetime(df['data_fim'],    dayfirst=True, errors='coerce')

    # Deduplica: mesmo id_campanha pode ter múltiplos fornecedores — fica uma linha por campanha
    colunas_campanha = ['id_campanha', 'campanha', 'data_inicio', 'data_fim',
                        'webdays', 'status', 'cd', 'categoria', 'setor',
                        'previsao_pecas', 'forecast_receita', 'estoque_total',
                        'modelo_negocio', 'gerencia', 'pvs_medio', 'tipologia']
    colunas_existentes = [c for c in colunas_campanha if c in df.columns]
    df = df[colunas_existentes].drop_duplicates(subset=['id_campanha'])
    df = df.dropna(subset=['data_inicio', 'data_fim'])
    return df

@st.cache_data
def carregar_vendas(arquivo):
    df = pd.read_csv(arquivo) if str(arquivo.name).endswith('.csv') else pd.read_excel(arquivo)
    df.columns = [c.strip() for c in df.columns]
    # Mapeia nomes reais do arquivo (Base Vendas Dia)
    renomear = {
        'ID':           'id_campanha',
        'Dia_click':    'data_venda',
        'Items':        'pecas_vendidas',
        'Orders':       'pedidos',
        'Revenue':      'receita',
        'Setor':        'setor',
        'Start_Date':   'data_inicio_camp',
        'End_Date':     'data_fim_camp',
    }
    df = df.rename(columns={k: v for k, v in renomear.items() if k in df.columns})
    if 'data_venda' in df.columns:
        df['data_venda'] = pd.to_datetime(df['data_venda'], errors='coerce')
    return df

def calcular_reforecast(calendario, vendas, data_hoje):
    """
    Para cada campanha e cada dia de vigência:
    - Se o dia já passou (realizado): usa venda real da base
    - Se a campanha está no ar hoje: misto (real até ontem + previsão proporcional nos dias restantes)
    - Se a campanha ainda não começou: usa previsão proporcional (peças / webdays)
    """
    linhas = []
    for _, camp in calendario.iterrows():
        inicio  = camp['data_inicio'].date()
        fim     = camp['data_fim'].date()
        webdays = int(camp['webdays']) if pd.notna(camp['webdays']) and camp['webdays'] > 0 else 1
        previsao_total = camp['previsao_pecas'] if pd.notna(camp['previsao_pecas']) else 0
        previsao_dia   = previsao_total / webdays

        datas = pd.date_range(inicio, fim).date

        vendas_camp = pd.DataFrame()
        if vendas is not None and 'id_campanha' in vendas.columns:
            vendas_camp = vendas[vendas['id_campanha'] == camp['id_campanha']].copy()

        for d in datas:
            if d < data_hoje:
                # Realizado
                if not vendas_camp.empty and 'data_venda' in vendas_camp.columns:
                    vd = vendas_camp[vendas_camp['data_venda'].dt.date == d]
                    pecas = float(vd['pecas_vendidas'].sum()) if not vd.empty else previsao_dia
                else:
                    pecas = previsao_dia
                tipo = 'realizado'
            elif d == data_hoje:
                # Misto — realizado se já disponível, senão previsão
                if not vendas_camp.empty and 'data_venda' in vendas_camp.columns:
                    vd = vendas_camp[vendas_camp['data_venda'].dt.date == d]
                    pecas = float(vd['pecas_vendidas'].sum()) if not vd.empty else previsao_dia
                else:
                    pecas = previsao_dia
                tipo = 'reforecast'
            else:
                # Futuro — previsão
                pecas = previsao_dia
                tipo = 'previsao'

            linhas.append({
                'id_campanha':   camp['id_campanha'],
                'campanha':      camp['campanha'],
                'data':          d,
                'modelo_negocio': camp.get('modelo_negocio', ''),
                'gerencia':      camp.get('gerencia', ''),
                'cd':            camp.get('cd', ''),
                'pecas':         round(pecas, 1),
                'tipo':          tipo,
            })

    return pd.DataFrame(linhas)

# ── Interface ────────────────────────────────────────────────────────────────

st.title("Simulador S&OP — Privalia")

with st.sidebar:
    st.header("Carregar dados")
    arq_calendario = st.file_uploader(
        "Calendário (Salesforce)", type=["xlsx", "xls"],
        help="Exporte do Salesforce e faça upload aqui"
    )
    arq_vendas = st.file_uploader(
        "Base de vendas (DBeaver)", type=["csv", "xlsx"],
        help="CSV ou Excel exportado da query do data lake"
    )
    st.divider()
    data_hoje = st.date_input("Data de referência (hoje)", value=date.today())
    st.caption("Altere para simular outros cenários sem salvar.")

# ── Processamento ────────────────────────────────────────────────────────────

if arq_calendario is None:
    st.info("Faça upload do calendário do Salesforce na barra lateral para começar.")
    st.stop()

with st.spinner("Carregando calendário..."):
    calendario = carregar_calendario(arq_calendario)

vendas = None
if arq_vendas is not None:
    with st.spinner("Carregando base de vendas..."):
        vendas = carregar_vendas(arq_vendas)

with st.spinner("Calculando reforecast..."):
    rf = calcular_reforecast(calendario, vendas, data_hoje)

# ── Filtros ──────────────────────────────────────────────────────────────────

st.subheader("Filtros")
col1, col2, col3 = st.columns(3)

modelos     = sorted(rf['modelo_negocio'].dropna().unique().tolist())
gerencias   = sorted(rf['gerencia'].dropna().unique().tolist())
cds         = sorted(rf['cd'].dropna().unique().tolist())

with col1:
    sel_modelo  = st.multiselect("Modelo de negócio", modelos, default=modelos)
with col2:
    sel_gerencia = st.multiselect("Gerência", gerencias, default=gerencias)
with col3:
    sel_cd      = st.multiselect("CD", cds, default=cds)

rf_filtrado = rf[
    rf['modelo_negocio'].isin(sel_modelo) &
    rf['gerencia'].isin(sel_gerencia) &
    rf['cd'].isin(sel_cd)
]

# ── Volume dia a dia ─────────────────────────────────────────────────────────

st.subheader("Volume dia a dia (peças)")

volume_dia = (
    rf_filtrado
    .groupby(['data', 'tipo'])['pecas']
    .sum()
    .reset_index()
)

# Pivot para visualização
pivot = volume_dia.pivot_table(index='data', columns='tipo', values='pecas', aggfunc='sum').fillna(0)
pivot.index = pd.to_datetime(pivot.index)
pivot = pivot.sort_index()

# Renomeia colunas para exibição amigável
pivot.columns.name = None
pivot = pivot.rename(columns={
    'realizado':   'Realizado',
    'reforecast':  'Reforecast (misto)',
    'previsao':    'Previsão',
})

st.bar_chart(pivot, use_container_width=True)

# ── Tabela resumo mensal ──────────────────────────────────────────────────────

st.subheader("Resumo por gerência e modelo de negócio")

resumo = (
    rf_filtrado
    .groupby(['gerencia', 'modelo_negocio'])
    .agg(
        campanhas=('id_campanha', 'nunique'),
        pecas_total=('pecas', 'sum'),
    )
    .reset_index()
    .rename(columns={
        'gerencia':        'Gerência',
        'modelo_negocio':  'Modelo',
        'campanhas':       'Campanhas',
        'pecas_total':     'Peças (total)',
    })
)
resumo['Peças (total)'] = resumo['Peças (total)'].round(0).astype(int)
st.dataframe(resumo, use_container_width=True, hide_index=True)

# ── Calendário de campanhas ───────────────────────────────────────────────────

st.subheader("Campanhas no período")

cal_exib = calendario[
    calendario['modelo_negocio'].isin(sel_modelo) &
    calendario['gerencia'].isin(sel_gerencia) &
    calendario['cd'].isin(sel_cd)
].copy()

cal_exib['data_inicio'] = cal_exib['data_inicio'].dt.strftime('%d/%m/%Y')
cal_exib['data_fim']    = cal_exib['data_fim'].dt.strftime('%d/%m/%Y')
cal_exib['previsao_pecas'] = cal_exib['previsao_pecas'].fillna(0).astype(int)

colunas_exib = ['campanha', 'data_inicio', 'data_fim', 'webdays',
                'modelo_negocio', 'gerencia', 'cd',
                'previsao_pecas', 'estoque_total', 'status']
colunas_exib = [c for c in colunas_exib if c in cal_exib.columns]

st.dataframe(
    cal_exib[colunas_exib].rename(columns={
        'campanha':        'Campanha',
        'data_inicio':     'Início',
        'data_fim':        'Fim',
        'webdays':         'Dias web',
        'modelo_negocio':  'Modelo',
        'gerencia':        'Gerência',
        'cd':              'CD',
        'previsao_pecas':  'Previsão (pçs)',
        'estoque_total':   'Estoque total',
        'status':          'Status',
    }),
    use_container_width=True,
    hide_index=True
)

# ── Export ────────────────────────────────────────────────────────────────────

st.subheader("Exportar")

col_exp1, col_exp2 = st.columns(2)

with col_exp1:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        pivot.to_excel(writer, sheet_name='Volume dia a dia')
        resumo.to_excel(writer, sheet_name='Resumo gerência', index=False)
        cal_exib[colunas_exib].to_excel(writer, sheet_name='Calendário', index=False)
    st.download_button(
        label="Baixar Excel",
        data=buf.getvalue(),
        file_name=f"simulador_sodp_{date.today().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

with col_exp2:
    csv_bytes = rf_filtrado.to_csv(index=False).encode('utf-8')
    st.download_button(
        label="Baixar CSV completo",
        data=csv_bytes,
        file_name=f"reforecast_{date.today().strftime('%Y%m%d')}.csv",
        mime="text/csv"
    )
