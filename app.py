import streamlit as st
import pandas as pd
from datetime import date
import io

st.set_page_config(
    page_title="Simulador S&OP — Privalia",
    page_icon="📦",
    layout="wide"
)

# ── Funções utilitárias ──────────────────────────────────────────────────────

def parse_numero_br(serie):
    """Converte números no formato brasileiro (1.234,56) para float."""
    if serie.dtype == object:
        serie = (serie.astype(str)
                 .str.replace('.', '', regex=False)
                 .str.replace(',', '.', regex=False))
    return pd.to_numeric(serie, errors='coerce')

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
    """
    CSV exportado do DBeaver — separador ponto e vírgula, números em formato BR.
    Colunas esperadas: ID, Campanha, Dia_Click, Start_Date, End_Date,
                       relative_day, Orders, Items, Revenue, Setor, Modelo_de_Negocio
    """
    df = pd.read_csv(arquivo, sep=';', dtype=str)
    df.columns = [c.strip() for c in df.columns]

    renomear = {
        'ID':                'id_campanha',
        'Campanha':          'campanha',
        'Dia_Click':         'data_venda',
        'Start_Date':        'data_inicio_camp',
        'End_Date':          'data_fim_camp',
        'relative_day':      'dia_relativo',
        'Orders':            'pedidos',
        'Items':             'pecas_vendidas',
        'Revenue':           'receita',
        'Setor':             'setor',
        'Modelo_de_Negocio': 'modelo_negocio',
        'Initial_Stock':     'estoque_inicial',
        'Stock_Sold':        'estoque_vendido',
    }
    df = df.rename(columns={k: v for k, v in renomear.items() if k in df.columns})

    # ID: remove pontos de milhar que podem aparecer (ex: 257.382 → 257382)
    if 'id_campanha' in df.columns:
        df['id_campanha'] = df['id_campanha'].str.replace('.', '', regex=False)
        df['id_campanha'] = pd.to_numeric(df['id_campanha'], errors='coerce')

    # Datas
    for col in ['data_venda', 'data_inicio_camp', 'data_fim_camp']:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], dayfirst=True, errors='coerce')

    # Números em formato BR
    for col in ['pecas_vendidas', 'pedidos', 'receita', 'estoque_inicial', 'estoque_vendido']:
        if col in df.columns:
            df[col] = parse_numero_br(df[col])

    return df

@st.cache_data
def carregar_odp_vertical(arquivo):
    """
    ODP Vertical — entradas planejadas de estoque no CD por campanha e data.
    Colunas: ID campanha, Nome, Categoria, Peças, Peças Convertidas, Status, Data
    As colunas Unnamed: são colunas extras vazias — descartamos.
    """
    df = pd.read_excel(arquivo)
    # Mantém só as colunas nomeadas úteis
    colunas_uteis = ['ID campanha', 'Nome', 'Categoria', 'Peças',
                     'Peças Convertidas', 'Status', 'Data']
    colunas_existentes = [c for c in colunas_uteis if c in df.columns]
    df = df[colunas_existentes].copy()

    df = df.rename(columns={
        'ID campanha':      'id_campanha',
        'Nome':             'campanha',
        'Categoria':        'categoria',
        'Peças':            'pecas',
        'Peças Convertidas':'pecas_convertidas',
        'Status':           'status_entrada',
        'Data':             'data_entrada',
    })
    df['data_entrada'] = pd.to_datetime(df['data_entrada'], errors='coerce')
    df = df.dropna(subset=['data_entrada'])
    return df

# ── Reforecast ───────────────────────────────────────────────────────────────

def calcular_reforecast(calendario, vendas, data_hoje):
    linhas = []
    for _, camp in calendario.iterrows():
        inicio  = camp['data_inicio'].date()
        fim     = camp['data_fim'].date()
        webdays = int(camp['webdays']) if pd.notna(camp.get('webdays')) and camp.get('webdays', 0) > 0 else 1
        previsao_total = camp['previsao_pecas'] if pd.notna(camp.get('previsao_pecas')) else 0
        previsao_dia   = previsao_total / webdays

        datas = pd.date_range(inicio, fim).date

        vendas_camp = pd.DataFrame()
        if vendas is not None and 'id_campanha' in vendas.columns:
            vendas_camp = vendas[vendas['id_campanha'] == camp['id_campanha']].copy()

        for d in datas:
            if d < data_hoje:
                if not vendas_camp.empty and 'data_venda' in vendas_camp.columns:
                    vd = vendas_camp[vendas_camp['data_venda'].dt.date == d]
                    pecas = float(vd['pecas_vendidas'].sum()) if not vd.empty else previsao_dia
                else:
                    pecas = previsao_dia
                tipo = 'realizado'
            elif d == data_hoje:
                if not vendas_camp.empty and 'data_venda' in vendas_camp.columns:
                    vd = vendas_camp[vendas_camp['data_venda'].dt.date == d]
                    pecas = float(vd['pecas_vendidas'].sum()) if not vd.empty else previsao_dia
                else:
                    pecas = previsao_dia
                tipo = 'reforecast'
            else:
                pecas = previsao_dia
                tipo = 'previsao'

            linhas.append({
                'id_campanha':    camp['id_campanha'],
                'campanha':       camp['campanha'],
                'data':           d,
                'modelo_negocio': camp.get('modelo_negocio', ''),
                'gerencia':       camp.get('gerencia', ''),
                'cd':             camp.get('cd', ''),
                'pecas':          round(pecas, 1),
                'tipo':           tipo,
            })

    return pd.DataFrame(linhas)

# ── Interface ────────────────────────────────────────────────────────────────

st.title("Simulador S&OP — Privalia")

with st.sidebar:
    st.header("Carregar dados")

    arq_calendario = st.file_uploader(
        "1. Calendário (Salesforce)", type=["xlsx", "xls"],
        help="Export do Salesforce — .xlsx"
    )
    arq_vendas = st.file_uploader(
        "2. Base de vendas (DBeaver)", type=["csv"],
        help="CSV exportado da query — separador ponto e vírgula"
    )
    arq_odp = st.file_uploader(
        "3. ODP Vertical", type=["xlsx", "xls"],
        help="Planilha de entradas planejadas no CD"
    )

    st.divider()
    data_hoje = st.date_input("Data de referência (hoje)", value=date.today())
    st.caption("Altere para simular cenários. Nada é salvo.")

# ── Carregamento ─────────────────────────────────────────────────────────────

if arq_calendario is None:
    st.info("Faça upload do calendário do Salesforce na barra lateral para começar.")
    st.stop()

with st.spinner("Carregando calendário..."):
    calendario = carregar_calendario(arq_calendario)

vendas = None
if arq_vendas is not None:
    with st.spinner("Carregando base de vendas..."):
        try:
            vendas = carregar_vendas(arq_vendas)
            st.sidebar.success(f"Vendas: {len(vendas):,} registros carregados")
        except Exception as e:
            st.sidebar.error(f"Erro ao carregar vendas: {e}")

odp = None
if arq_odp is not None:
    with st.spinner("Carregando ODP Vertical..."):
        try:
            odp = carregar_odp_vertical(arq_odp)
            st.sidebar.success(f"ODP Vertical: {len(odp):,} registros carregados")
        except Exception as e:
            st.sidebar.error(f"Erro ao carregar ODP: {e}")

with st.spinner("Calculando reforecast..."):
    rf = calcular_reforecast(calendario, vendas, data_hoje)

# ── Filtros ───────────────────────────────────────────────────────────────────

st.subheader("Filtros")
col1, col2, col3 = st.columns(3)

modelos   = sorted(rf['modelo_negocio'].dropna().unique().tolist())
gerencias = sorted(rf['gerencia'].dropna().unique().tolist())
cds       = sorted(rf['cd'].dropna().unique().tolist())

with col1:
    sel_modelo   = st.multiselect("Modelo de negócio", modelos, default=modelos)
with col2:
    sel_gerencia = st.multiselect("Gerência", gerencias, default=gerencias)
with col3:
    sel_cd       = st.multiselect("CD", cds, default=cds)

rf_filtrado = rf[
    rf['modelo_negocio'].isin(sel_modelo) &
    rf['gerencia'].isin(sel_gerencia) &
    rf['cd'].isin(sel_cd)
]

# ── Abas principais ───────────────────────────────────────────────────────────

aba1, aba2, aba3, aba4 = st.tabs([
    "Reforecast de vendas",
    "Entradas ODP Vertical",
    "Calendário de campanhas",
    "Exportar"
])

# ─── ABA 1: Reforecast ────────────────────────────────────────────────────────
with aba1:
    st.subheader("Volume de vendas dia a dia (peças)")

    volume_dia = (
        rf_filtrado
        .groupby(['data', 'tipo'])['pecas']
        .sum()
        .reset_index()
    )
    pivot = (
        volume_dia
        .pivot_table(index='data', columns='tipo', values='pecas', aggfunc='sum')
        .fillna(0)
    )
    pivot.index    = pd.to_datetime(pivot.index)
    pivot          = pivot.sort_index()
    pivot.columns.name = None
    pivot = pivot.rename(columns={
        'realizado':  'Realizado',
        'reforecast': 'Reforecast',
        'previsao':   'Previsão',
    })
    st.bar_chart(pivot, use_container_width=True)

    st.subheader("Resumo por gerência e modelo de negócio")
    resumo = (
        rf_filtrado
        .groupby(['gerencia', 'modelo_negocio'])
        .agg(campanhas=('id_campanha', 'nunique'), pecas_total=('pecas', 'sum'))
        .reset_index()
        .rename(columns={
            'gerencia':       'Gerência',
            'modelo_negocio': 'Modelo',
            'campanhas':      'Campanhas',
            'pecas_total':    'Peças (total)',
        })
    )
    resumo['Peças (total)'] = resumo['Peças (total)'].round(0).astype(int)
    st.dataframe(resumo, use_container_width=True, hide_index=True)

# ─── ABA 2: ODP Vertical ──────────────────────────────────────────────────────
with aba2:
    if odp is None:
        st.info("Faça upload do arquivo ODP Vertical na barra lateral.")
    else:
        st.subheader("Entradas planejadas no CD — dia a dia (peças convertidas)")

        # Gráfico de entradas por dia
        entradas_dia = (
            odp
            .groupby('data_entrada')[['pecas', 'pecas_convertidas']]
            .sum()
            .reset_index()
            .rename(columns={
                'data_entrada':     'Data',
                'pecas':            'Peças brutas',
                'pecas_convertidas':'Peças convertidas',
            })
            .set_index('Data')
            .sort_index()
        )
        st.bar_chart(entradas_dia[['Peças convertidas']], use_container_width=True)

        # Entradas por categoria
        st.subheader("Entradas por categoria")
        por_cat = (
            odp
            .groupby('categoria')
            .agg(
                campanhas=('id_campanha', 'nunique'),
                pecas=('pecas', 'sum'),
                pecas_conv=('pecas_convertidas', 'sum'),
            )
            .reset_index()
            .rename(columns={
                'categoria':  'Categoria',
                'campanhas':  'Campanhas',
                'pecas':      'Peças brutas',
                'pecas_conv': 'Peças convertidas',
            })
            .sort_values('Peças brutas', ascending=False)
        )
        for col in ['Peças brutas', 'Peças convertidas']:
            por_cat[col] = por_cat[col].round(0).astype(int)
        st.dataframe(por_cat, use_container_width=True, hide_index=True)

        # Tabela detalhada
        st.subheader("Detalhe por campanha")
        odp_exib = odp.copy()
        odp_exib['data_entrada']     = odp_exib['data_entrada'].dt.strftime('%d/%m/%Y')
        odp_exib['pecas']            = odp_exib['pecas'].round(0).astype(int)
        odp_exib['pecas_convertidas']= odp_exib['pecas_convertidas'].round(0).astype(int)
        st.dataframe(
            odp_exib.rename(columns={
                'id_campanha':      'ID',
                'campanha':         'Campanha',
                'categoria':        'Categoria',
                'pecas':            'Peças',
                'pecas_convertidas':'Peças convertidas',
                'status_entrada':   'Status',
                'data_entrada':     'Data entrada CD',
            }),
            use_container_width=True,
            hide_index=True
        )

# ─── ABA 3: Calendário ────────────────────────────────────────────────────────
with aba3:
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
            'campanha':       'Campanha',
            'data_inicio':    'Início',
            'data_fim':       'Fim',
            'webdays':        'Dias web',
            'modelo_negocio': 'Modelo',
            'gerencia':       'Gerência',
            'cd':             'CD',
            'previsao_pecas': 'Previsão (pçs)',
            'estoque_total':  'Estoque total',
            'status':         'Status',
        }),
        use_container_width=True,
        hide_index=True
    )

# ─── ABA 4: Exportar ──────────────────────────────────────────────────────────
with aba4:
    st.subheader("Exportar resultados")
    st.caption("Os dados exportados refletem o estado atual da simulação. Nada é salvo no sistema.")

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        pivot.to_excel(writer, sheet_name='Reforecast dia a dia')
        resumo.to_excel(writer, sheet_name='Resumo gerência', index=False)
        if odp is not None:
            odp_exib.to_excel(writer, sheet_name='ODP Vertical', index=False)
        cal_exib[colunas_exib].to_excel(writer, sheet_name='Calendário', index=False)

    st.download_button(
        label="Baixar Excel completo",
        data=buf.getvalue(),
        file_name=f"simulador_sodp_{date.today().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    csv_bytes = rf_filtrado.to_csv(index=False).encode('utf-8')
    st.download_button(
        label="Baixar CSV — reforecast detalhado",
        data=csv_bytes,
        file_name=f"reforecast_{date.today().strftime('%Y%m%d')}.csv",
        mime="text/csv"
    )
