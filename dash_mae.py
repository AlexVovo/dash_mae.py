import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
from io import BytesIO
from fpdf import FPDF

st.set_page_config(page_title="Controle de Absenteísmo", layout="wide")

# ==========================
# 🔗 LINK DIRETO PARA O GOOGLE SHEETS (compartilhado como Leitor)
# ==========================
sheet_id = "1hz8m06SdFVMvrk2-rkvfeyCvtWjXaWUUxYEj99_1JSk"
gid = "774671515"
csv_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&gid={gid}"

@st.cache_data(ttl=300)
def carregar_dados():
    df = pd.read_csv(csv_url)

    # Converter datas no formato brasileiro (dia/mês/ano)
    df['Data Início'] = pd.to_datetime(df['Data Início'], errors='coerce', dayfirst=True)
    df['Data Fim'] = pd.to_datetime(df['Data Fim'], errors='coerce', dayfirst=True)

    # --- Cálculo correto: Dias de Atestado (calendário, inclusivo) ---
    dias_atestado = np.full(len(df), np.nan)
    mask = df['Data Início'].notna() & df['Data Fim'].notna()

    if mask.any():
        delta = (df.loc[mask, 'Data Fim'].values.astype('datetime64[D]') -
                 df.loc[mask, 'Data Início'].values.astype('datetime64[D]')).astype('timedelta64[D]').astype(int)
        dias_atestado[mask] = delta + 1

    mask_inicio_only = df['Data Início'].notna() & df['Data Fim'].isna()
    dias_atestado[mask_inicio_only] = 1

    # Evitar negativos (erros de digitação)
    dias_atestado = np.where(dias_atestado < 0, np.nan, dias_atestado)

    # --- Cálculo de Dias de Afastamento (úteis, inclusivo) ---
    dias_afastamento = np.full(len(df), np.nan)
    if mask.any():
        starts = df.loc[mask, 'Data Início'].values.astype('datetime64[D]')
        ends_plus1 = (df.loc[mask, 'Data Fim'].values.astype('datetime64[D]') + np.timedelta64(1, 'D'))
        dias_afastamento[mask] = np.busday_count(starts, ends_plus1)

    dias_afastamento[mask_inicio_only] = 1
    dias_afastamento = np.where(dias_afastamento < 0, np.nan, dias_afastamento)

    df['Dias de Atestado'] = pd.Series(dias_atestado).astype('Int64')
    df['Dias de Afastamento'] = pd.Series(dias_afastamento).astype('Int64')

    # Criar coluna para marcar erros de data
    df['Erro de Data'] = np.where(
        (df['Data Fim'] < df['Data Início']) & df['Data Fim'].notna(), True, False
    )

    return df

df = carregar_dados()

# ==========================
# 🚨 Verificação de erros de data
# ==========================
if df['Erro de Data'].any():
    st.error("⚠️ Existem linhas com **Data Fim menor que Data Início** — verifique a planilha.")
    st.dataframe(df[df['Erro de Data']][['Matrícula', 'Colaborador', 'Setor', 'Data Início', 'Data Fim']])

# ==========================
# 🎨 TÍTULO E DESCRIÇÃO
# ==========================
st.title("📊 Controle de Absenteísmo")
st.caption("Atualizado automaticamente a partir da planilha pública do Google Sheets.")

# ==========================
# 🧱 FILTROS
# ==========================
col1, col2, col3 = st.columns(3)
setores = ['Todos'] + sorted(df['Setor'].dropna().unique().tolist())
cids = ['Todos'] + sorted(df['CID'].dropna().unique().tolist())

setor = col1.selectbox("Filtrar por Setor", setores)
cid = col2.selectbox("Filtrar por CID", cids)
periodo = col3.date_input("Período (Data Início)", [])

df_filtrado = df.copy()
if setor != 'Todos':
    df_filtrado = df_filtrado[df_filtrado['Setor'] == setor]
if cid != 'Todos':
    df_filtrado = df_filtrado[df_filtrado['CID'] == cid]
if len(periodo) == 2:
    df_filtrado = df_filtrado[
        (df_filtrado['Data Início'] >= pd.Timestamp(periodo[0])) &
        (df_filtrado['Data Início'] <= pd.Timestamp(periodo[1]))
    ]

# ==========================
# 📈 INDICADORES
# ==========================
total_dias_atestado = int(df_filtrado['Dias de Atestado'].fillna(0).sum())
total_dias_afast = int(df_filtrado['Dias de Afastamento'].fillna(0).sum())
media_dias = round(df_filtrado['Dias de Atestado'].dropna().mean(), 1) if not df_filtrado['Dias de Atestado'].dropna().empty else 0
colabs = df_filtrado['Colaborador'].nunique()

st.subheader("📈 Indicadores Gerais")
k1, k2, k3, k4 = st.columns(4)
k1.metric("Total de Dias de Atestado (calendário)", total_dias_atestado)
k2.metric("Total de Dias de Afastamento (úteis)", total_dias_afast)
k3.metric("Média de Dias por Colaborador", media_dias)
k4.metric("Colaboradores com Atestado", colabs)

# ==========================
# 📊 GRÁFICOS
# ==========================
col4, col5 = st.columns(2)

if not df_filtrado.empty:
    # Gráfico: Total de Atestado por Setor
    graf1 = px.bar(
        df_filtrado.groupby('Setor', as_index=False)['Dias de Atestado'].sum(),
        x='Setor', y='Dias de Atestado',
        title="📍 Dias de Atestado por Setor (calendário)",
        text='Dias de Atestado'
    )
    graf1.update_traces(textposition='outside')
    col4.plotly_chart(graf1, use_container_width=True)

    # Gráfico: Total de Atestado por Colaborador
    graf2 = px.bar(
        df_filtrado.groupby('Colaborador', as_index=False)['Dias de Atestado'].sum(),
        x='Colaborador', y='Dias de Atestado',
        title="👤 Dias de Atestado por Colaborador",
        text='Dias de Atestado'
    )
    graf2.update_traces(textposition='outside')
    col5.plotly_chart(graf2, use_container_width=True)

    # Evolução Mensal
    df_filtrado['Mês'] = df_filtrado['Data Início'].dt.to_period('M').astype(str)
    evolucao = df_filtrado.groupby('Mês', as_index=False)['Dias de Atestado'].sum()
    graf3 = px.line(evolucao, x='Mês', y='Dias de Atestado', markers=True,
                    title="📅 Evolução Mensal dos Afastamentos (calendário)")
    st.plotly_chart(graf3, use_container_width=True)

    # Resumo por Setor (tabela)
    resumo_setor = df_filtrado.groupby('Setor', as_index=False)[['Dias de Atestado', 'Dias de Afastamento']].sum()
    st.markdown("### 🧮 Totais por Setor")
    st.dataframe(resumo_setor)

else:
    st.warning("Nenhum dado encontrado com os filtros selecionados.")

# ==========================
# 📋 TABELA DETALHADA
# ==========================
st.subheader("📋 Dados Detalhados")
cols_to_show = ['Matrícula', 'Colaborador', 'Setor', 'CID', 'Data Início', 'Data Fim', 'Dias de Atestado', 'Dias de Afastamento']
available_cols = [c for c in cols_to_show if c in df_filtrado.columns]
st.dataframe(df_filtrado[available_cols])

# ==========================
# 📤 EXPORTAÇÕES (Excel e PDF)
# ==========================
st.subheader("📦 Exportar Dados")

colA, colB = st.columns(2)

# ---- Excel ----
excel_buffer = BytesIO()
with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
    df_filtrado.to_excel(writer, index=False, sheet_name='Absenteismo')

colA.download_button(
    label="📤 Exportar para Excel",
    data=excel_buffer.getvalue(),
    file_name="Controle_Absenteismo.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

# ---- PDF ----
class PDF(FPDF):
    def header(self):
        try:
            self.image("logo.png", 10, 8, 25)
        except:
            pass
        self.set_font("Helvetica", "B", 14)
        self.cell(0, 10, "Controle de Absenteísmo", border=False, ln=True, align="C")
        self.ln(6)

    def footer(self):
        self.set_y(-15)
        self.set_font("Helvetica", "I", 8)
        self.cell(0, 10, f"Página {self.page_no()}", 0, 0, "C")

pdf = PDF()
pdf.add_page()
pdf.set_font("Helvetica", "", 10)

header_cols = ["Matrícula", "Colaborador", "Setor", "CID", "Início", "Fim", "Dias (cal)", "Dias (úteis)"]
col_widths = [25, 45, 30, 20, 20, 20, 22, 22]
pdf.set_font("Helvetica", "B", 10)
for h, w in zip(header_cols, col_widths):
    pdf.cell(w, 8, h, border=1, align='C')
pdf.ln()
pdf.set_font("Helvetica", "", 9)
for _, row in df_filtrado.iterrows():
    vals = [
        str(row.get('Matrícula', '')),
        str(row.get('Colaborador', '')),
        str(row.get('Setor', '')),
        str(row.get('CID', '')),
        row['Data Início'].strftime('%d/%m/%Y') if pd.notna(row.get('Data Início')) else '',
        row['Data Fim'].strftime('%d/%m/%Y') if pd.notna(row.get('Data Fim')) else '',
        str(row.get('Dias de Atestado', '')),
        str(row.get('Dias de Afastamento', ''))
    ]
    for v, w in zip(vals, col_widths):
        pdf.cell(w, 7, v, border=1)
    pdf.ln()

pdf_buffer = BytesIO(pdf.output(dest="S"))

colB.download_button(
    label="🧾 Gerar PDF (com logo)",
    data=pdf_buffer,
    file_name="Controle_Absenteismo.pdf",
    mime="application/pdf"
)

st.caption("💡 O Excel e o PDF incluem apenas os dados filtrados.")
