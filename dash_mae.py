import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO
from fpdf import FPDF
from datetime import datetime

st.set_page_config(page_title="Controle de Absenteísmo - AESC", layout="wide")

# ==========================
# 🔗 LINK DO GOOGLE SHEETS
# ==========================
sheet_id = "1DgzNkglGSliXuAfgRp55YQmKrDXpk6Nu"
csv_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv"

@st.cache_data(ttl=300)
def carregar_dados():
    df = pd.read_csv(csv_url)
    df.columns = df.columns.str.strip().str.upper()

    # Renomear colunas conhecidas
    df.rename(columns={
        'ESTABELECIMENTOI': 'ESTABELECIMENTO',
        'FUNÇÃO': 'FUNCAO',
        'INÍCIO': 'INICIO',
        'TÉRMINO': 'TERMINO',
        'SUBGRUPO CID-10': 'CID10'
    }, inplace=True, errors='ignore')

    # Converter datas
    if 'INICIO' in df.columns:
        df['INICIO'] = pd.to_datetime(df['INICIO'], errors='coerce', dayfirst=True)
    if 'TERMINO' in df.columns:
        df['TERMINO'] = pd.to_datetime(df['TERMINO'], errors='coerce', dayfirst=True)

    # 🔹 Tratamento seguro da coluna DIAS
    if 'DIAS' in df.columns:
        # Converter para numérico, forçando erros para NaN (caso haja '#VALUE!')
        df['DIAS'] = pd.to_numeric(df['DIAS'], errors='coerce')
    else:
        # Calcular se a planilha não tiver DIAS
        mask = df['INICIO'].notna() & df['TERMINO'].notna()
        df.loc[mask, 'DIAS'] = (df.loc[mask, 'TERMINO'] - df.loc[mask, 'INICIO']).dt.days + 1

    # Substituir NaN por 1 e converter para inteiro
    df['DIAS'] = df['DIAS'].fillna(1).astype(int)

    # Criar coluna de MÊS (período)
    df['MES'] = df['INICIO'].dt.to_period('M').astype(str)

    return df

df = carregar_dados()

# ==========================
# 🎨 TÍTULO
# ==========================
st.title("📊 Controle de Absenteísmo - AESC")
st.caption("Dados atualizados automaticamente da planilha *AtestadosAesc2025*.")

# ==========================
# 🧱 FILTROS
# ==========================
col1, col2, col3 = st.columns(3)
estabs = ['Todos'] + sorted(df['ESTABELECIMENTO'].dropna().unique().tolist())
setores = ['Todos'] + sorted(df['SETOR'].dropna().unique().tolist())
funcoes = ['Todos'] + sorted(df['FUNCAO'].dropna().unique().tolist())
cids = ['Todos'] + sorted(df['CID10'].dropna().unique().tolist())

estab = col1.selectbox("🏥 Estabelecimento", estabs)
setor = col2.selectbox("🏢 Setor", setores)
funcao = col3.selectbox("👔 Função", funcoes)

col4, col5 = st.columns(2)
cid = col4.selectbox("🧬 CID10", cids)
periodo = col5.date_input("📆 Período (Início e Fim)", [])

df_filtrado = df.copy()
if estab != 'Todos':
    df_filtrado = df_filtrado[df_filtrado['ESTABELECIMENTO'] == estab]
if setor != 'Todos':
    df_filtrado = df_filtrado[df_filtrado['SETOR'] == setor]
if funcao != 'Todos':
    df_filtrado = df_filtrado[df_filtrado['FUNCAO'] == funcao]
if cid != 'Todos':
    df_filtrado = df_filtrado[df_filtrado['CID10'] == cid]
if len(periodo) == 2:
    df_filtrado = df_filtrado[
        (df_filtrado['INICIO'] >= pd.Timestamp(periodo[0])) &
        (df_filtrado['INICIO'] <= pd.Timestamp(periodo[1]))
    ]

# ==========================
# 📊 SELEÇÃO DE CORRELAÇÃO
# ==========================
st.markdown("### 🔍 Selecione a correlação que deseja visualizar:")
opcao = st.radio(
    "Escolha uma relação:",
    [
        "🏥 Estabelecimento × Setor × Dias",
        "🏥 Estabelecimento × Função × Dias",
        "🏥 Estabelecimento × Setor × CID",
        "🏥 Estabelecimento × Setor × CID × Dias",
        "📆 Estabelecimento × Setor × Dias × Mês"
    ]
)

# ==========================
# 📈 GRÁFICOS
# ==========================
if df_filtrado.empty:
    st.warning("Nenhum dado encontrado com os filtros selecionados.")
else:
    if opcao == "🏥 Estabelecimento × Setor × Dias":
        resumo = df_filtrado.groupby(['ESTABELECIMENTO', 'SETOR'], as_index=False)['DIAS'].sum()
        graf = px.bar(resumo, x='SETOR', y='DIAS', color='ESTABELECIMENTO',
                      title="🏥 Correlação: Estabelecimento × Setor × Dias", text='DIAS')
    elif opcao == "🏥 Estabelecimento × Função × Dias":
        resumo = df_filtrado.groupby(['ESTABELECIMENTO', 'FUNCAO'], as_index=False)['DIAS'].sum()
        graf = px.bar(resumo, x='FUNCAO', y='DIAS', color='ESTABELECIMENTO',
                      title="👔 Correlação: Estabelecimento × Função × Dias", text='DIAS')
    elif opcao == "🏥 Estabelecimento × Setor × CID":
        resumo = df_filtrado.groupby(['ESTABELECIMENTO', 'SETOR', 'CID10'], as_index=False).size()
        graf = px.bar(resumo, x='SETOR', y='size', color='CID10',
                      title="🧬 Correlação: Estabelecimento × Setor × CID", text='size')
    elif opcao == "🏥 Estabelecimento × Setor × CID × Dias":
        resumo = df_filtrado.groupby(['ESTABELECIMENTO', 'SETOR', 'CID10'], as_index=False)['DIAS'].sum()
        graf = px.bar(resumo, x='SETOR', y='DIAS', color='CID10',
                      title="🧩 Correlação: Estabelecimento × Setor × CID × Dias", text='DIAS')
    else:  # Correlação mensal corrigida
        resumo = df_filtrado.groupby(['ESTABELECIMENTO', 'SETOR', 'MES'], as_index=False)['DIAS'].sum()
        graf = px.bar(resumo, x='MES', y='DIAS', color='SETOR',
                      title="📆 Correlação: Estabelecimento × Setor × Dias × Mês", text='DIAS')

    graf.update_traces(texttemplate='%{text:.0f}', textposition='outside')
    graf.update_yaxes(title="Total de Dias", tickformat="d")
    graf.update_xaxes(title="Categoria")
    st.plotly_chart(graf, use_container_width=True)

    # ==========================
    # 📋 TABELA DETALHADA
    # ==========================
    st.markdown("### 📋 Dados Detalhados")
    df_filtrado['DIAS'] = df_filtrado['DIAS'].astype(int)
    st.dataframe(df_filtrado)

    # ==========================
    # 📦 EXPORTAR DADOS
    # ==========================
    st.markdown("### 📦 Exportar Dados")
    colA, colB = st.columns(2)

    # Excel
    excel_buffer = BytesIO()
    with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
        df_filtrado.to_excel(writer, index=False, sheet_name='AtestadosAesc2025')

    colA.download_button(
        "📤 Exportar para Excel",
        data=excel_buffer.getvalue(),
        file_name="AtestadosAesc2025.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # PDF PROFISSIONAL
    class PDF(FPDF):
        def header(self):
            try:
                self.image("logo.png", 10, 8, 25)
            except:
                pass
            self.set_font("Helvetica", "B", 15)
            self.cell(0, 10, "Relatório de Absenteísmo - AESC", ln=True, align="C")
            self.set_font("Helvetica", "", 10)
            self.cell(0, 8, f"Gerado em: {datetime.now().strftime('%d/%m/%Y %H:%M')}", ln=True, align="C")
            self.ln(4)
            self.set_draw_color(180, 180, 180)
            self.line(10, 28, 200, 28)
            self.ln(6)

        def footer(self):
            self.set_y(-15)
            self.set_font("Helvetica", "I", 9)
            self.set_text_color(120, 120, 120)
            self.cell(0, 10, f"Página {self.page_no()} / {{nb}}", align="C")

    pdf = PDF()
    pdf.alias_nb_pages()
    pdf.add_page()
    pdf.set_font("Helvetica", "", 10)

    header = ['ESTABELECIMENTO', 'SETOR', 'FUNCAO', 'CID10', 'INICIO', 'TERMINO', 'DIAS', 'MES']
    pdf.set_font("Helvetica", "B", 10)
    pdf.set_fill_color(230, 230, 230)
    for col in header:
        pdf.cell(25, 8, col[:12], border=1, align='C', fill=True)
    pdf.ln()

    pdf.set_font("Helvetica", "", 9)
    for _, row in df_filtrado[header].iterrows():
        for v in row:
            pdf.cell(25, 7, str(v)[:12], border=1)
        pdf.ln()

    pdf_bytes = pdf.output(dest="S").encode('latin1')
    pdf_buffer = BytesIO(pdf_bytes)

    colB.download_button(
        "🧾 Gerar PDF Profissional",
        data=pdf_buffer,
        file_name="Relatorio_Absenteismo_AESC.pdf",
        mime="application/pdf"
    )
