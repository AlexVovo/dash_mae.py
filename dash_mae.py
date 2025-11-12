import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO
from fpdf import FPDF
from datetime import datetime
import os

st.set_page_config(page_title="Controle de Absenteísmo - AESC", layout="wide")

# ==========================
# 📂 LEITURA LOCAL DO CSV
# ==========================
@st.cache_data(ttl=300)
def carregar_dados():
    # 🔗 Link direto para a aba correta do Google Sheets
    sheet_id = "1DgzNkglGSliXuAfgRp55YQmKrDXpk6Nu"
    csv_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&gid=1955342039"

    # Leitura online do CSV
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

    # Tratar a coluna DIAS
    if 'DIAS' in df.columns:
        df['DIAS'] = pd.to_numeric(df['DIAS'], errors='coerce')
    else:
        mask = df['INICIO'].notna() & df['TERMINO'].notna()
        df.loc[mask, 'DIAS'] = (df.loc[mask, 'TERMINO'] - df.loc[mask, 'INICIO']).dt.days + 1

    df['DIAS'] = df['DIAS'].fillna(1).astype(int)
    df['MES'] = df['INICIO'].dt.to_period('M').astype(str)

    return df



df = carregar_dados()

# ==========================
# 🎨 TÍTULO
# ==========================
st.title("📊 Controle de Absenteísmo - AESC")
st.caption("Dados atualizados automaticamente do arquivo local *AtestadosAesc2025.csv*.")

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
    else:
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

    # ==========================
    # 🧾 GERAR PDF PROFISSIONAL
    # ==========================
    class PDF(FPDF):
        def header(self):
            try:
                self.image("logo.png", 10, 6, 30)
            except:
                pass
            self.set_font("Helvetica", "B", 16)
            self.cell(0, 10, "Relatório de Absenteísmo - AESC", ln=True, align="C")
            self.set_font("Helvetica", "", 11)
            self.cell(0, 8, f"Gerado em: {datetime.now().strftime('%d/%m/%Y %H:%M')}", ln=True, align="C")
            self.ln(4)
            self.set_draw_color(150, 150, 150)
            self.line(10, 25, 287, 25)
            self.ln(8)

        def footer(self):
            self.set_y(-15)
            self.set_font("Helvetica", "I", 9)
            self.set_text_color(120, 120, 120)
            self.cell(0, 10, f"Página {self.page_no()} / {{nb}}", align="C")

    pdf = PDF(orientation='L', unit='mm', format='A4')
    pdf.alias_nb_pages()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.set_font("Helvetica", "", 10)

    header = ['COLABORADOR', 'ESTABELECIMENTO', 'SETOR', 'FUNCAO', 'INICIO', 'TERMINO', 'DIAS', 'CID10', 'MES']
    col_widths = [25, 45, 45, 40, 25, 25, 10, 25, 20]

    pdf.set_font("Helvetica", "B", 10)
    pdf.set_fill_color(220, 220, 220)
    for i, col in enumerate(header):
        pdf.cell(col_widths[i], 8, col, border=1, align='C', fill=True)
    pdf.ln()

    def fmt(v):
        if pd.isna(v):
            return ""
        if isinstance(v, (int, float)):
            return str(int(v))
        if isinstance(v, (pd.Timestamp, datetime)):
            return v.strftime("%d/%m/%Y")
        try:
            tmp = pd.to_datetime(v, errors='coerce')
            if not pd.isna(tmp):
                return tmp.strftime("%d/%m/%Y")
        except:
            pass
        return str(v)

    pdf.set_font("Helvetica", "", 9)
    line_height = 4
    fill = False

    for _, row in df_filtrado[header].iterrows():
        texts = [fmt(row[h]) for h in header]
        max_lines = max(1, max(int((pdf.get_string_width(txt) / max(col_widths[i]-2,1)) + 0.9999) for i, txt in enumerate(texts)))
        row_h = line_height * max_lines
        x_start, y_start = pdf.get_x(), pdf.get_y()
        for i, txt in enumerate(texts):
            pdf.set_xy(x_start + sum(col_widths[:i]), y_start)
            align = 'C' if header[i] == 'DIAS' else 'L'
            pdf.multi_cell(col_widths[i], line_height, txt, border=1, align=align, fill=fill)
        pdf.set_xy(x_start, y_start + row_h)
        fill = not fill
        if pdf.get_y() > pdf.h - 25:
            pdf.add_page()
            pdf.set_font("Helvetica", "B", 10)
            pdf.set_fill_color(220, 220, 220)
            for i, col in enumerate(header):
                pdf.cell(col_widths[i], 8, col, border=1, align='C', fill=True)
            pdf.ln()
            pdf.set_font("Helvetica", "", 9)

    pdf_output = pdf.output(dest="S")
    if isinstance(pdf_output, str):
        pdf_output = pdf_output.encode("latin1")
    pdf_buffer = BytesIO(pdf_output)

    colB.download_button(
        "🧾 Gerar PDF Profissional",
        data=pdf_buffer,
        file_name=f"Relatorio_Absenteismo_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf",
        mime="application/pdf"
    )
