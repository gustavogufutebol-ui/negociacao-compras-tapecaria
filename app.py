import streamlit as st
import pandas as pd
import pdfplumber
import re
import traceback
import plotly.express as px
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment
from openpyxl.utils import get_column_letter

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Portal de Negociação de Compras", layout="wide")

# --- TELA DE SENHA (OPCIONAL) ---
SENHA_ACESSO = "Tape2026"
if "liberado_negociacao" not in st.session_state:
    st.session_state.liberado_negociacao = False

if not st.session_state.liberado_negociacao:
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.markdown("<br><br><br><br>", unsafe_allow_html=True)
        try:
            st.image("logo.png", width=150)
        except Exception:
            st.markdown("<h1 style='text-align: center;'>🤝</h1>", unsafe_allow_html=True)
        
        st.markdown("<h2 style='text-align: center;'>Portal de Negociação</h2>", unsafe_allow_html=True)
        senha = st.text_input("Senha de Acesso", type="password")
        if st.button("Entrar", use_container_width=True):
            if senha == SENHA_ACESSO:
                st.session_state.liberado_negociacao = True
                st.rerun()
            else:
                st.error("Senha incorreta.")
    st.stop()

# --- FUNÇÕES DE SUPORTE ---
def limpar_v(val):
    """Converte valores numéricos no formato brasileiro (ex: '1.234,56' ou '0,00') para float."""
    if not val or val in ['-', 'S/N', 'N/A']:
        return 0.0
    val_limpo = str(val).replace('.', '').replace(',', '.')
    try:
        return float(val_limpo)
    except ValueError:
        return 0.0

def extrair_dados_proposta_pdf(file):
    """
    Lê os PDFs de propostas/tabelas de preços de fornecedores.
    """
    dados = []
    nome_arquivo = file.name.replace(".pdf", "").upper()
    fornecedor_atual = "DESCONHECIDO"

    with pdfplumber.open(file) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue

            linhas = text.split('\n')
            for l in linhas:
                l_str = l.strip()

                # Busca nome do Fornecedor no cabeçalho do PDF
                if "FORNECEDOR:" in l_str.upper():
                    partes_forn = l_str.upper().split("FORNECEDOR:")
                    if len(partes_forn) > 1:
                        fornecedor_atual = partes_forn[1].split("-")[0].strip()

                # Limpeza de caracteres ou asteriscos no código
                l_limpa = re.sub(r'^\*+\s*', '', l_str)
                match_cod = re.search(r'\b\d{3,6}\b', l_limpa)

                if match_cod:
                    codigo = match_cod.group(0)
                    partes = l_limpa.split()

                    # Verifica se contém preço atual e preço proposto
                    if len(partes) >= 4:
                        try:
                            preco_atual = limpar_v(partes[-2])
                            preco_novo = limpar_v(partes[-1])
                            
                            # Filtra apenas linhas válidas com valores de preço
                            if preco_atual > 0 or preco_novo > 0:
                                item_dict = {
                                    'FORNECEDOR': fornecedor_atual,
                                    'CODIGO': codigo,
                                    'DESCRICAO': " ".join([p for p in partes if not p.replace(',', '.').replace('-', '').replace('.', '').isdigit() and p != codigo])[:50],
                                    'PRECO_ATUAL': preco_atual,
                                    'PRECO_PROPOSTO': preco_novo,
                                    'ARQUIVO': nome_arquivo
                                }
                                dados.append(item_dict)
                        except Exception:
                            continue

    return pd.DataFrame(dados)

def pintar_tabela_negociacao(row):
    """Destaca aumentos e reduções na tabela interativa."""
    var = row.get('VARIACAO_%', 0)
    if var > 5.0:
        return ['background-color: #ffcccc; color: black'] * len(row) # Aumento Alto (Vermelho)
    elif var > 0:
        return ['background-color: #fff2cc; color: black'] * len(row) # Aumento Moderado (Amarelo)
    elif var < 0:
        return ['background-color: #d4edda; color: black'] * len(row) # Desconto / Redução (Verde)
    return [''] * len(row)

# --- BARRA LATERAL ---
with st.sidebar:
    try:
        st.image("logo.png", use_container_width=True)
    except Exception:
        pass

    st.markdown("---")
    st.header("📂 Propostas para Analisar")
    uploaded_files = st.file_uploader("Selecione os PDFs das Propostas", type="pdf", accept_multiple_files=True)
    st.markdown("---")

    with st.expander("⚙️ Regras de Negociação e Frete"):
        tipo_frete = st.selectbox("Modalidade de Frete", ["CIF (Incluso pelo Fornecedor)", "FOB (Adicionar Frete)"])
        percentual_frete = st.number_input("Adicional de Frete (%)", min_value=0.0, value=0.0, step=0.5) if "FOB" in tipo_frete else 0.0
        meta_reajuste = st.number_input("Limite Máximo Aceitável de Reajuste (%)", value=3.0, step=0.5)
        nome_sugerido = st.text_input("Nome do Arquivo Gerado", value="Relatorio_Negociacao_Precos")
        nome_final_xlsx = nome_sugerido if nome_sugerido.endswith(".xlsx") else f"{nome_sugerido}.xlsx"

# --- CORPO PRINCIPAL ---
col1, col2 = st.columns([1, 15])
with col1:
    try:
        st.image("simbolo.png", width=50)
    except Exception:
        pass
with col2:
    st.title("Portal de Negociação de Compras")

st.markdown("##### Tabulação de Propostas, Análise de Reajustes e Variação de Tabelas")
st.markdown("<br>", unsafe_allow_html=True)

if uploaded_files:
    with st.spinner("🔍 Lendo arquivos PDF e comparando tabelas de preços..."):
        todos_dados = []
        for f in uploaded_files:
            df_pdf = extrair_dados_proposta_pdf(f)
            if not df_pdf.empty:
                todos_dados.append(df_pdf)

        if not todos_dados:
            st.error("⚠️ Não foram encontrados itens com preços válidos nos PDFs enviados.")
        else:
            try:
                df_global = pd.concat(todos_dados).reset_index(drop=True)

                # Cálculo de frete e variação
                df_global['PRECO_FINAL_PROPOSTO'] = df_global['PRECO_PROPOSTO'] * (1 + (percentual_frete / 100))
                df_global['DIFERENCA_R$'] = df_global['PRECO_FINAL_PROPOSTO'] - df_global['PRECO_ATUAL']
                
                df_global['VARIACAO_%'] = df_global.apply(
                    lambda r: ((r['PRECO_FINAL_PROPOSTO'] - r['PRECO_ATUAL']) / r['PRECO_ATUAL'] * 100) if r['PRECO_ATUAL'] > 0 else 0.0,
                    axis=1
                )

                df_global['PARECER'] = df_global['VARIACAO_%'].apply(
                    lambda v: "🔴 REJEITAR (Acima do Limite)" if v > meta_reajuste else ("🟢 ACEITAR" if v <= 0 else "🟡 NEGOCIAR")
                )

                # --- DASHBOARD DE MÉTRICAS ---
                m1, m2, m3, m4 = st.columns(4)
                m1.metric("Itens Analisados", len(df_global))
                m2.metric("Variação Médial (%)", f"{df_global['VARIACAO_%'].mean():.2f}%")
                m3.metric("Itens Acima da Meta", len(df_global[df_global['VARIACAO_%'] > meta_reajuste]))
                m4.metric("Itens c/ Redução de Preço", len(df_global[df_global['VARIACAO_%'] < 0]))

                st.markdown("---")

                # --- ABAS DE ANÁLISE COM PLOTLY ---
                tab1, tab2 = st.tabs(["📊 Gráficos de Variação", "📋 Tabela Comparativa Completa"])

                with tab1:
                    st.subheader("Variação Média de Preços por Fornecedor (%)")
                    df_grafico = df_global.groupby("FORNECEDOR")["VARIACAO_%"].mean().reset_index()
                    
                    fig = px.bar(
                        df_grafico,
                        x="FORNECEDOR",
                        y="VARIACAO_%",
                        color="VARIACAO_%",
                        title="Variação Percentual Média da Proposta em Relação à Tabela Atual",
                        color_continuous_scale=["green", "yellow", "red"],
                        labels={"VARIACAO_%": "Variação Média (%)", "FORNECEDOR": "Fornecedor"}
                    )
                    st.plotly_chart(fig, use_container_width=True)

                with tab2:
                    st.subheader("Resumo Negocial por Item")
                    df_view = df_global[['FORNECEDOR', 'CODIGO', 'DESCRICAO', 'PRECO_ATUAL', 'PRECO_FINAL_PROPOSTO', 'DIFERENCA_R$', 'VARIACAO_%', 'PARECER']]
                    st.dataframe(df_view.style.apply(pintar_tabela_negociacao, axis=1), use_container_width=True)

                # --- GERADOR DE EXCEL (OPENPYXL) ---
                buffer = BytesIO()
                wb = Workbook()
                ws = wb.active
                ws.title = "Analise_Negociacao"

                header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
                header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
                border_thin = Border(left=Side(style='thin', color='D9D9D9'),
                                     right=Side(style='thin', color='D9D9D9'),
                                     top=Side(style='thin', color='D9D9D9'),
                                     bottom=Side(style='thin', color='D9D9D9'))

                colunas = list(df_global.columns)
                ws.append(colunas)
                for col_num, col_name in enumerate(colunas, 1):
                    cell = ws.cell(row=1, column=col_num)
                    cell.fill = header_fill
                    cell.font = header_font
                    cell.alignment = Alignment(horizontal="center", vertical="center")

                for row in df_global.itertuples(index=False):
                    ws.append(list(row))

                for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=len(colunas)):
                    for cell in row:
                        cell.border = border_thin
                        if isinstance(cell.value, (int, float)):
                            cell.number_format = '#,##0.00'

                for col in ws.columns:
                    max_len = max(len(str(cell.value or '')) for cell in col)
                    col_letter = get_column_letter(col[0].column)
                    ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

                wb.save(buffer)
                buffer.seek(0)

                st.markdown("---")
                st.download_button(
                    label="📥 Baixar Análise de Negociação em Excel",
                    data=buffer,
                    file_name=nome_final_xlsx,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            except Exception as e:
                st.error(f"Erro no processamento da negociação: {e}")
                st.text(traceback.format_exc())

else:
    st.info("Aguardando propostas. Faça o upload dos PDFs na barra lateral para calcular os reajustes de negociação.")
