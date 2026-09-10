import streamlit as st
import pandas as pd
import pdfplumber
import re
import math
import traceback
from io import BytesIO
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ==========================================
# --- CONFIGURAÇÃO DA PÁGINA ---
# ==========================================
st.set_page_config(page_title="Negociações de Compras", layout="wide")

# ==========================================
# --- VARIÁVEIS DE SESSÃO / REGRAS DE EMBALAGEM ---
# ==========================================
if "df_regras" not in st.session_state:
    st.session_state.df_regras = pd.DataFrame([
        {"FORNECEDOR": "YORK", "MULTIPLO": 50, "TOLERANCIA": 20, "PALAVRA_CHAVE": ""},
        {"FORNECEDOR": "CORTTEX", "MULTIPLO": 50, "TOLERANCIA": 20, "PALAVRA_CHAVE": ""},
        {"FORNECEDOR": "TEX COMPANY", "MULTIPLO": 50, "TOLERANCIA": 20, "PALAVRA_CHAVE": ""},
        {"FORNECEDOR": "CIPATEX", "MULTIPLO": 50, "TOLERANCIA": 20, "PALAVRA_CHAVE": ""},
        {"FORNECEDOR": "ROMPLAS", "MULTIPLO": 30, "TOLERANCIA": 15, "PALAVRA_CHAVE": "URUGUA"},
        {"FORNECEDOR": "ROMA DUBLADOS", "MULTIPLO": 10, "TOLERANCIA": 5, "PALAVRA_CHAVE": ""}
    ])

if "uploader_key" not in st.session_state:
    st.session_state.uploader_key = 0

# ==========================================
# --- FUNÇÃO DE LEITURA DE PDF ---
# ==========================================
def extrair_dados_pdf(file_bytes):
    """Extrai código, descrição, estoque, vendas e fornecedor dos PDFs de relatório."""
    dados = []
    try:
        with pdfplumber.open(file_bytes) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if not text:
                    continue
                for line in text.split('\n'):
                    match = re.search(r'(\d+)\s+(.*?)\s+(\d+[\.,]?\d*)\s+(\d+[\.,]?\d*)\s+(.*)', line)
                    if match:
                        cod, desc, est, media, forn = match.groups()
                        try:
                            est_val = float(est.replace('.', '').replace(',', '.'))
                            media_val = float(media.replace('.', '').replace(',', '.'))
                        except:
                            est_val, media_val = 0.0, 0.0
                            
                        dados.append({
                            'CODIGO': cod.strip(),
                            'DESCRICAO': desc.strip(),
                            'ESTOQUE': est_val,
                            'MEDIA_VENDAS': media_val,
                            'FORNECEDOR': forn.strip().upper()
                        })
    except Exception as e:
        st.error(f"Erro ao extrair PDF: {e}")
    return pd.DataFrame(dados)

# ==========================================
# --- FUNÇÃO DE ARREDONDAMENTO POR MÚLTIPLO ---
# ==========================================
def aplicar_multiplo(row, quantidade_sugerida):
    if quantidade_sugerida <= 0:
        return 0
    
    forn = str(row.get('FORNECEDOR', '')).upper()
    desc = str(row.get('DESCRICAO', '')).upper()
    
    mult, tol = 1, 0
    encontrou_regra = False
    
    for _, regra in st.session_state.df_regras.iterrows():
        f_regra = str(regra.get('FORNECEDOR', '')).upper()
        if f_regra and f_regra in forn:
            p_chave = str(regra.get('PALAVRA_CHAVE', '')).upper().strip()
            if p_chave and p_chave not in desc:
                continue
            try:
                mult = int(regra['MULTIPLO'])
                tol = int(regra['TOLERANCIA'])
                encontrou_regra = True
                break
            except:
                pass
                
    if encontrou_regra and mult > 0:
        base = (int(quantidade_sugerida) // mult) * mult
        resto = quantidade_sugerida % mult
        return int(base + mult) if resto >= tol else int(base)
    
    return int(math.ceil(quantidade_sugerida))

# ==========================================
# --- GERADOR DE EXCEL ---
# ==========================================
def gerar_excel(df_resultado, fornecedor_nome, vol_total):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_resultado.to_excel(writer, index=False, sheet_name="Sugestao_Negociacao")
        
        ws = writer.sheets["Sugestao_Negociacao"]
        header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        
        for col in range(1, len(df_resultado.columns) + 1):
            cell = ws.cell(row=1, column=col)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
            ws.column_dimensions[get_column_letter(col)].width = 24
            
    return output.getvalue()

# ==========================================
# --- BARRA LATERAL (PARÂMETROS DE NEGOCIAÇÃO) ---
# ==========================================
with st.sidebar:
    st.header("🤝 Parâmetros da Negociação")
    fornecedor_alvo = st.text_input("Fornecedor Selecionado", value="YORK")
    
    st.markdown("---")
    st.subheader("🎯 Filtro de Linha/Coleção")
    filtro_palavra = st.text_input("Filtrar descrição por (opcional)", value="", placeholder="Ex: URUGUAI, SUEDE, CORINO")
    
    st.markdown("---")
    st.subheader("📦 Volume Fechado")
    volume_total = st.number_input("Volume Total Acordado (Mts / Un)", min_value=1, value=30000, step=1000)
    
    st.subheader("📅 Cronograma de Entregas")
    c1, c2 = st.columns(2)
    with c1:
        rotulo_lote1 = st.text_input("Nome Lote 1", value="Setembro")
        pct_lote1 = st.number_input("% Lote 1", min_value=0, max_value=100, value=60)
    with c2:
        rotulo_lote2 = st.text_input("Nome Lote 2", value="Outubro")
        pct_lote2 = st.number_input("% Lote 2", min_value=0, max_value=100, value=40)
        
    st.markdown("---")
    uploaded_files = st.file_uploader(
        "Carregue os PDFs dos Produtos", 
        type="pdf", 
        accept_multiple_files=True, 
        key=f"pdf_uploader_{st.session_state.uploader_key}"
    )
    
    if st.button("🧹 Limpar Dados para Nova Negociação"):
        st.session_state.uploader_key += 1
        st.rerun()

    with st.expander("🏭 Tabela de Múltiplos por Fornecedor"):
        st.caption("Ajuste as embalagens mínimas e tolerâncias:")
        st.session_state.df_regras = st.data_editor(
            st.session_state.df_regras, num_rows="dynamic", use_container_width=True, hide_index=True
        )

# ==========================================
# --- CORPO PRINCIPAL DO PROGRAMA ---
# ==========================================
st.title("🎯 Gerenciador de Negociações e Compras em Lote")
st.caption(f"Distribuição inteligente de volume negociado para {fornecedor_alvo}")

if uploaded_files:
    dados_totais = []
    for f in uploaded_files:
        df_pdf = extrair_dados_pdf(f)
        if not df_pdf.empty:
            dados_totais.append(df_pdf)
            
    if dados_totais:
        df_base = pd.concat(dados_totais).reset_index(drop=True)
        
        # 1. Filtrar Fornecedor
        if fornecedor_alvo:
            df_forn = df_base[df_base['FORNECEDOR'].str.contains(fornecedor_alvo.upper(), na=False)].copy()
        else:
            df_forn = df_base.copy()
            
        # 2. Filtrar Por Palavra-Chave (Solução 3)
        if filtro_palavra.strip():
            df_forn = df_forn[df_forn['DESCRICAO'].str.contains(filtro_palavra.upper().strip(), na=False)].copy()
            
        if df_forn.empty:
            st.warning(f"Nenhum produto encontrado para o fornecedor '{fornecedor_alvo}' com o filtro '{filtro_palavra}'.")
        else:
            # Agrupar itens repetidos entre filiais
            df_agrupado = df_forn.groupby(['CODIGO', 'DESCRICAO', 'FORNECEDOR'], as_index=False).agg({
                'ESTOQUE': 'sum',
                'MEDIA_VENDAS': 'sum'
            })
            
            # Apenas itens com venda (Curva A)
            df_curva_a = df_agrupado[df_agrupado['MEDIA_VENDAS'] > 0].copy()
            df_curva_a['LABEL_PRODUTO'] = df_curva_a['CODIGO'] + " - " + df_curva_a['DESCRICAO']
            
            lista_opcoes = df_curva_a['LABEL_PRODUTO'].tolist()
            
            st.markdown("---")
            st.subheader("📌 Seleção de Produtos Participantes da Negociação (Solução 1)")
            st.caption("Marque apenas os itens negociados. O volume de 30.000m será distribuído exclusivamente entre eles.")
            
            # Controle de seleção (Marcar / Desmarcar Todos)
            col_b1, col_b2, _ = st.columns([1, 1, 4])
            if "skus_selecionados" not in st.session_state:
                st.session_state.skus_selecionados = lista_opcoes
                
            if col_b1.button("✅ Selecionar Todos"):
                st.session_state.skus_selecionados = lista_opcoes
                st.rerun()
            if col_b2.button("❌ Desmarcar Todos"):
                st.session_state.skus_selecionados = []
                st.rerun()

            produtos_escolhidos = st.multiselect(
                "Itens Incluídos no Lote:",
                options=lista_opcoes,
                default=[p for p in st.session_state.skus_selecionados if p in lista_opcoes]
            )
            
            df_filtrado_final = df_curva_a[df_curva_a['LABEL_PRODUTO'].isin(produtos_escolhidos)].copy()
            
            if df_filtrado_final.empty:
                st.info("⚠️ Nenhum produto selecionado para a negociação. Selecione pelo menos um item acima.")
            else:
                soma_vendas = df_filtrado_final['MEDIA_VENDAS'].sum()
                
                # --- RATEIO PROPORCIONAL RE-CALCULADO ---
                df_filtrado_final['PARTICIPACAO_%'] = (df_filtrado_final['MEDIA_VENDAS'] / soma_vendas)
                df_filtrado_final['QTD_BRUTA_SUGERIDA'] = df_filtrado_final['PARTICIPACAO_%'] * volume_total
                
                # Lotes
                df_filtrado_final['QTD_LOTE1_BRUTA'] = df_filtrado_final['QTD_BRUTA_SUGERIDA'] * (pct_lote1 / 100.0)
                df_filtrado_final['QTD_LOTE2_BRUTA'] = df_filtrado_final['QTD_BRUTA_SUGERIDA'] * (pct_lote2 / 100.0)
                
                # Múltiplos
                df_filtrado_final[f'SUGESTAO_{rotulo_lote1.upper()}'] = df_filtrado_final.apply(
                    lambda r: aplicar_multiplo(r, r['QTD_LOTE1_BRUTA']), axis=1
                )
                df_filtrado_final[f'SUGESTAO_{rotulo_lote2.upper()}'] = df_filtrado_final.apply(
                    lambda r: aplicar_multiplo(r, r['QTD_LOTE2_BRUTA']), axis=1
                )
                
                df_filtrado_final['TOTAL_SUGERIDO'] = (
                    df_filtrado_final[f'SUGESTAO_{rotulo_lote1.upper()}'] + df_filtrado_final[f'SUGESTAO_{rotulo_lote2.upper()}']
                )
                
                # --- DASHBOARD DE RESULTADOS ---
                st.markdown("---")
                st.markdown("### 📊 Resumo da Distribuição do Pedido")
                m1, m2, m3, m4 = st.columns(4)
                
                total_alocado = df_filtrado_final['TOTAL_SUGERIDO'].sum()
                m1.metric("Meta Negociada", f"{volume_total:,} un.")
                m2.metric("Total Alocado (Com Múltiplos)", f"{total_alocado:,} un.")
                m3.metric(f"Lote 1 ({rotulo_lote1})", f"{df_filtrado_final[f'SUGESTAO_{rotulo_lote1.upper()}'].sum():,} un.")
                m4.metric(f"Lote 2 ({rotulo_lote2})", f"{df_filtrado_final[f'SUGESTAO_{rotulo_lote2.upper()}'].sum():,} un.")
                
                st.subheader("📋 Pedido Fechado por Item")
                cols_exibicao = [
                    'CODIGO', 'DESCRICAO', 'ESTOQUE', 'MEDIA_VENDAS',
                    f'SUGESTAO_{rotulo_lote1.upper()}', f'SUGESTAO_{rotulo_lote2.upper()}', 'TOTAL_SUGERIDO'
                ]
                
                st.dataframe(
                    df_filtrado_final[cols_exibicao].sort_values(by='MEDIA_VENDAS', ascending=False), 
                    use_container_width=True
                )
                
                excel_bytes = gerar_excel(df_filtrado_final[cols_exibicao], fornecedor_alvo, volume_total)
                st.download_button(
                    label="📥 Exportar Negociação para Excel",
                    data=excel_bytes,
                    file_name=f"Negociacao_{fornecedor_alvo}_{volume_total}Mts.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
    else:
        st.error("Não foi possível ler dados válidos nos PDFs.")
else:
    st.info("👈 Utilize a barra lateral para definir os parâmetros e carregar os PDFs.")
