import streamlit as st
import pandas as pd
import pdfplumber
import re
import math
import traceback
import plotly.express as px
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment
from openpyxl.utils import get_column_letter

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Portal Compras - Tapeçaria", layout="wide")

# --- INICIALIZAÇÃO DE ESTADO ---
if "uploader_key" not in st.session_state:
    st.session_state.uploader_key = 0

if "analise_concluida" not in st.session_state:
    st.session_state.analise_concluida = False

if "df_regras" not in st.session_state:
    st.session_state.df_regras = pd.DataFrame([
        {"FORNECEDOR": "CORTTEX", "MULTIPLO": 50, "TOLERANCIA": 20, "PALAVRA_CHAVE": ""},
        {"FORNECEDOR": "TEX COMPANY", "MULTIPLO": 50, "TOLERANCIA": 20, "PALAVRA_CHAVE": ""},
        {"FORNECEDOR": "CIPATEX", "MULTIPLO": 50, "TOLERANCIA": 20, "PALAVRA_CHAVE": ""},
        {"FORNECEDOR": "KARSTEN", "MULTIPLO": 50, "TOLERANCIA": 20, "PALAVRA_CHAVE": ""},
        {"FORNECEDOR": "ETRURIA", "MULTIPLO": 50, "TOLERANCIA": 20, "PALAVRA_CHAVE": ""},
        {"FORNECEDOR": "TELLAIO", "MULTIPLO": 50, "TOLERANCIA": 20, "PALAVRA_CHAVE": ""},
        {"FORNECEDOR": "OBER", "MULTIPLO": 50, "TOLERANCIA": 20, "PALAVRA_CHAVE": ""},
        {"FORNECEDOR": "TEXTIL J. SERRANO", "MULTIPLO": 50, "TOLERANCIA": 20, "PALAVRA_CHAVE": ""},
        {"FORNECEDOR": "CKS", "MULTIPLO": 50, "TOLERANCIA": 20, "PALAVRA_CHAVE": ""},
        {"FORNECEDOR": "AGRO QUIMICA", "MULTIPLO": 45, "TOLERANCIA": 20, "PALAVRA_CHAVE": ""},
        {"FORNECEDOR": "ROMPLAS", "MULTIPLO": 30, "TOLERANCIA": 15, "PALAVRA_CHAVE": "URUGUA"},
        {"FORNECEDOR": "ROMA DUBLADOS", "MULTIPLO": 10, "TOLERANCIA": 5, "PALAVRA_CHAVE": ""}
    ])

def limpar_dados():
    st.session_state.uploader_key += 1
    st.session_state.analise_concluida = False
    st.rerun()

# --- FUNÇÕES DE SUPORTE E LEITURA DE PDF ---
def limpar_v(val):
    if not val or val in ['-', 'S/N', 'N/A']:
        return 0.0
    val_limpo = str(val).replace('.', '').replace(',', '.')
    try:
        return float(val_limpo)
    except ValueError:
        return 0.0

def extrair_dados_pdf_web(file):
    dados = []
    meses_cabecalho = []
    nome_filial = file.name.replace(".pdf", "").upper()
    fornecedor_atual = "DESCONHECIDO"

    with pdfplumber.open(file) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue

            linhas = text.split('\n')
            for l in linhas:
                l_str = l.strip()

                if "FORNECEDOR:" in l_str.upper():
                    partes_forn = l_str.upper().split("FORNECEDOR:")
                    if len(partes_forn) > 1:
                        fornecedor_atual = partes_forn[1].split("-")[0].strip()

                if "CÓDIGO" in l_str.upper() and "DESCRIÇÃO" in l_str.upper():
                    partes_h = l_str.split()
                    cand_meses = [p for p in partes_h if len(p) == 3 and p.isalpha()]
                    if len(cand_meses) >= 4 and not meses_cabecalho:
                        meses_cabecalho = [m.upper() for m in cand_meses[:4]]

                l_limpa = re.sub(r'^\*+\s*', '', l_str)
                match_cod = re.search(r'\b\d{3,6}\b', l_limpa)

                if match_cod:
                    codigo = match_cod.group(0)
                    partes = l_limpa.split()

                    if len(partes) >= 12:
                        try:
                            item_dict = {
                                'CODIGO': codigo,
                                'DESCRICAO': " ".join([p for p in partes if not p.replace(',', '.').replace('-', '').replace('.', '').isdigit() and p != codigo])[:50],
                                'EMB.': partes[-12],
                                'MES_1': limpar_v(partes[-11]),
                                'MES_2': limpar_v(partes[-10]),
                                'MES_3': limpar_v(partes[-9]),
                                'MES_4': limpar_v(partes[-8]),
                                'MEDIA_SISTEMA': limpar_v(partes[-7]),
                                'ESTOQUE': limpar_v(partes[-6]),
                                'RESERVA': limpar_v(partes[-5]),
                                'COMPRADA': limpar_v(partes[-4]),
                                'SITUACAO': partes[-2],
                                'MESES_ESTOQUE': limpar_v(partes[-1]),
                                'FILIAL_NOME': nome_filial,
                                'FORNECEDOR': fornecedor_atual
                            }
                            dados.append(item_dict)
                        except Exception:
                            continue

    df_res = pd.DataFrame(dados)
    if not meses_cabecalho or len(meses_cabecalho) < 4:
        meses_cabecalho = ["MÊS 1", "MÊS 2", "MÊS 3", "MÊS 4"]

    return df_res, meses_cabecalho

# --- INTERFACE WEB (BARRA LATERAL) ---
with st.sidebar:
    try:
        st.image("logo.png", use_container_width=True)
    except Exception:
        pass

    st.markdown("---")
    st.header("📂 Nova Compra")
    uploaded_files = st.file_uploader(
        "Selecione os 4 PDFs das Unidades", 
        type="pdf", 
        accept_multiple_files=True,
        key=f"pdf_uploader_{st.session_state.uploader_key}"
    )
    
    st.button("🧹 Limpar Dados para Nova Compra", on_click=limpar_dados, use_container_width=True)
    st.markdown("---")

    with st.expander("⚙️ Configurações Avançadas"):
        meta = st.number_input("Meta de estoque (meses)", min_value=1, value=2)
        meses_parado = st.number_input("Considerar estoque parado após (meses)", min_value=1, value=3, step=1)
        fator_pico = st.number_input("Sensibilidade de Pico (x vezes a média)", min_value=1.5, value=2.5, step=0.5)
        nome_sugerido = st.text_input("Nome do ficheiro Excel", value="Relatorio_Compras_Tapecaria")
        nome_final_xlsx = nome_sugerido if nome_sugerido.endswith(".xlsx") else f"{nome_sugerido}.xlsx"

    with st.expander("🏭 Fornecedores e Múltiplos"):
        st.caption("Edite ou adicione regras na última linha vazia.")
        df_regras_editado = st.data_editor(st.session_state.df_regras, num_rows="dynamic", use_container_width=True, hide_index=True)
        st.session_state.df_regras = df_regras_editado

# --- CORPO DO SITE ---
col1, col2 = st.columns([1, 15])
with col1:
    try:
        st.image("simbolo.png", width=50)
    except Exception:
        pass
with col2:
    st.title("Inteligência de Compras")

st.markdown("##### Portal Operacional - Tapeçaria")
st.markdown("<br>", unsafe_allow_html=True)

# --- PROCESSAMENTO AUTOMÁTICO ---
if uploaded_files:
    with st.spinner("🔍 Lendo arquivos PDF e executando cálculo das regras de compras..."):
        dfs_por_filial = {}
        todos_dados = []
        meses_globais = []

        for f in uploaded_files:
            df, meses = extrair_dados_pdf_web(f)
            if not df.empty:
                dfs_por_filial[f.name.replace(".pdf", "").upper()] = df
                todos_dados.append(df)
                if len(meses) >= 4 and not meses_globais:
                    meses_globais = meses[:4]

        if not meses_globais:
            meses_globais = ["MÊS 1", "MÊS 2", "MÊS 3", "MÊS 4"]

        if not todos_dados:
            st.error("⚠️ O sistema não encontrou produtos compatíveis nos PDFs.")
        else:
            try:
                df_global = pd.concat(todos_dados).reset_index(drop=True)
                df_global['ESTOQUE_DISPONIVEL'] = df_global['ESTOQUE']

                vendas_recentes = df_global['MES_1'] + df_global['MES_2'] + df_global['MES_3'] + df_global['MES_4']
                df_global['TOTAL_VENDAS_RECENTES'] = vendas_recentes

                tracker_estoque = {}
                for _, row in df_global.iterrows():
                    f_nome = row['FILIAL_NOME']
                    c = row['CODIGO']
                    est = float(row['ESTOQUE'])
                    med = float(row['MEDIA_SISTEMA'])
                    excesso = est if med == 0 else max(0.0, est - (med * meta))
                    tracker_estoque[(f_nome, c)] = {'EXCEDENTE': excesso, 'MEDIA': med, 'ESTOQUE_FINAL': est}

                dash_qtd_comprar = 0
                dash_qtd_transferida = 0
                dash_itens_pico = 0
                dash_itens_ruptura = 0

                regras_dict = dict(zip(st.session_state.df_regras['FORNECEDOR'].str.upper(), st.session_state.df_regras['MULTIPLO']))

                for f_nome, df_f in dfs_por_filial.items():
                    suge_compra = []
                    trans_interna = []
                    ruptura_critica = []
                    venda_atipica = []
                    est_parado = []

                    for _, row in df_f.iterrows():
                        c = row['CODIGO']
                        med = float(row['MEDIA_SISTEMA'])
                        est = float(row['ESTOQUE'])
                        comp = float(row['COMPRADA'])
                        res = float(row['RESERVA'])
                        fornecedor = str(row['FORNECEDOR']).upper()

                        # Análise de Ruptura
                        if est == 0 and comp == 0 and med > 0:
                            ruptura_critica.append("🚨 CRÍTICA")
                            dash_itens_ruptura += 1
                        else:
                            ruptura_critica.append("OK")

                        # Análise de Pico
                        max_venda = max(row['MES_1'], row['MES_2'], row['MES_3'], row['MES_4'])
                        if max_venda > (med * fator_pico) and med > 0:
                            venda_atipica.append("⚠️ SIM")
                            dash_itens_pico += 1
                        else:
                            venda_atipica.append("NÃO")

                        # Análise de Parado
                        vendas_tot = row['MES_1'] + row['MES_2'] + row['MES_3'] + row['MES_4']
                        if vendas_tot == 0 and est > 0:
                            est_parado.append("🛑 SIM")
                        else:
                            est_parado.append("NÃO")

                        # Cálculo de Necessidade e Transferência
                        necessidade = max(0.0, (med * meta) - (est + comp - res))
                        qtd_transf = 0.0
                        origem = ""

                        if necessidade > 0:
                            for (outra_f, cod_item), d_est in tracker_estoque.items():
                                if cod_item == c and outra_f != f_nome and d_est['EXCEDENTE'] > 0:
                                    atend = min(necessidade, d_est['EXCEDENTE'])
                                    qtd_transf += atend
                                    d_est['EXCEDENTE'] -= atend
                                    necessidade -= atend
                                    origem = outra_f
                                    dash_qtd_transferida += atend
                                    if necessidade == 0:
                                        break

                        mult = regras_dict.get(fornecedor, 1)
                        qtd_compra = math.ceil(necessidade / mult) * mult if necessidade > 0 else 0
                        dash_qtd_comprar += qtd_compra

                        suge_compra.append(qtd_compra)
                        trans_interna.append(f"{qtd_transf:.0f} DE {origem}" if qtd_transf > 0 else "0")

                    df_f['SUGESTAO COMPRA'] = suge_compra
                    df_f['TRANS INTERNA'] = trans_interna
                    df_f['RUPTURA CRÍTICA'] = ruptura_critica
                    df_f['VENDA_ATIPICA'] = venda_atipica
                    df_f['ESTOQUE PARADO'] = est_parado
                    df_f['MEDIA'] = df_f['MEDIA_SISTEMA']

                df_p = df_global[(df_global['TOTAL_VENDAS_RECENTES'] == 0) & (df_global['ESTOQUE_DISPONIVEL'] > 0)].copy()

                st.session_state.dfs_por_filial = dfs_por_filial
                st.session_state.dash_qtd_comprar = dash_qtd_comprar
                st.session_state.dash_qtd_transferida = dash_qtd_transferida
                st.session_state.dash_itens_pico = dash_itens_pico
                st.session_state.dash_itens_ruptura = dash_itens_ruptura
                st.session_state.df_p = df_p
                st.session_state.analise_concluida = True

            except Exception as e:
                st.error(f"🚨 Ocorreu um erro durante os cálculos: {e}")
                st.code(traceback.format_exc())

# --- RENDERIZAÇÃO DAS ABAS ---
if st.session_state.analise_concluida:
    dfs_por_filial = st.session_state.dfs_por_filial
    dash_qtd_comprar = st.session_state.dash_qtd_comprar
    dash_qtd_transferida = st.session_state.dash_qtd_transferida
    dash_itens_pico = st.session_state.dash_itens_pico
    dash_itens_ruptura = st.session_state.dash_itens_ruptura
    df_p = st.session_state.df_p

    tab1, tab2, tab3, tab4 = st.tabs(["📊 Visão Geral", "🚨 Top Urgentes", "📦 Estoque Parado", "🔍 Prévia por Filial"])

    with tab1:
        st.subheader("Indicadores de Desempenho")
        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("🛒 Sugestão de Compra", f"{int(dash_qtd_comprar)} un.")
        c2.metric("🔄 Economia (Transf.)", f"{int(dash_qtd_transferida)} un.")
        c3.metric("⚠️ Picos de Vendas", f"{int(dash_itens_pico)} itens")
        c4.metric("🚨 Rupturas Críticas", f"{int(dash_itens_ruptura)} itens")

        if not df_p.empty:
            f_p = df_p.groupby('FILIAL_NOME')['ESTOQUE_DISPONIVEL'].sum().idxmax()
        else:
            f_p = "Nenhuma"
        c5.metric("📦 Maior Estoque Parado", f_p)

        st.success("✅ Processamento concluído com sucesso!")

    with tab2:
        df_all = pd.concat(dfs_por_filial.values())
        df_rupturas = df_all[df_all['RUPTURA CRÍTICA'] == "🚨 CRÍTICA"].sort_values(by='MEDIA', ascending=False)
        if not df_rupturas.empty:
            st.error("🚨 PRODUTOS EM RUPTURA CRÍTICA DETECTADOS (Estoque Zero + Sem Pedido em Andamento)")
            st.dataframe(df_rupturas[['CODIGO', 'DESCRICAO', 'FILIAL_NOME', 'MEDIA', 'SUGESTAO COMPRA', 'FORNECEDOR']], use_container_width=True)
        else:
            st.success("✅ Nenhuma ruptura crítica absoluta detectada nas filiais!")

        st.markdown("<br><hr>", unsafe_allow_html=True)
        st.subheader("🛒 Maior Volume de Compra Sugerido (Top 15)")
        top_compra = df_all[df_all['SUGESTAO COMPRA'] > 0].sort_values(by='SUGESTAO COMPRA', ascending=False).head(15)
        st.dataframe(top_compra[['CODIGO', 'DESCRICAO', 'FILIAL_NOME', 'SUGESTAO COMPRA', 'FORNECEDOR']], use_container_width=True)

    with tab3:
        st.subheader("Distribuição de Estoque Excedente / Sem Giro")
        if not df_p.empty:
            grafico_dados = df_p.groupby('FILIAL_NOME')['ESTOQUE_DISPONIVEL'].sum().reset_index()
            fig = px.bar(
                grafico_dados, 
                x='FILIAL_NOME', 
                y='ESTOQUE_DISPONIVEL', 
                title="Volume de Estoque Acima do Limite de Giro por Filial", 
                color='ESTOQUE_DISPONIVEL', 
                color_continuous_scale='Reds'
            )
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Nenhum estoque crítico/parado foi detectado com base nos parâmetros configurados.")

    with tab4:
        st.subheader("Prévia Colorida dos Dados por Filial")
        sel_f = st.selectbox("Selecione a Filial para visualizar:", list(dfs_por_filial.keys()))
        df_view = dfs_por_filial[sel_f].copy()

        def pintar_tabela(row):
            cols = row.index
            estilos = [''] * len(cols)
            def f_idx(nome): return cols.get_loc(nome) if nome in cols else -1

            i_parado = f_idx('ESTOQUE PARADO')
            i_estoque = f_idx('ESTOQUE')
            i_atipica = f_idx('VENDA_ATIPICA')
            i_compra = f_idx('SUGESTAO COMPRA')
            i_transf = f_idx('TRANS INTERNA')
            i_comprada = f_idx('COMPRADA')
            i_ruptura = f_idx('RUPTURA CRÍTICA')

            if i_parado >= 0 and '🛑 SIM' in str(row.get('ESTOQUE PARADO', '')):
                estilos[i_parado] = 'background-color: #F4CCCC; color: black;'
                if i_estoque >= 0: estilos[i_estoque] = 'background-color: #F4CCCC; color: black;'
            if i_atipica >= 0 and '⚠️ SIM' in str(row.get('VENDA_ATIPICA', '')): estilos[i_atipica] = 'background-color: #FFF2CC; color: black;'
            if i_compra >= 0 and pd.to_numeric(row.get('SUGESTAO COMPRA', 0), errors='coerce') > 0: estilos[i_compra] = 'background-color: #D9EAD3; color: black;'
            if i_transf >= 0 and str(row.get('TRANS INTERNA', '')) not in ['0', 'None', '', 'nan']: estilos[i_transf] = 'background-color: #C9DAF8; color: black;'
            if i_comprada >= 0 and pd.to_numeric(row.get('COMPRADA', 0), errors='coerce') > 0: estilos[i_comprada] = 'background-color: #FCE5CD; color: black;'
            if i_ruptura >= 0 and '🚨 CRÍTICA' in str(row.get('RUPTURA CRÍTICA', '')):
                estilos[i_ruptura] = 'background-color: #FFD2D2; color: black; font-weight: bold;'
                if i_estoque >= 0: estilos[i_estoque] = 'background-color: #FFD2D2; color: black;'
            return estilos

        st.dataframe(df_view.style.apply(pintar_tabela, axis=1), use_container_width=True)

else:
    st.info("Aguardando documentos. Por favor, selecione os ficheiros PDF na barra lateral para iniciar a análise.")
