import streamlit as st
import pandas as pd
import pdfplumber
import re
import math
import traceback
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment
from openpyxl.utils import get_column_letter

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Portal Compras - Tapeçaria", layout="wide")

# --- INICIALIZAÇÃO DE ESTADO (MÚLTIPLOS E REGRAS) ---
if "df_regras" not in st.session_state:
    st.session_state.df_regras = pd.DataFrame([
        {"FORNECEDOR": "FORNECEDOR A", "MULTIPLO": 10},
        {"FORNECEDOR": "FORNECEDOR B", "MULTIPLO": 50},
        {"FORNECEDOR": "GERAL", "MULTIPLO": 1}
    ])

# --- FUNÇÕES DE SUPORTE E LEITURA DE PDF ---
def limpar_v(val):
    """Converte valores numéricos no formato brasileiro (ex: '1.234,56' ou '0,00') para float."""
    if not val or val == '-' or val == 'S/N':
        return 0.0
    val_limpo = str(val).replace('.', '').replace(',', '.')
    try:
        return float(val_limpo)
    except ValueError:
        return 0.0

def extrair_dados_pdf_web(file):
    """
    Lê o PDF de forma flexível e resiliente, tratando colunas dinâmicas,
    códigos com asteriscos (***) e a coluna adicional de SITUAÇÃO.
    """
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

                # Identificação de Fornecedor no Cabeçalho
                if "FORNECEDOR:" in l_str.upper():
                    partes_forn = l_str.upper().split("FORNECEDOR:")
                    if len(partes_forn) > 1:
                        fornecedor_atual = partes_forn[1].split("-")[0].strip()

                # Identificação dos Meses no Cabeçalho da Tabela
                if "CÓDIGO" in l_str.upper() and "DESCRIÇÃO" in l_str.upper():
                    partes_h = l_str.split()
                    cand_meses = [p for p in partes_h if len(p) == 3 and p.isalpha()]
                    if len(cand_meses) >= 4 and not meses_cabecalho:
                        meses_cabecalho = [m.upper() for m in cand_meses[:4]]

                # Limpeza de asteriscos no início do código (ex: ***16759)
                l_limpa = re.sub(r'^\*+\s*', '', l_str)

                # Busca por código numérico válido de produto (3 a 6 dígitos)
                match_cod = re.search(r'\b\d{3,6}\b', l_limpa)
                if match_cod:
                    codigo = match_cod.group(0)
                    partes = l_limpa.split()

                    # Garante que a linha tenha os elementos da tabela
                    if len(partes) >= 12:
                        try:
                            # Leitura dinâmica pegando as posições a partir do final
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

def pintar_tabela(val):
    """Aplica cores dinâmicas nas linhas do Dataframe no Streamlit."""
    status = val.get('STATUS', '')
    if 'RUPTURA' in str(status):
        return ['background-color: #ffcccc'] * len(val)
    elif 'TRANSFERIR' in str(status):
        return ['background-color: #e6f2ff'] * len(val)
    elif 'EXCESSO' in str(status):
        return ['background-color: #fff2cc'] * len(val)
    return [''] * len(val)

# --- INTERFACE WEB (BARRA LATERAL) ---
with st.sidebar:
    try:
        st.image("logo.png", use_container_width=True)
    except Exception:
        pass

    st.markdown("---")
    st.header("📂 Nova Compra")
    uploaded_files = st.file_uploader("Selecione os 4 PDFs das Unidades", type="pdf", accept_multiple_files=True)
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

# ====================================================
# === PROCESSAMENTO AUTOMÁTICO ("PISCOU, MUDOU") ===
# ====================================================
if uploaded_files:
    with st.spinner("🔍 Processando arquivos PDF e calculando regras de estoque..."):
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

                # Rastreador de excedentes para cálculo de transferências
                tracker_estoque = {}
                for _, row in df_global.iterrows():
                    f_nome = row['FILIAL_NOME']
                    c = row['CODIGO']
                    est = float(row['ESTOQUE'])
                    med = float(row['MEDIA_SISTEMA'])
                    excesso = est if med == 0 else max(0.0, est - (med * meta))
                    tracker_estoque[(f_nome, c)] = {'EXCEDENTE': excesso, 'MEDIA': med, 'ESTOQUE_FINAL': est}

                # Lógica de Sugestão de Compras e Transferências
                resultados = []
                dash_qtd_comprar = 0
                dash_qtd_transferida = 0
                dash_itens_pico = 0
                dash_itens_ruptura = 0

                # Mapeamento de múltiplos de fornecedor
                regras_dict = dict(zip(st.session_state.df_regras['FORNECEDOR'].str.upper(), st.session_state.df_regras['MULTIPLO']))
                multiplo_padrao = regras_dict.get('GERAL', 1)

                for _, row in df_global.iterrows():
                    f_nome = row['FILIAL_NOME']
                    c = row['CODIGO']
                    med = float(row['MEDIA_SISTEMA'])
                    est = float(row['ESTOQUE'])
                    res = float(row['RESERVA'])
                    comp = float(row['COMPRADA'])
                    fornecedor = str(row['FORNECEDOR']).upper()

                    # Identificação de Picos
                    max_venda = max(row['MES_1'], row['MES_2'], row['MES_3'], row['MES_4'])
                    eh_pico = max_venda > (med * fator_pico) and med > 0
                    if eh_pico:
                        dash_itens_pico += 1

                    # Necessidade bruta
                    necessidade = max(0.0, (med * meta) - (est + comp - res))

                    qtd_transf = 0.0
                    origem_transf = ""

                    # Tenta buscar transferência de outras filiais com excesso
                    if necessidade > 0:
                        dash_itens_ruptura += 1
                        for (outra_filial, cod_item), dados_est in tracker_estoque.items():
                            if cod_item == c and outra_filial != f_nome and dados_est['EXCEDENTE'] > 0:
                                qtd_atendida = min(necessidade, dados_est['EXCEDENTE'])
                                qtd_transf += qtd_atendida
                                dados_est['EXCEDENTE'] -= qtd_atendida
                                necessidade -= qtd_atendida
                                origem_transf = outra_filial
                                dash_qtd_transferida += qtd_atendida
                                if necessidade == 0:
                                    break

                    # Múltiplo de embalagem/fornecedor
                    mult = regras_dict.get(fornecedor, multiplo_padrao)
                    qtd_comprar = math.ceil(necessidade / mult) * mult if necessidade > 0 else 0
                    dash_qtd_comprar += qtd_comprar

                    # Status visual
                    if qtd_comprar > 0:
                        status = "🔴 RUPTURA / COMPRAR"
                    elif qtd_transf > 0:
                        status = f"🔵 TRANSFERIR DE {origem_transf}"
                    elif med == 0 and est > 0:
                        status = "🟡 EXCESSO / SEM GIRO"
                    else:
                        status = "🟢 OK"

                    resultados.append({
                        'FILIAL': f_nome,
                        'CODIGO': c,
                        'DESCRICAO': row['DESCRICAO'],
                        'FORNECEDOR': fornecedor,
                        'MEDIA': med,
                        'ESTOQUE': est,
                        'COMPRADA': comp,
                        'SUG_COMPRA': qtd_comprar,
                        'SUG_TRANSF': qtd_transf,
                        'ORIGEM_TRANSF': origem_transf,
                        'STATUS': status,
                        'MES_1': row['MES_1'],
                        'MES_2': row['MES_2'],
                        'MES_3': row['MES_3'],
                        'MES_4': row['MES_4']
                    })

                df_final = pd.DataFrame(resultados)

                # --- PAINEL DE MÉTRICAS (DASHBOARD) ---
                m1, m2, m3, m4 = st.columns(4)
                m1.metric("Unidades a Comprar", f"{dash_qtd_comprar:,.0f}".replace(",", "."))
                m2.metric("Unidades a Transferir", f"{dash_qtd_transferida:,.0f}".replace(",", "."))
                m3.metric("Itens em Ruptura", dash_itens_ruptura)
                m4.metric("Alertas de Pico", dash_itens_pico)

                st.markdown("---")

                # --- TABELA DE VISUALIZAÇÃO NO STREAMLIT ---
                df_view = df_final[['FILIAL', 'CODIGO', 'DESCRICAO', 'FORNECEDOR', 'MEDIA', 'ESTOQUE', 'COMPRADA', 'SUG_COMPRA', 'SUG_TRANSF', 'STATUS']]
                st.dataframe(df_view.style.apply(pintar_tabela, axis=1), use_container_width=True)

                # --- GERADOR DE EXCEL (OPENPYXL) ---
                buffer = BytesIO()
                wb = Workbook()
                ws = wb.active
                ws.title = "Plano_de_Compras"

                # Estilos do Excel
                header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
                header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
                border_thin = Border(left=Side(style='thin', color='D9D9D9'),
                                     right=Side(style='thin', color='D9D9D9'),
                                     top=Side(style='thin', color='D9D9D9'),
                                     bottom=Side(style='thin', color='D9D9D9'))

                # Cabeçalhos
                colunas = list(df_final.columns)
                ws.append(colunas)
                for col_num, col_name in enumerate(colunas, 1):
                    cell = ws.cell(row=1, column=col_num)
                    cell.fill = header_fill
                    cell.font = header_font
                    cell.alignment = Alignment(horizontal="center", vertical="center")

                # Dados
                for row in df_final.itertuples(index=False):
                    ws.append(list(row))

                # Formatação de bordas e alinhamentos
                for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=len(colunas)):
                    for cell in row:
                        cell.border = border_thin
                        if isinstance(cell.value, (int, float)):
                            cell.number_format = '#,##0.00'

                # Auto-ajuste de largura das colunas
                for col in ws.columns:
                    max_len = max(len(str(cell.value or '')) for cell in col)
                    col_letter = get_column_letter(col[0].column)
                    ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

                wb.save(buffer)
                buffer.seek(0)

                # Botão de Download
                st.download_button(
                    label="📥 Baixar Relatório em Excel",
                    data=buffer,
                    file_name=nome_final_xlsx,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            except Exception as e:
                st.error(f"Erro ao processar arquivos: {e}")
                st.text(traceback.format_exc())

else:
    st.info("A aguardar documentos. Por favor, carregue os ficheiros PDF na barra lateral para iniciar.")
