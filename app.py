import streamlit as st
import fitz  # PyMuPDF
import re
import pandas as pd
from io import BytesIO
import time

# --- Configuração da Página ---
st.set_page_config(
    page_title="Consolidador de Extratos",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- Estilização CSS ---
st.markdown("""
    <style>
    .main { background-color: #f8f9fa; }
    h1 { color: #2c3e50; }
    div.stButton > button {
        background-color: #003399;
        color: white;
        border: none;
        padding: 10px 24px;
        font-size: 16px;
        border-radius: 8px;
        width: 100%;
    }
    div.stButton > button:hover {
        background-color: #002266;
        color: white;
    }
    /* Ajuste para alinhar métricas */
    [data-testid="stMetricValue"] {
        font-size: 1.5rem;
    }
    </style>
    """, unsafe_allow_html=True)

# --- FUNÇÃO DE FORMATAÇÃO (BR) ---
def formatar_moeda_br(valor):
    """
    Recebe um float (ex: 1234.56) e retorna string BR (ex: R$ 1.234,56)
    """
    if valor is None:
        return "R$ 0,00"
    # Formata como padrão US primeiro (1,234.56)
    s = f"{valor:,.2f}"
    # Troca vírgula por X, ponto por vírgula, X por ponto
    return f"R$ {s.replace(',', 'X').replace('.', ',').replace('X', '.')}"

# --- FUNÇÕES AUXILIARES DE LIMPEZA ---
def limpar_valor_caixa(valor_str):
    if not valor_str: return 0.0
    valor_limpo = valor_str.strip().upper()
    valor_limpo = re.sub(r'[A-Z]', '', valor_limpo).strip()
    valor_limpo = valor_limpo.replace('.', '').replace(',', '.')
    try: return float(valor_limpo)
    except ValueError: return 0.0

def limpar_valor_geral(valor_str):
    if valor_str:
        valor_limpo = valor_str.replace('.', '').replace(',', '.')
        try: return float(valor_limpo)
        except ValueError: return 0.0
    return 0.0

# --- FUNÇÕES DE PROCESSAMENTO ---

@st.cache_data(show_spinner=False)
def processar_caixa(arquivo_bytes, nome_arquivo):
    try:
        doc = fitz.open(stream=arquivo_bytes, filetype="pdf")
        texto_completo = doc[0].get_text()
        doc.close()

        conta = re.search(r"Conta Corrente\s*\n\s*([\d\.\-]+)", texto_completo)
        conta = conta.group(1) if conta else "Não identificada"

        saldo_ant = re.search(r"Saldo Anterior\s*\n\s*([\d\.\,]+[CD]?)", texto_completo)
        rend = re.search(r"Rendimento Bruto no Mês\s*\n\s*([\d\.\,]+[CD]?)", texto_completo)
        saldo_atual = re.search(r"Saldo Bruto\*?\s*\n\s*([\d\.\,]+[CD]?)", texto_completo)

        return {
            "Nome do Arquivo": nome_arquivo,
            "Banco": "Caixa",
            "Conta": conta,
            "Saldo Anterior": limpar_valor_caixa(saldo_ant.group(1) if saldo_ant else ""),
            "Rendimento": limpar_valor_caixa(rend.group(1) if rend else ""),
            "Saldo Atual": limpar_valor_caixa(saldo_atual.group(1) if saldo_atual else "")
        }
    except Exception as e:
        return {"Nome do Arquivo": nome_arquivo, "Banco": "Caixa", "Conta": f"Erro: {str(e)}", "Saldo Atual": 0.0}

@st.cache_data(show_spinner=False)
def processar_bb(arquivo_bytes, nome_arquivo):
    try:
        doc = fitz.open(stream=arquivo_bytes, filetype="pdf")
        texto_completo = "".join([pag.get_text() for pag in doc])
        doc.close()

        def busca_dupla(padrao1, padrao2, txt):
            m = re.search(padrao1, txt)
            return m if m else re.search(padrao2, txt)

        match_conta = busca_dupla(r"Conta\n([\d\-]+)", r"Conta\s*,\s*\"([\d\-]+)", texto_completo)
        conta = match_conta.group(1).strip() if match_conta else "Não encontrado"

        if "NÃO HOUVE MOVIMENTO" in texto_completo:
            return {
                "Nome do Arquivo": nome_arquivo, 
                "Banco": "Banco do Brasil",
                "Conta": conta, "Saldo Anterior": 0.0, "Rendimento": 0.0, "Saldo Atual": 0.0
            }

        m_ant = busca_dupla(r"SALDO ANTERIOR\n([\d\.,]+)", r"SALDO ANTERIOR\s*,\s*\"([\d\.,]+)", texto_completo)
        m_rend = busca_dupla(r"RENDIMENTO LÍQUIDO\n([\d\.,]+)", r"RENDIMENTO LÍQUIDO\s*,\s*\"([\d\.,]+)", texto_completo)
        m_atual = busca_dupla(r"SALDO ATUAL =\n([\d\.,]+)", r"SALDO ATUAL =\s*,\s*\"([\d\.,]+)", texto_completo)

        return {
            "Nome do Arquivo": nome_arquivo,
            "Banco": "Banco do Brasil",
            "Conta": conta,
            "Saldo Anterior": limpar_valor_geral(m_ant.group(1) if m_ant else ""),
            "Rendimento": limpar_valor_geral(m_rend.group(1) if m_rend else ""),
            "Saldo Atual": limpar_valor_geral(m_atual.group(1) if m_atual else "")
        }
    except Exception as e:
         return {"Nome do Arquivo": nome_arquivo, "Banco": "BB", "Conta": f"Erro: {str(e)}", "Saldo Atual": 0.0}

# --- BARRA LATERAL ---
with st.sidebar:
    st.title("Menu")
    st.info("ℹ️ **Como usar:**\n\n1. Selecione o Banco.\n2. Arraste os arquivos PDF.\n3. Clique em Iniciar.\n4. Baixe o Excel.")
    st.divider()
    tipo_extrato = st.selectbox("Selecione o Banco:", ["Extrato CAIXA", "Extrato BB"], index=0)

# --- ÁREA PRINCIPAL ---
st.title("📊 Consolidador de Extratos")
st.markdown(f"Importação e análise de dados para: **{tipo_extrato}**")
st.divider()

col_upload, col_btn = st.columns([3, 1])
with col_upload:
    arquivos_carregados = st.file_uploader("Carregue os arquivos PDF:", type=["pdf"], accept_multiple_files=True)
with col_btn:
    st.write("") 
    st.write("") 
    if arquivos_carregados:
        botao_processar = st.button("▶️ Iniciar", type="primary")
    else:
        botao_processar = False

if botao_processar and arquivos_carregados:
    
    dados_consolidados = []
    barra = st.progress(0, text="Lendo arquivos...")
    
    for i, arquivo in enumerate(arquivos_carregados):
        bytes_arquivo = arquivo.read()
        if tipo_extrato == "Extrato CAIXA":
            dados = processar_caixa(bytes_arquivo, arquivo.name)
        else:
            dados = processar_bb(bytes_arquivo, arquivo.name)
        
        dados_consolidados.append(dados)
        barra.progress((i + 1) / len(arquivos_carregados), text=f"Processando {i+1}/{len(arquivos_carregados)}")

    time.sleep(0.5)
    barra.empty()
    
    df = pd.DataFrame(dados_consolidados)
    
    # --- RESULTADOS ---
    st.divider()
    st.subheader("📈 Resultado Consolidado")
    
    # Métricas (Usando a função formatar_moeda_br)
    total_saldo = df["Saldo Atual"].sum()
    total_rendimento = df["Rendimento"].sum()
    total_contas = df["Conta"].nunique()

    k1, k2, k3 = st.columns(3)
    k1.metric("Saldo Total", formatar_moeda_br(total_saldo))
    k2.metric("Rendimento Total", formatar_moeda_br(total_rendimento))
    k3.metric("Contas Únicas", total_contas)

    # Abas
    tab1, tab2 = st.tabs(["Tabela", "Gráfico"])

    with tab1:
        # Aplicando formatação visual na Tabela do Streamlit
        st.dataframe(
            df.style.format({
                "Saldo Anterior": formatar_moeda_br,
                "Rendimento": formatar_moeda_br,
                "Saldo Atual": formatar_moeda_br
            }),
            use_container_width=True,
            height=400
        )

    with tab2:
        if not df.empty:
            st.bar_chart(df.set_index("Conta")["Saldo Atual"])
            st.caption("Saldo Atual por Conta")

    # --- DOWNLOAD ---
    st.divider()
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Extratos')
    
    st.download_button(
        label="📥 Baixar Excel",
        data=buffer.getvalue(),
        file_name=f"Extratos_{tipo_extrato.replace(' ', '_')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
