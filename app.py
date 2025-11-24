import streamlit as st
import fitz  # PyMuPDF
import re
import pandas as pd
from io import BytesIO
import time

# --- Configuração da Página (Deve ser o primeiro comando) ---
st.set_page_config(
    page_title="Consolidador de Extratos",
    page_icon="📊",
    layout="wide", # Usa a tela inteira, melhor para tabelas grandes
    initial_sidebar_state="expanded"
)

# --- Estilização CSS para limpar o visual ---
st.markdown("""
    <style>
    .main {
        background-color: #f8f9fa;
    }
    h1 {
        color: #2c3e50;
    }
    div.stButton > button {
        background-color: #4CAF50;
        color: white;
        border: none;
        padding: 10px 24px;
        font-size: 16px;
        border-radius: 8px;
        width: 100%;
    }
    div.stButton > button:hover {
        background-color: #45a049;
        color: white;
    }
    </style>
    """, unsafe_allow_html=True)

# --- FUNÇÕES DE EXTRAÇÃO (Lógica Mantida) ---

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

@st.cache_data(show_spinner=False) # Cache para não reprocessar o mesmo arquivo
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
            "Conta": conta,
            "Saldo Anterior": limpar_valor_caixa(saldo_ant.group(1) if saldo_ant else ""),
            "Rendimento": limpar_valor_caixa(rend.group(1) if rend else ""),
            "Saldo Atual": limpar_valor_caixa(saldo_atual.group(1) if saldo_atual else "")
        }
    except Exception as e:
        return {"Nome do Arquivo": nome_arquivo, "Conta": f"Erro: {str(e)}", "Saldo Atual": 0.0}

@st.cache_data(show_spinner=False)
def processar_padrao_2(arquivo_bytes, nome_arquivo):
    try:
        doc = fitz.open(stream=arquivo_bytes, filetype="pdf")
        texto_completo = "".join([pag.get_text() for pag in doc])
        doc.close()

        # Regex unificados
        def busca_dupla(padrao1, padrao2, txt):
            m = re.search(padrao1, txt)
            return m if m else re.search(padrao2, txt)

        match_conta = busca_dupla(r"Conta\n([\d\-]+)", r"Conta\s*,\s*\"([\d\-]+)", texto_completo)
        conta = match_conta.group(1).strip() if match_conta else "Não encontrado"

        if "NÃO HOUVE MOVIMENTO" in texto_completo:
            return {"Nome do Arquivo": nome_arquivo, "Conta": conta, "Saldo Anterior": 0.0, "Rendimento": 0.0, "Saldo Atual": 0.0}

        m_ant = busca_dupla(r"SALDO ANTERIOR\n([\d\.,]+)", r"SALDO ANTERIOR\s*,\s*\"([\d\.,]+)", texto_completo)
        m_rend = busca_dupla(r"RENDIMENTO LÍQUIDO\n([\d\.,]+)", r"RENDIMENTO LÍQUIDO\s*,\s*\"([\d\.,]+)", texto_completo)
        m_atual = busca_dupla(r"SALDO ATUAL =\n([\d\.,]+)", r"SALDO ATUAL =\s*,\s*\"([\d\.,]+)", texto_completo)

        return {
            "Nome do Arquivo": nome_arquivo,
            "Conta": conta,
            "Saldo Anterior": limpar_valor_geral(m_ant.group(1) if m_ant else ""),
            "Rendimento": limpar_valor_geral(m_rend.group(1) if m_rend else ""),
            "Saldo Atual": limpar_valor_geral(m_atual.group(1) if m_atual else "")
        }
    except Exception as e:
         return {"Nome do Arquivo": nome_arquivo, "Conta": f"Erro: {str(e)}", "Saldo Atual": 0.0}

# --- BARRA LATERAL (SIDEBAR) ---
with st.sidebar:
    st.image("https://cdn-icons-png.flaticon.com/512/2991/2991112.png", width=80)
    st.title("Menu de Opções")
    
    st.info("ℹ️ **Instruções:**\n\n1. Escolha o tipo de extrato.\n2. Arraste os arquivos PDF.\n3. Clique em Processar.\n4. Baixe o Excel final.")
    
    st.divider()
    
    tipo_extrato = st.selectbox(
        "Selecione o Modelo de Extrato:",
        ["Extrato CAIXA", "Extrato Padrão 2 (Investimentos)"],
        index=0
    )

# --- ÁREA PRINCIPAL ---

st.title("📊 Consolidador de Extratos Bancários")
st.markdown("Bem-vindo. Utilize esta ferramenta para transformar seus extratos PDF em planilhas Excel organizadas.")
st.divider()

# Upload
col_upload, col_info = st.columns([2, 1])

with col_upload:
    arquivos_carregados = st.file_uploader(
        "Carregue seus arquivos PDF aqui:", 
        type=["pdf"], 
        accept_multiple_files=True,
        help="Você pode selecionar múltiplos arquivos de uma vez."
    )

with col_info:
    if arquivos_carregados:
        st.success(f"📂 {len(arquivos_carregados)} arquivos identificados.")
        botao_processar = st.button("Iniciar Processamento", type="primary")
    else:
        st.warning("Aguardando arquivos...")
        botao_processar = False

# Lógica de Processamento
if botao_processar and arquivos_carregados:
    
    dados_consolidados = []
    barra = st.progress(0, text="Lendo arquivos...")
    
    start_time = time.time()

    for i, arquivo in enumerate(arquivos_carregados):
        bytes_arquivo = arquivo.read()
        
        if "CAIXA" in tipo_extrato:
            dados = processar_caixa(bytes_arquivo, arquivo.name)
        else:
            dados = processar_padrao_2(bytes_arquivo, arquivo.name)
        
        dados_consolidados.append(dados)
        barra.progress((i + 1) / len(arquivos_carregados), text=f"Processando {i+1}/{len(arquivos_carregados)}")

    time.sleep(0.5) # Pequena pausa para efeito visual de conclusão
    barra.empty() # Remove a barra de progresso
    
    # DataFrame
    df = pd.DataFrame(dados_consolidados)
    
    # --- DASHBOARD DE RESULTADOS ---
    st.divider()
    st.subheader("📈 Resultado da Análise")
    
    # Métricas (KPIs)
    total_saldo = df["Saldo Atual"].sum()
    total_rendimento = df["Rendimento"].sum()
    total_contas = df["Conta"].nunique()

    kpi1, kpi2, kpi3 = st.columns(3)
    kpi1.metric("Saldo Total Consolidado", f"R$ {total_saldo:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
    kpi2.metric("Rendimento Total", f"R$ {total_rendimento:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
    kpi3.metric("Contas Únicas", total_contas)

    # Abas para visualização
    aba_tabela, aba_grafico = st.tabs(["📄 Tabela de Dados", "📊 Gráfico de Saldos"])

    with aba_tabela:
        st.dataframe(
            df.style.format({
                "Saldo Anterior": "R$ {:,.2f}", 
                "Rendimento": "R$ {:,.2f}", 
                "Saldo Atual": "R$ {:,.2f}"
            }),
            use_container_width=True,
            height=400
        )

    with aba_grafico:
        if not df.empty:
            # Gráfico simples de barras
            st.bar_chart(df.set_index("Conta")["Saldo Atual"])
            st.caption("Visualização do Saldo Atual por número de conta.")

    # --- ÁREA DE DOWNLOAD ---
    st.divider()
    col_dl_1, col_dl_2, col_dl_3 = st.columns([1, 2, 1])
    
    with col_dl_2:
        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Extratos')
        
        st.download_button(
            label="📥 Baixar Planilha Excel Completa",
            data=buffer.getvalue(),
            file_name="Extratos_Consolidados.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
