import streamlit as st
import fitz  # PyMuPDF
import re
import pandas as pd
from io import BytesIO

# --- Configuração da Página ---
st.set_page_config(page_title="Extrator de Extratos", page_icon="📄")

st.title("📄 Extrator de Extratos Bancários")
st.markdown("""
Faça o upload dos seus extratos em PDF para gerar uma tabela consolidada em Excel.
""")

# --- FUNÇÕES DE LIMPEZA E EXTRAÇÃO ---

def limpar_valor_caixa(valor_str):
    """Limpa valores no padrão da Caixa (ex: 1.000,00C)."""
    if not valor_str:
        return 0.0
    valor_limpo = valor_str.strip().upper()
    valor_limpo = re.sub(r'[A-Z]', '', valor_limpo).strip()
    valor_limpo = valor_limpo.replace('.', '').replace(',', '.')
    try:
        return float(valor_limpo)
    except ValueError:
        return 0.0

def limpar_valor_geral(valor_str):
    """Limpa valores no padrão comum (ex: 1.000,00)."""
    if valor_str:
        valor_limpo = valor_str.replace('.', '').replace(',', '.')
        try:
            return float(valor_limpo)
        except ValueError:
            return 0.0
    return 0.0

def processar_caixa(arquivo_bytes, nome_arquivo):
    """Lógica de extração específica para CAIXA."""
    try:
        # Abre o PDF a partir da memória (bytes)
        doc = fitz.open(stream=arquivo_bytes, filetype="pdf")
        pagina = doc[0]
        texto_completo = pagina.get_text()
        doc.close()

        # Regex Caixa
        regex_conta = r"Conta Corrente\s*\n\s*([\d\.\-]+)"
        regex_saldo_anterior = r"Saldo Anterior\s*\n\s*([\d\.\,]+[CD]?)"
        regex_rendimento = r"Rendimento Bruto no Mês\s*\n\s*([\d\.\,]+[CD]?)"
        regex_saldo_atual = r"Saldo Bruto\*?\s*\n\s*([\d\.\,]+[CD]?)"

        def buscar(padrao, texto):
            match = re.search(padrao, texto, re.MULTILINE)
            return match.group(1) if match else None

        conta = buscar(regex_conta, texto_completo)
        saldo_ant_raw = buscar(regex_saldo_anterior, texto_completo)
        rendimento_raw = buscar(regex_rendimento, texto_completo)
        saldo_atual_raw = buscar(regex_saldo_atual, texto_completo)

        if not conta: conta = "Não identificada"

        return {
            "Arquivo": nome_arquivo,
            "Conta": conta,
            "Saldo Anterior": limpar_valor_caixa(saldo_ant_raw),
            "Rendimento": limpar_valor_caixa(rendimento_raw),
            "Saldo Atual": limpar_valor_caixa(saldo_atual_raw)
        }
    except Exception as e:
        return {"Arquivo": nome_arquivo, "Conta": f"Erro: {str(e)}"}

def processar_padrao_2(arquivo_bytes, nome_arquivo):
    """Lógica de extração para o segundo modelo (Investimentos/Outros)."""
    try:
        doc = fitz.open(stream=arquivo_bytes, filetype="pdf")
        texto_completo = ""
        for pagina in doc:
            texto_completo += pagina.get_text()
        doc.close()

        # Regex Padrão 2
        regex_conta_novo = r"Conta\n([\d\-]+)"
        regex_conta_antigo = r"Conta\s*,\s*\"([\d\-]+)"
        
        regex_saldo_anterior_novo = r"SALDO ANTERIOR\n([\d\.,]+)"
        regex_saldo_anterior_antigo = r"SALDO ANTERIOR\s*,\s*\"([\d\.,]+)"

        regex_rend_liquido_novo = r"RENDIMENTO LÍQUIDO\n([\d\.,]+)"
        regex_rend_liquido_antigo = r"RENDIMENTO LÍQUIDO\s*,\s*\"([\d\.,]+)"

        regex_saldo_atual_novo = r"SALDO ATUAL =\n([\d\.,]+)"
        regex_saldo_atual_antigo = r"SALDO ATUAL =\s*,\s*\"([\d\.,]+)"

        def buscar_dado(padrao_novo, padrao_antigo, texto):
            match = re.search(padrao_novo, texto)
            if not match:
                match = re.search(padrao_antigo, texto)
            return match

        match_conta = buscar_dado(regex_conta_novo, regex_conta_antigo, texto_completo)
        conta = match_conta.group(1).strip() if match_conta else "Não encontrado"

        if "NÃO HOUVE MOVIMENTO NO PERÍODO SOLICITADO" in texto_completo:
            return {"Arquivo": nome_arquivo, "Conta": conta, "Saldo Anterior": 0.0, "Rendimento": 0.0, "Saldo Atual": 0.0}

        match_saldo_ant = buscar_dado(regex_saldo_anterior_novo, regex_saldo_anterior_antigo, texto_completo)
        match_rend = buscar_dado(regex_rend_liquido_novo, regex_rend_liquido_antigo, texto_completo)
        match_saldo_atual = buscar_dado(regex_saldo_atual_novo, regex_saldo_atual_antigo, texto_completo)

        return {
            "Arquivo": nome_arquivo,
            "Conta": conta,
            "Saldo Anterior": limpar_valor_geral(match_saldo_ant.group(1) if match_saldo_ant else ""),
            "Rendimento": limpar_valor_geral(match_rend.group(1) if match_rend else ""),
            "Saldo Atual": limpar_valor_geral(match_saldo_atual.group(1) if match_saldo_atual else "")
        }
    except Exception as e:
         return {"Arquivo": nome_arquivo, "Conta": f"Erro: {str(e)}"}

# --- INTERFACE PRINCIPAL ---

# Barra lateral para configuração
with st.sidebar:
    st.header("Configurações")
    tipo_extrato = st.radio(
        "Selecione o modelo do extrato:",
        ("Extrato CAIXA", "Extrato Padrão 2")
    )

# Área de Upload
uploaded_files = st.file_uploader(
    "Arraste os arquivos PDF aqui", 
    type=["pdf"], 
    accept_multiple_files=True
)

if uploaded_files:
    st.write(f"📂 {len(uploaded_files)} arquivos carregados.")
    
    if st.button("Processar Arquivos"):
        dados_consolidados = []
        barra_progresso = st.progress(0)
        
        for i, arquivo in enumerate(uploaded_files):
            # Lê os bytes do arquivo carregado
            bytes_arquivo = arquivo.read()
            
            # Decide qual função usar baseado na escolha da sidebar
            if tipo_extrato == "Extrato CAIXA":
                dados = processar_caixa(bytes_arquivo, arquivo.name)
            else:
                dados = processar_padrao_2(bytes_arquivo, arquivo.name)
            
            dados_consolidados.append(dados)
            # Atualiza barra de progresso
            barra_progresso.progress((i + 1) / len(uploaded_files))

        # Criação do DataFrame
        df = pd.DataFrame(dados_consolidados)
        
        st.success("Processamento concluído!")
        
        # Mostra a tabela na tela
        st.subheader("Resultados:")
        st.dataframe(df.style.format({
            "Saldo Anterior": "{:,.2f}", 
            "Rendimento": "{:,.2f}", 
            "Saldo Atual": "{:,.2f}"
        }))

        # --- BOTÃO DE DOWNLOAD EXCEL ---
        # Converter DF para Excel em memória (Buffer)
        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Dados')
        
        st.download_button(
            label="📥 Baixar Excel Consolidado",
            data=buffer.getvalue(),
            file_name="extratos_processados.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
