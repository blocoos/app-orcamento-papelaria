import streamlit as st
import pandas as pd
import openpyxl
import os
from datetime import datetime
import unicodedata
import json
from io import BytesIO  # Essencial para ler/escrever arquivos em memória

# --- NOVAS BIBLIOTECAS PARA O PDF (Orçamento) ---
import base64
from xhtml2pdf import pisa
import sys
import subprocess
# --- FIM DAS NOVAS BIBLIOTECAS ---

# --- BIBLIOTECAS PARA GOOGLE DRIVE / PLANILHAS ---
import gspread
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
# --- FIM GOOGLE DRIVE ---

# --- NOVAS BIBLIOTECAS PARA O PDF (Vale / Livro) ---
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.platypus import Frame, FrameBreak
from reportlab.platypus.flowables import PageBreak
from reportlab.platypus.doctemplate import PageTemplate
from reportlab.platypus.flowables import Image as RLImage
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.units import cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
# --- FIM DAS NOVAS BIBLIOTECAS ---

# --- NOVAS BIBLIOTECAS PARA O GOOGLE ---
import gspread
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload, MediaIoBaseDownload
# --- FIM DAS NOVAS BIBLIOTECAS ---


# --- CAMINHOS (AGORA SÃO RELATIVOS OU CONSTANTES) ---
# CAMINHO_BASE não é mais necessário
ARQUIVO_BASE = "Planilha Salva Sono_1.xlsx" 
LOGO_PATH = "Logo.png" # Lê o logo do repositório GitHub
HEADER_DA_PLANILHA = 14 
# Caminhos de salvar PDF/Rascunho não são mais usados.
# --- FIM DA MUDANÇA ---


# --- CONFIGURAÇÕES DO GOOGLE DRIVE (COM SEUS IDs) ---
PLANILHA_MESTRE_ID = "1P2aJCePtRVaqx9pnw2t_L1vwiKeIFoP-FkLAGXXb7rY" 
PASTA_DRIVE_PRINCIPAL_ID = "1tzmGPT9mvsfCrBW8vP81f-1kQ6J9iv1b"
PASTA_DRIVE_ARQUIVO_MORTO_ID = "1tiKYOxH5reHkTnvSboY5TioNl-Uk-Gbk"
# --- MUDANÇA (SEU ID FOI ADICIONADO) ---
PASTA_DRIVE_PLANILHAS_ID = "1iCofZQcZKSSMqEef7xGdghBHlSJYq2Sh"
# --- FIM DA MUDANÇA ---
SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
# CLIENT_SECRET_FILE não é mais um arquivo, usaremos st.secrets
# --- FIM DAS CONFIGURAÇÕES GOOGLE ---

COR_PRINCIPAL_PAPELARIA = "#FF6600" # Laranja Forte
# --- FIM DA COR ---

# --- (NOVO) REGISTRO DA FONTE ARIAL (LENDO DO REPOSITÓRIO) ---
try:
    caminho_arial_normal = "arial.ttf"
    caminho_arial_bold = "arialbd.ttf"
    
    pdfmetrics.registerFont(TTFont('Arial', caminho_arial_normal))
    pdfmetrics.registerFont(TTFont('Arial-Bold', caminho_arial_bold))
    
    pdfmetrics.registerFontFamily('Arial', normal='Arial', bold='Arial-Bold', italic=None, boldItalic=None)
    
    styles = getSampleStyleSheet()
    styles['Normal'].fontName = 'Arial'
    styles['Heading3'].fontName = 'Arial-Bold'
    
except Exception as e:
    # Em nuvem, se a fonte falhar, o app quebra.
    st.error(f"Erro FATAL: Não foi possível carregar as fontes 'arial.ttf' ou 'arialbd.ttf'. Verifique se elas estão no repositório GitHub. Erro: {e}")
    st.stop()
# --- FIM DO REGISTRO DA FONTE ---


# --- FUNÇÕES DE ARQUIVO (MODIFICADAS PARA NUVEM) ---

# ABRE O ARQUIVO (NÃO É MAIS USADO, MAS MANTIDO POR PRECAUÇÃO)
def abrir_arquivo(caminho):
    st.warning("Função 'abrir_arquivo' não é suportada na nuvem. Use o botão de download.")

# --- NOVA FUNÇÃO HELPER (DOWNLOAD DO DRIVE) ---
@st.cache_data(ttl=600) # Cache de 10 minutos
def download_excel_bytes(_drive_service, folder_id, file_name):
    """Encontra um arquivo Excel no Drive pelo nome e o baixa para a memória (BytesIO)."""
    try:
        # 1. Encontrar o arquivo
        query = f"name='{file_name}' and '{folder_id}' in parents and trashed=false"
        response = _drive_service.files().list(q=query, spaces='drive', fields='files(id, name)').execute()
        files = response.get('files', [])
        
        file_id = files[0].get('id') if files else None
        
        if not file_id:
            st.error(f"Erro Crítico: Arquivo '{file_name}' não encontrado no Google Drive (Pasta ID: {folder_id}).")
            return None
            
        # 2. Baixar o arquivo para a memória
        request = _drive_service.files().get_media(fileId=file_id)
        file_buffer = BytesIO()
        downloader = MediaIoBaseDownload(file_buffer, request)
        
        done = False
        while done is False:
            status, done = downloader.next_chunk()
            
        file_buffer.seek(0)
        return file_buffer
        
    except Exception as e:
        st.error(f"Erro ao baixar o arquivo '{file_name}' do Drive: {e}")
        st.exception(e)
        return None

# --- FUNÇÃO ATUALIZADA (LÊ DO DRIVE) ---
@st.cache_data(ttl=3600) # Cache de 1 hora
def extrair_data_validade(_drive_service, folder_id):
    try:
        file_buffer = download_excel_bytes(_drive_service, folder_id, ARQUIVO_BASE)
        if file_buffer is None:
            return "Erro ao ler base"
            
        wb = openpyxl.load_workbook(file_buffer, data_only=True)
        ws = wb["Base"] 
        data = ws["F5"].value
        if isinstance(data, datetime):
            return data.strftime("%d/%m/%Y")
        return str(data) if data else "Não definida"
    except Exception as e:
        st.error(f"Erro ao ler data de validade do Drive: {e}"); return "Erro"

# --- FUNÇÃO ATUALIZADA (LÊ DO DRIVE) ---
@st.cache_data(ttl=600) # Cache de 10 minutos
def extrair_observacoes_iniciais(_drive_service, folder_id, file_name, aba):
    nao_trabalhamos = []
    para_escolher = []
    try:
        file_buffer = download_excel_bytes(_drive_service, folder_id, file_name)
        if file_buffer is None:
            return "", ""
            
        wb = openpyxl.load_workbook(file_buffer, data_only=True)
        ws = wb[aba]
        for col in range(3, 7):
            val_nt = ws.cell(row=12, column=col).value
            val_pe = ws.cell(row=13, column=col).value
            if val_nt: nao_trabalhamos.append(str(val_nt).strip())
            if val_pe: para_escolher.append(str(val_pe).strip())
    except Exception as e:
        st.warning(f"Erro ao extrair observações iniciais de '{file_name}' (Aba: {aba}): {e}")
    return "\n".join(nao_trabalhamos), "\n".join(para_escolher)

def normalizar_texto(texto):
    if not isinstance(texto, str): return ""
    texto = texto.strip().upper()
    texto = unicodedata.normalize("NFD", texto)
    texto = "".join(ch for ch in texto if unicodedata.category(ch) != "Mn")
    return texto

def sanitizar_nome_arquivo(nome):
    invalidos = '\\/*?:"<>|'
    for char in invalidos:
        nome = nome.replace(char, '')
    return nome.strip()

def formatar_telefone(tel_str):
    numeros = "".join(filter(str.isdigit, tel_str))
    if len(numeros) == 11:
        return f"({numeros[:2]}) {numeros[2:7]}.{numeros[7:]}"
    elif len(numeros) == 10:
        return f"({numeros[:2]}) {numeros[2:6]}.{numeros[6:]}"
    return tel_str

# --- FUNÇÃO ATUALIZADA (LÊ DO DRIVE) ---
@st.cache_data(ttl=600) # Cache de 10 minutos
def carregar_itens(_drive_service, folder_id, file_name, aba):
    try:
        file_buffer = download_excel_bytes(_drive_service, folder_id, file_name)
        if file_buffer is None:
            st.error(f"Não foi possível carregar o arquivo '{file_name}' do Drive.")
            return pd.DataFrame()
            
        df = pd.read_excel(file_buffer, sheet_name=aba, header=HEADER_DA_PLANILHA, dtype=str, engine='openpyxl')
    except Exception as e:
        st.error(f"Erro ao ler {aba} de '{file_name}': {e}"); return pd.DataFrame()
        
    df.columns = [normalizar_texto(col) for col in df.columns.astype(str)]
    if not any("COD" in c for c in df.columns) or not any("TIPO" in c for c in df.columns):
        st.error(f"❌ Coluna 'COD' ou 'TIPO' não encontrada na aba '{aba}' do arquivo '{file_name}'."); return pd.DataFrame()
    cod_col = [c for c in df.columns if "COD" in c][0]
    qtd_col_nome = [c for c in df.columns if "QTD" in c or "QTDE" in c or "QUANT" in c]
    col_tipo = [c for c in df.columns if "TIPO" in c][0]
    colunas_essenciais = [col_tipo, cod_col]
    qtd_col = qtd_col_nome[0] if qtd_col_nome else "QTD"
    if qtd_col_nome: colunas_essenciais.append(qtd_col)
    col_desc = [c for c in df.columns if "DESC" in c] 
    if col_desc: colunas_essenciais.append(col_desc[0])
    outras_cols = [c for c in df.columns if c not in colunas_essenciais]
    df = df[colunas_essenciais + outras_cols]
    df = df.rename(columns={cod_col: "COD", qtd_col: "QTD"})
    df["COD"] = df["COD"].astype(str).str.strip() 
    df[col_tipo] = df[col_tipo].astype(str).apply(normalizar_texto)
    if "QTD" not in df.columns: df["QTD"] = 1
    df = df[df["COD"].str.isdigit().fillna(False)] 
    return df

# --- FUNÇÃO ATUALIZADA (LÊ DO DRIVE) ---
# O @st.cache_data foi removido daqui e colocado no helper 'download_excel_bytes'
def carregar_base_dados(drive_service, folder_id):
    try:
        # --- MUDANÇA PRINCIPAL ---
        file_buffer = download_excel_bytes(drive_service, folder_id, ARQUIVO_BASE)
        if file_buffer is None:
            st.error("Não foi possível carregar a base de dados do Google Drive.")
            return None, None
        df_base = pd.read_excel(file_buffer, sheet_name="Base", dtype=str, engine='openpyxl')
        # --- FIM DA MUDANÇA ---
            
        df_base.columns = [normalizar_texto(col) for col in df_base.columns.astype(str)]
        col_cod_base = [c for c in df_base.columns if "COD" in c][0]
        col_desc_base = [c for c in df_base.columns if "DESC" in c][0]
        col_valor_base = None
        for col_nome in ["VALOR", "PRECO", "UNIT"]:
            cols_encontradas = [c for c in df_base.columns if col_nome in c and "TOTAL" not in c]
            if cols_encontradas:
                col_valor_base = cols_encontradas[0]; break
        if not col_valor_base:
            st.error("Erro na Base de Dados: Não achei a coluna de Valor Unitário."); return None, None
        df_base = df_base.rename(columns={col_cod_base: "COD_BASE", col_desc_base: "DESC_BASE", col_valor_base: "VALOR_BASE"})
        df_base = df_base[["COD_BASE", "DESC_BASE", "VALOR_BASE"]]
        df_base["COD_BASE"] = df_base["COD_BASE"].astype(str).str.strip()
        df_base = df_base.dropna(subset=["COD_BASE"])
        
        df_base["VALOR_NUM"] = pd.to_numeric(df_base["VALOR_BASE"], errors="coerce").fillna(0)
        df_base["BUSCA_STR"] = "[" + df_base["COD_BASE"] + "] - " + \
                                df_base["DESC_BASE"].str.strip() + \
                                " | R$ " + df_base["VALOR_NUM"].apply(lambda x: f"{x:,.2f}")
                                
        lista_busca = sorted(df_base["BUSCA_STR"].tolist())
        
        base_dados_dict = {}
        for _, row in df_base.iterrows():
            base_dados_dict[row["COD_BASE"].upper()] = {
                "descricao": row["DESC_BASE"],
                "valor": row["VALOR_NUM"]
            }
        
        st.success("Base de dados (do Google Drive) carregada!")
        return base_dados_dict, lista_busca 
    except Exception as e:
        st.error(f"Erro FATAL ao processar a Base de Dados (do Drive): {e}"); return None, None

def configurar_e_calcular_tabela(df_entrada, base_dados):
    if not isinstance(df_entrada, pd.DataFrame):
        df_para_editar = pd.DataFrame(columns=["COD", "QTD", "DESCRICAO", "VALOR UNITARIO"])
    else:
        df_para_editar = df_entrada.copy()
    
    col_valor_unit_padrao = "VALOR UNITARIO"
    col_desc_padrao = "DESCRICAO"
    
    col_valor_unit_lista = [c for c in df_para_editar.columns if "UNIT" in c or ("VALOR" in c and "TOTAL" not in c) or "PRECO" in c]
    col_desc_lista = [c for c in df_para_editar.columns if "DESC" in c]
    colunas_lixo = [c for c in df_para_editar.columns if "TOTAL" in c or "UNNAMED" in c]
    
    col_valor_unit = col_valor_unit_lista[0] if col_valor_unit_lista else col_valor_unit_padrao
    col_desc = col_desc_lista[0] if col_desc_lista else col_desc_padrao
    
    col_tipo_lista = [c for c in df_para_editar.columns if "TIPO" in c]
    col_tipo_real = col_tipo_lista[0] if col_tipo_lista else None
    
    if 'COD' not in df_para_editar.columns: df_para_editar['COD'] = None
    if 'QTD' not in df_para_editar.columns: df_para_editar['QTD'] = 1
    if col_desc not in df_para_editar.columns: df_para_editar[col_desc] = ""
    if col_valor_unit not in df_para_editar.columns: df_para_editar[col_valor_unit] = 0.0

    df_para_editar["COD"] = df_para_editar["COD"].astype(str).str.strip().str.upper()
    df_para_editar = df_para_editar[
        df_para_editar["COD"].notna() &
        (df_para_editar["COD"] != "") &
        (df_para_editar["COD"] != "NAN") &
        (df_para_editar["COD"] != "NONE") &
        (df_para_editar["COD"] != "NONE") 
    ].copy()
    
    df_para_editar["QTD"] = pd.to_numeric(df_para_editar["QTD"], errors="coerce").fillna(1)
    df_para_editar[col_valor_unit] = pd.to_numeric(df_para_editar[col_valor_unit], errors="coerce").fillna(0)
    
    for index, linha in df_para_editar.iterrows():
        cod_atual = str(linha["COD"]).strip().upper()
        desc_atual = linha[col_desc]
        
        tipo_atual = ""
        if col_tipo_real and col_tipo_real in linha:
            tipo_atual = str(linha[col_tipo_real]).strip().upper()
        
        info_base = base_dados.get(cod_atual)
        
        if info_base:
            desc_base = info_base["descricao"]
            valor_base = info_base["valor"] 

            tipos_livro = ["LIVRO", "DICIONARIO", "LIVROS", "DICIONÁRIO"]
            
            if tipo_atual not in tipos_livro:
                df_para_editar.at[index, col_valor_unit] = valor_base

            is_empty = pd.isna(desc_atual) or str(desc_atual).strip() == "" or "CODIGO NAO ENCONTRADO" in str(desc_atual)
            is_standard = (str(desc_atual).strip().upper() == str(desc_base).strip().upper())
            
            if is_empty or is_standard:
                df_para_editar.at[index, col_desc] = desc_base
            
        else:
            df_para_editar.at[index, col_desc] = "--- CODIGO NAO ENCONTRADO ---"
            df_para_editar.at[index, col_valor_unit] = 0.0
            
    df_para_editar["QTD"] = pd.to_numeric(df_para_editar["QTD"], errors="coerce").fillna(1)
    df_para_editar[col_valor_unit] = pd.to_numeric(df_para_editar[col_valor_unit], errors="coerce").fillna(0)
    df_para_editar["Subtotal"] = (df_para_editar["QTD"] * df_para_editar[col_valor_unit]).round(2)
    valor_total = df_para_editar["Subtotal"].sum()
    
    config_colunas = {
        "COD": st.column_config.Column("COD", width="small"),
        col_desc: st.column_config.Column("Descrição", width="large"),
        "QTD": st.column_config.NumberColumn("QTD", width="small"),
        col_valor_unit: st.column_config.NumberColumn("Valor Unit.", format="R$ %.2f", disabled=True, width="small"),
        "Subtotal": st.column_config.NumberColumn("Subtotal", format="R$ %.2f", disabled=True, width="small"),
        "TIPO": None
    }
    
    for col in colunas_lixo: 
        if col not in config_colunas:
            config_colunas[col] = None

    ordem_colunas = ["COD", col_desc, "QTD", col_valor_unit, "Subtotal"]
    
    return df_para_editar, config_colunas, valor_total, ordem_colunas


def set_add_flag():
    item_str = st.session_state.get("busca_item")
    cod = item_str.split(']')[0].replace('[', '') if item_str else None
    qtd = st.session_state.get("qtd_adicionar", 1)
    
    if st.session_state.orcamento_mode == "Gerador de Vale":
        tipo = "VALE"
    elif st.session_state.orcamento_mode == "Pedido de Livro":
        tipo = "LIVRO"
    else:
        tipo = st.session_state.get("tipo_adicionar", "Material") 

    if item_str and cod:
        st.session_state.item_para_adicionar = {"COD": cod, "QTD": qtd, "TIPO": tipo.upper(), "ITEM_STR": item_str}
        st.session_state.busca_item = None
    else:
        st.session_state.item_para_adicionar = "ERRO"


# --- FUNÇÕES PDF (MODIFICADAS PARA NUVEM) ---

@st.cache_data
def converter_imagem_base64(caminho_imagem):
    try:
        with open(caminho_imagem, "rb") as img_file:
            return base64.b64encode(img_file.read()).decode('utf-8')
    except Exception as e:
        st.error(f"Erro ao carregar o logo '{caminho_imagem}': {e}")
        return None

def gerar_html_para_pdf(logo_b64, escola, serie, nome_cliente, data_val, df_mat, df_vale, df_livro, df_integral, df_bilingue, obs_nt_str, obs_pe_str, obs_outras_str, totais):
    html_style = f"""
    <style>
        @page {{ size: A4; margin: 1.5cm; }}
        body {{ font-family: Arial, sans-serif; font-size: 9pt; color: #333; }}
        .header-table {{ width: 100%; padding-bottom: 10px; }}
        .header-table td {{ vertical-align: bottom; padding: 0; width: 33.33%; }}
        .header-table img {{ max-width: 180px; max-height: 80px; }}
        .header-center {{ text-align: center; padding-bottom: 5px; }}
        .header-center .escola {{ font-size: 20pt; font-weight: bold; color: #333; }}
        .header-center .serie {{ font-size: 16pt; color: #555; }}
        .header-right {{ text-align: right; font-size: 10pt; line-height: 1.6; padding-bottom: 5px; }}
        .contato-info {{ text-align: center; padding: 8px 0; font-size: 9pt; color: #555; }}
        .obs-container {{ display: flex; justify-content: space-between; margin-top: 15px; }}
        .obs-box {{ width: 48%; font-size: 10pt; }}
        .obs-box h3 {{ font-size: 13pt; margin-top: 0; margin-bottom: 5px; color: {COR_PRINCIPAL_PAPELARIA}; }}
        .obs-box ul {{ padding-left: 20px; margin: 0; }}
        .obs-outras {{ margin-top: 15px; font-size: 10pt; }}
        .obs-outras h3 {{ font-size: 13pt; margin-top: 0; margin-bottom: 5px; color: {COR_PRINCIPAL_PAPELARIA}; }}
        .obs-outras ul {{ padding-left: 20px; margin: 0; }}
        h2 {{ color: #333; border-bottom: 1px solid #ccc; padding-bottom: 2px; font-size: 14pt; margin-top: 20px; }}
        .item-table {{ width: 100%; border-collapse: collapse; margin-top: 10px; }}
        .item-table th {{ text-align: center; padding: 2px 2px; border-bottom: 2px solid #333; font-size: 8pt; text-transform: uppercase; }}
        .item-table td {{ padding: 2px; vertical-align: middle; }}
        .col-cod {{ width: 12%; }}
        .col-desc {{ width: 50%; }}
        .col-qtd {{ width: 8%; text-align: center; }}
        .col-valor {{ width: 15%; text-align: center; }}
        .col-total {{ width: 15%; text-align: center; }}
        .subtotal-final {{ text-align: right; font-size: 11pt; font-weight: bold; margin-top: 5px; }}
        .total-section {{ margin-top: 25px; width: 100%; border-top: 2px solid {COR_PRINCIPAL_PAPELARIA}; padding-top: 10px; }}
        .subtotal-linha {{ display: none; }}
        .total-geral {{ display: flex; justify-content: space-between; font-size: 16pt; font-weight: bold; color: {COR_PRINCIPAL_PAPELARIA}; padding-top: 8px; width: 350px; margin-left: auto; margin-right: 0; }}
        .obs-finais {{ clear: both; margin-top: 80px; padding-top: 15px; border-top: 0px; font-size: 10pt; color: #555; line-height: 1.2; }}
        .obs-finais h3 {{ font-size: 11pt; color: #333; margin-bottom: 2px; }}
        .obs-finais ul {{ padding-left: 20px; margin-top: 1px; margin-bottom: 0px; padding-top: 1px; padding-bottom: 0px; }}
        .obs-finais ul li {{ margin-bottom: 3px; line-height: 1.2; }}
        .footer-info {{ text-align: center; font-size: 9pt; color: #888; margin-top: 30px; border-top: 1px solid #ccc; padding-top: 5px; }}
    </style>
    """
    
    logo_html = f'<img src="data:image/jpeg;base64,{logo_b64}">' if logo_b64 else f'<h1 style="color:{COR_PRINCIPAL_PAPELARIA};">Blocoos Papelaria</h1>'
    html_body = f"""
    <table class="header-table">
        <tr>
            <td>{logo_html}</td>
            <td class="header-center"><div class="escola">{escola}</div><div class="serie">{serie}</div></td>
            <td class="header-right"><strong>Cliente:</strong> {nome_cliente}<br><strong>Validade:</strong> {data_val}</td>
        </tr>
    </table>
    <div class="contato-info">Rua Souza Pereira, 214 - Centro - Sorocaba/SP | E-mail: escolar@blocoos.com.br | Fone/Whatsapp: (15) 3233-8329</div>
    """
    obs_nt_list = [item for item in obs_nt_str.split('\n') if item.strip()]
    obs_pe_list = [item for item in obs_pe_str.split('\n') if item.strip()]
    obs_outras_list = [item for item in obs_outras_str.split('\n') if item.strip()]
    obs_nt_html = "".join([f"<li>{item}</li>" for item in obs_nt_list])
    obs_pe_html = "".join([f"<li>{item}</li>" for item in obs_pe_list])
    obs_outras_html = "".join([f"<li>{item}</li>" for item in obs_outras_list])
    html_body += f"""
    <div class="obs-container">
        <div class="obs-box"><h3 style='color: {COR_PRINCIPAL_PAPELARIA};'>NÃO ESTÁ INCLUSO</h3><ul>{obs_nt_html or "<li>Nenhum item.</li>"}</ul></div>
        <div class="obs-box"><h3 style='color: {COR_PRINCIPAL_PAPELARIA};'>PARA ESCOLHER</h3><ul>{obs_pe_html or "<li>Nenhum item.</li>"}</ul></div>
    </div>
    """
    if obs_outras_html:
        html_body += f"""<div class="obs-outras"><h3 style='color: {COR_PRINCIPAL_PAPELARIA};'>OUTRAS OBSERVAÇÕES</h3><ul>{obs_outras_html}</ul></div>"""
    
    def criar_tabela_html(df, totais_chave):
        if df.empty:
            return ""
        
        tabela_html = '<table class="item-table">'
        tabela_html += '<thead><tr><th class="col-cod">CÓD.</th><th class="col-desc">DESCRIÇÃO</th><th class="col-qtd">QTD</th><th class.col-valor">VLR. UNIT.</th><th class="col-total">VLR. TOTAL</th></tr></thead>'
        tabela_html += '<tbody>'
        
        col_desc_real = [c for c in df.columns if "DESC" in c][0]
        col_valor_unit_real = [c for c in df.columns if "UNIT" in c or ("VALOR" in c and "TOTAL" not in c) or "PRECO" in c][0]
        
        for _, row in df.iterrows():
            tabela_html += f"""
            <tr>
                <td class="col-cod">{row['COD']}</td>
                <td class="col-desc">{row[col_desc_real]}</td>
                <td class="col-qtd">{row['QTD']}</td>
                <td class="col-valor">R$ {row[col_valor_unit_real]:,.2f}</td>
                <td class="col-total">R$ {row['Subtotal']:,.2f}</td>
            </tr>
            """
        tabela_html += '</tbody></table>'
        if totais[totais_chave] > 0:
            subtotal_formatado = f"{totais[totais_chave]:,.2f}"
            tabela_html += f"<div class='subtotal-final'>Subtotal: R$ {subtotal_formatado}</div>"
        return tabela_html

    if not df_mat.empty:
        html_body += "<h2>Itens de Material Individual</h2>"
        html_body += criar_tabela_html(df_mat, "material")

    if not df_vale.empty:
        html_body += "<h2>Vale de Material Coletivo</h2>"
        html_body += criar_tabela_html(df_vale, "vale")
            
    if not df_livro.empty:
        html_body += "<h2>Livros sob Encomenda</h2>"
        html_body += criar_tabela_html(df_livro, "livro")
    
    if not df_integral.empty:
        html_body += "<h2>Itens do Período Integral</h2>"
        html_body += criar_tabela_html(df_integral, "integral")

    if not df_bilingue.empty:
        html_body += "<h2>Itens do Programa Bilíngue</h2>"
        html_body += criar_tabela_html(df_bilingue, "bilingue")
            
    html_body += f"""
    <div class="total-section">
        <div class="subtotal-linha"><span>Subtotal material:</span><span>R$ {totais['material']:,.2f}</span></div>
        <div class="subtotal-linha"><span>Subtotal vale:</span><span>R$ {totais['vale']:,.2f}</span></div>
        <div class="subtotal-linha"><span>Subtotal livro:</span><span>R$ {totais['livro']:,.2f}</span></div>
        <div class="subtotal-linha"><span>Subtotal integral:</span><span>R$ {totais['integral']:,.2f}</span></div>
        <div class="subtotal-linha"><span>Subtotal bilingue:</span><span>R$ {totais['bilingue']:,.2f}</span></div>
        <div class="total-geral"><span>Valor total do orçamento:</span><span>R$ {totais['geral']:,.2f}</span></div>
    </div>
    """
    obs_finais_lista = [
        "Parcelamos em até 6x (parcela mínima R$ 80,00), nos cartões Visa e Mastercard.",
        "Para lista a partir de R$ 250,00 - desconto de 3% pagamento em dinheiro/PIX (exceto livros didáticos).",
        "Livros sob encomenda mediante pagamento antecipado (consultar prazo de entrega e disponibilidade)",
        "Delivery para pedidos acima de R$30,00 + frete (consultar valor). Prazo de entrega de até 2 dias úteis (troca somente na loja física)"
    ]
    html_obs_finais = "<ul>"
    for item in obs_finais_lista:
        if "Livros sob encomenda" in item:
            html_obs_finais += f"<li><span style='color: red; font-weight: bold;'>{item}</span></li>"
        else:
            html_obs_finais += f"<li>{item}</li>"
    html_obs_finais += "</ul>"
    html_body += f"""
    <div class="obs-finais" style='clear: both;'>
        <h3>OBSERVAÇÕES</h3>
        {html_obs_finais}
    </div>
    """
    html_body += f"""
    <div class="footer-info">
        Orçamento gerado em {datetime.now().strftime('%d/%m/%Y %H:%M')} por Blocoos Papelaria
    </div>
    """
    return f"<html><head>{html_style}</head><body>{html_body}</body></html>"


def gerar_vale_pdf_reportlab(caminho_logo, escola, serie, aluno, responsavel, telefone, df_vale, total_vale):
    """Monta um PDF de Vale Avulso (retorna BytesIO)."""
    
    pdf_buffer = BytesIO()
    doc = SimpleDocTemplate(
        pdf_buffer,
        pagesize=A4,
        rightMargin=1.5*cm, leftMargin=1.5*cm, topMargin=1.5*cm, bottomMargin=1.5*cm
    )
    styles = getSampleStyleSheet()
    elementos = []

    style_escola = ParagraphStyle(name="Escola", parent=styles["Normal"], fontName="Arial-Bold", fontSize=14, alignment=TA_CENTER)
    style_serie = ParagraphStyle(name="Serie", parent=styles["Normal"], fontName="Arial", fontSize=12, alignment=TA_CENTER)
    
    logo_flowable = Spacer(1, 1) 
    try:
        logo_flowable = RLImage(caminho_logo, width=180, height=80) 
        logo_flowable.hAlign = 'LEFT'
    except Exception:
        pass 
            
    header_data = [[logo_flowable, [Paragraph(f"ESCOLA: {escola.upper()}", style_escola), Spacer(1, 12), Paragraph(f"SÉRIE: {serie.upper()}", style_serie)]]]
    header_table = Table(header_data, colWidths=[180 + 10, None]) 
    header_table.setStyle(TableStyle([('VALIGN', (0, 0), (-1, -1), 'MIDDLE'), ('ALIGN', (1, 0), (1, 0), 'CENTER')]))
    elementos.append(header_table)
    elementos.append(Spacer(1, 0.5*cm))

    style_info_aluno = ParagraphStyle(name="InfoAluno", parent=styles["Normal"], fontName="Arial", fontSize=11, leading=14)
    info_text = (f"<b>RESPONSÁVEL:</b> {responsavel.upper()}<br/>"
                 f"<b>ALUNO:</b> {aluno.upper()}<br/>"
                 f"<b>TELEFONE:</b> {telefone}")
    info = Paragraph(info_text, style_info_aluno)
    elementos += [info, Spacer(1, 0.5*cm)]

    col_desc_real = [c for c in df_vale.columns if "DESC" in c][0]
    col_valor_unit_real = [c for c in df_vale.columns if "UNIT" in c or ("VALOR" in c and "TOTAL" not in c) or "PRECO" in c][0]
    
    style_normal_left = ParagraphStyle(name="NormalLeft", parent=styles["Normal"], fontName="Arial", alignment=TA_LEFT, fontSize=9)
    style_normal_center = ParagraphStyle(name="NormalCenter", parent=styles["Normal"], fontName="Arial", alignment=TA_CENTER, fontSize=9)
    
    dados_tabela = [["COD", "DESCRIÇÃO", "QTD", "VLR. UNITÁRIO", "VLR. TOTAL"]]

    for _, row in df_vale.iterrows():
        cod = Paragraph(str(row["COD"]), style_normal_center)
        desc = Paragraph(str(row[col_desc_real]), style_normal_left)
        qtd = Paragraph(str(int(row["QTD"])), style_normal_center)
        
        valor_unit = row[col_valor_unit_real]
        valor_unit_par = Paragraph(f"R$ {valor_unit:,.2f}", style_normal_center) if valor_unit != 0 else Paragraph(f"<font color='red'>R$ 0,00</font>", style_normal_center)
        total_par = Paragraph(f"R$ {row['Subtotal']:,.2f}", style_normal_center)
        
        dados_tabela.append([cod, desc, qtd, valor_unit_par, total_par])

    tabela = Table(dados_tabela, colWidths=[60, 260, 45, 80, 80])
    tabela.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor(COR_PRINCIPAL_PAPELARIA)), 
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white), ('FONTNAME', (0, 0), (-1, 0), "Arial-Bold"),
        ('ALIGN', (0, 0), (-1, 0), "CENTER"), ('VALIGN', (0, 0), (-1, -1), "MIDDLE"),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey), ('ALIGN', (1, 1), (1, -1), "LEFT"), 
        ('ALIGN', (3, 1), (-1, -1), "CENTER"), ('FONTSIZE', (0, 1), (-1, -1), 9), 
        ('TOPPADDING', (0, 0), (-1, -1), 1), ('BOTTOMPADDING', (0, 0), (-1, -1), 1), 
        ('FONTNAME', (0, 1), (-1, -1), 'Arial')
    ]))
    elementos += [tabela, Spacer(1, 12)]

    total_geral = df_vale["Subtotal"].sum()
    style_total = ParagraphStyle(name="Total", parent=styles["Heading3"], fontName="Arial-Bold", alignment=TA_LEFT, fontSize=14)
    total_paragraph = Paragraph(f"<b>Total Geral:</b> R$ {total_geral:,.2f}", style_total)
    elementos.append(total_paragraph)

    style_footer_contato = ParagraphStyle(name="FooterContato", parent=styles["Normal"], fontName="Arial", fontSize=8, textColor=colors.grey, alignment=TA_CENTER, borderTopWidth=1, borderTopColor=colors.lightgrey, paddingTop=5, marginTop=20)
    contato_texto = "Rua Souza Pereira, 214 - Centro - Sorocaba/SP | E-mail: blocoos@blocoos.com.br | Fone/Whatsapp: (15) 3233-8329"
    elementos.append(Paragraph(contato_texto, style_footer_contato))
    
    doc.build(elementos)
    pdf_buffer.seek(0)
    return pdf_buffer.getvalue()


def _build_one_copy_story(caminho_logo, cliente, telefone, df_livro, total_livro, obs_livro, styles, cor_principal):
    """Helper que constrói os 'flowables' para uma única via."""
    elementos = []
    
    style_cliente = ParagraphStyle(name="Cliente", parent=styles["Normal"], fontName="Arial-Bold", fontSize=14, alignment=TA_CENTER)
    style_info = ParagraphStyle(name="Info", parent=styles["Normal"], fontName="Arial", fontSize=11, alignment=TA_CENTER, leading=14)
    style_normal_left = ParagraphStyle(name="NormalLeft", parent=styles["Normal"], fontName="Arial", alignment=TA_LEFT, fontSize=9)
    style_normal_center = ParagraphStyle(name="NormalCenter", parent=styles["Normal"], fontName="Arial", alignment=TA_CENTER, fontSize=9)
    style_checkbox = ParagraphStyle(name="Checkbox", parent=styles["Normal"], fontName="Arial", alignment=TA_CENTER, fontSize=12)
    style_total = ParagraphStyle(name="Total", parent=styles["Heading3"], fontName="Arial-Bold", alignment=TA_LEFT, fontSize=14)
    style_obs_fixa = ParagraphStyle(name="ObsFixa", parent=styles["Normal"], fontName="Arial-Bold", fontSize=9, leading=11, textColor=colors.red, alignment=TA_CENTER)
    
    logo_flowable = Spacer(1, 1)
    try:
        logo_flowable = RLImage(caminho_logo, width=150, height=67)
        logo_flowable.hAlign = 'CENTER'
    except Exception:
        pass
    elementos.append(logo_flowable)
    elementos.append(Spacer(1, 0.5*cm))

    elementos.append(Paragraph("PEDIDO DE LIVROS", style_cliente))
    elementos.append(Spacer(1, 0.2*cm))
    info_text = f"<b>CLIENTE:</b> {cliente.upper()}<br/><b>TELEFONE:</b> {telefone}"
    elementos.append(Paragraph(info_text, style_info))
    elementos.append(Spacer(1, 0.2*cm)) 

    if obs_livro and obs_livro.strip() != "":
        style_obs = ParagraphStyle(name="Obs", parent=styles["Normal"], fontName="Arial", fontSize=9, leading=11, borderWidth=0.5, borderColor=colors.grey, padding=(5, 5, 5, 5), borderRadius=2)
        obs_formatada = obs_livro.replace('\n', '<br/>')
        elementos.append(Paragraph(f"<b>OBSERVAÇÃO:</b><br/>{obs_formatada}", style_obs))
        elementos.append(Spacer(1, 0.3*cm))
    
    elementos.append(Spacer(1, 0.3*cm))
    
    col_desc_real = [c for c in df_livro.columns if "DESC" in c][0]
    
    dados_tabela = [["QTD", "DESCRIÇÃO", "VLR. TOTAL", "ENC.", "ENTR."]]
    checkbox = Paragraph("▢", style_checkbox)

    for _, row in df_livro.iterrows():
        desc = Paragraph(str(row[col_desc_real]), style_normal_left)
        qtd = Paragraph(str(int(row["QTD"])), style_normal_center)
        total_par = Paragraph(f"R$ {row['Subtotal']:,.2f}", style_normal_center)
        dados_tabela.append([qtd, desc, total_par, checkbox, checkbox])

    tabela = Table(dados_tabela, colWidths=[45, 295, 80, 35, 35]) 
    tabela.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor(cor_principal)),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white), ('FONTNAME', (0, 0), (-1, 0), "Arial-Bold"),
        ('ALIGN', (0, 0), (-1, 0), "CENTER"), ('VALIGN', (0, 0), (-1, -1), "MIDDLE"),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey), ('ALIGN', (1, 1), (1, -1), "LEFT"),
        ('ALIGN', (0, 1), (0, -1), "CENTER"), ('ALIGN', (2, 1), (-1, -1), "CENTER"), 
        ('FONTSIZE', (0, 1), (-1, -1), 9), ('TOPPADDING', (0, 0), (-1, -1), 1), 
        ('BOTTOMPADDING', (0, 0), (-1, -1), 1), ('FONTNAME', (0, 1), (-1, -1), 'Arial')
    ]))
    elementos += [tabela, Spacer(1, 12)]

    total_paragraph = Paragraph(f"<b>Total Geral:</b> R$ {total_livro:,.2f}", style_total)
    elementos.append(total_paragraph)
    
    elementos.append(Spacer(1, 0.5*cm))
    obs_fixa_texto = "Prazo de encomenda é aproximadamente 15 dias úteis e sujeito a disponibilidade do livro no fornecedor."
    elementos.append(Paragraph(obs_fixa_texto, style_obs_fixa))
    
    return elementos


def gerar_pedido_livro_pdf_reportlab(caminho_logo, cliente, telefone, df_livro, total_livro, obs_livro):
    """Monta um PDF de Pedido de Livro com 2 vias (retorna BytesIO)."""
    
    pdf_buffer = BytesIO()
    doc = SimpleDocTemplate(
        pdf_buffer,
        pagesize=A4, 
        rightMargin=1.5*cm, leftMargin=1.5*cm, topMargin=1.5*cm, bottomMargin=1.5*cm
    )
    styles = getSampleStyleSheet()
    
    largura_pagina, altura_pagina = A4
    altura_frame = (altura_pagina - 3*cm) / 2 
    largura_frame = largura_pagina - 3*cm
    
    frame_cima = Frame(x1=doc.leftMargin, y1=doc.bottomMargin + altura_frame + 1.5*cm, width=largura_frame, height=altura_frame, id='frame_cima')
    frame_baixo = Frame(x1=doc.leftMargin, y1=doc.bottomMargin, width=largura_frame, height=altura_frame, id='frame_baixo')
    
    def linha_divisoria(canvas, doc):
        canvas.saveState()
        canvas.setStrokeColorRGB(0.7, 0.7, 0.7)
        canvas.setDash(1, 2)
        meio_pagina_y = doc.bottomMargin + altura_frame + (1.5*cm / 2)
        canvas.line(doc.leftMargin, meio_pagina_y, largura_pagina - doc.rightMargin, meio_pagina_y)
        canvas.restoreState()

    template_duplo = PageTemplate(id='DuasVias', frames=[frame_cima, frame_baixo], onPage=linha_divisoria)
    doc.addPageTemplates([template_duplo])

    telefone_formatado = formatar_telefone(telefone)
    
    elementos_uma_via = _build_one_copy_story(
        caminho_logo, cliente, telefone_formatado, df_livro, total_livro, obs_livro, styles, COR_PRINCIPAL_PAPELARIA
    )
    
    historia_completa = elementos_uma_via + [FrameBreak()] + elementos_uma_via
    
    doc.build(historia_completa)
    pdf_buffer.seek(0)
    return pdf_buffer.getvalue()


def converter_html_para_pdf(html_string):
    """Converte uma string de HTML em bytes de PDF."""
    pdf_buffer = BytesIO()
    pisa_status = pisa.CreatePDF(html_string, dest=pdf_buffer, encoding='utf-8')
    if not pisa_status.err:
        pdf_buffer.seek(0)
        return pdf_buffer.getvalue()
    else:
        st.error(f"Erro ao gerar PDF: {pisa_status.err}")
        return None
# --- FIM DAS FUNÇÕES PDF ---

@st.cache_resource
def autenticar_google():
    """Autentica com Google usando o client_secret.json armazenado em st.secrets."""
    try:
        if "google_creds" not in st.secrets:
            st.error("Erro: 'google_creds' (client_secret.json) não encontrado nos Segredos do Streamlit.")
            st.info("Adicione o conteúdo do seu client_secret.json em 'google_creds' no painel de segredos.")
            return None, None

        # Converte string JSON em dicionário
        creds_json = st.secrets["google_creds"]
        if isinstance(creds_json, str):
            creds_json = json.loads(creds_json)

        # Cria o fluxo OAuth
        flow = InstalledAppFlow.from_client_secrets_info(creds_json, SCOPES)

        # Executa o login via console (funciona no Streamlit Cloud)
        creds = flow.run_console()

        # Conecta com os serviços Google Drive e Sheets
        drive_service = build('drive', 'v3', credentials=creds)
        sheets_service = gspread.authorize(creds)

        return drive_service, sheets_service

    except Exception as e:
        st.error(f"Erro na autenticação do Google: {e}")
        st.exception(e)
        return None, None

drive_service, sheets_service = autenticar_google()

def find_file_in_drive(drive_service, pasta_id, nome_arquivo):
    """Procura um arquivo pelo nome exato dentro de uma pasta específica."""
    query = f"name='{nome_arquivo}' and '{pasta_id}' in parents and trashed=false"
    response = drive_service.files().list(q=query, spaces='drive', fields='files(id, name)').execute()
    files = response.get('files', [])
    return files[0].get('id') if files else None

# --- FUNÇÃO ATUALIZADA (LÊ DO DRIVE) ---
@st.cache_data(ttl=600)
def build_full_database(_drive_service, _base_dados):
    """Varre todos os Excels no Google Drive para criar um 'Super DF' de busca."""
    
    lista_completa = []
    
    try:
        # 1. Listar todos os arquivos na pasta
        query = f"'{PASTA_DRIVE_PLANILHAS_ID}' in parents and mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' and trashed=false"
        response = _drive_service.files().list(q=query, spaces='drive', fields='files(id, name)').execute()
        files = response.get('files', [])
        
        if not files:
            st.warning("Nenhuma planilha de escola encontrada no Google Drive.")
            return pd.DataFrame()

    except Exception as e:
        st.error(f"Erro ao listar arquivos do Google Drive: {e}")
        return pd.DataFrame()

    st.toast(f"Encontrados {len(files)} arquivos de escola. Iniciando busca global...")
    progress_bar = st.progress(0.0)
    
    for i, file in enumerate(files):
        nome_arquivo_excel = file.get('name')
        if nome_arquivo_excel == ARQUIVO_BASE or nome_arquivo_excel.startswith("~$"):
            continue
            
        progress_bar.progress((i+1) / len(files), text=f"Lendo: {nome_arquivo_excel}")

        try:
            file_buffer = download_excel_bytes(_drive_service, PASTA_DRIVE_PLANILHAS_ID, nome_arquivo_excel)
            if file_buffer is None:
                continue
                
            nome_escola_limpo = nome_arquivo_excel.split('.')[0].split('_')[0]
            # Carrega o arquivo em memória para pegar os nomes das abas
            xl = pd.ExcelFile(file_buffer, engine='openpyxl')
            abas = xl.sheet_names
            
            for nome_aba in abas:
                # Re-lê o buffer para cada aba (necessário pois pd.read_excel fecha o buffer)
                file_buffer_aba = download_excel_bytes(_drive_service, PASTA_DRIVE_PLANILHAS_ID, nome_arquivo_excel)
                df_itens = pd.read_excel(file_buffer_aba, sheet_name=nome_aba, header=HEADER_DA_PLANILHA, dtype=str, engine='openpyxl')
                
                # O carregar_itens() faz muita validação, vamos simplificar aqui
                df_itens.columns = [normalizar_texto(col) for col in df_itens.columns.astype(str)]
                if "COD" not in df_itens.columns or "TIPO" not in df_itens.columns:
                    continue # Pula aba mal formatada
                
                df_itens = df_itens.rename(columns={
                    [c for c in df_itens.columns if "COD" in c][0]: "COD",
                    [c for c in df_itens.columns if "TIPO" in c][0]: "TIPO"
                })
                df_itens["COD"] = df_itens["COD"].astype(str).str.strip().str.upper()
                if "QTD" not in df_itens.columns: df_itens["QTD"] = 1
                
                df_itens['QTD'] = pd.to_numeric(df_itens['QTD'], errors='coerce').fillna(0)

                for _, row in df_itens.iterrows():
                    cod = row['COD']
                    info_base = _base_dados.get(cod)
                    
                    if info_base:
                        desc_planilha = row.get("DESCRICAO", "")
                        desc_base = info_base['descricao']
                        
                        is_empty = pd.isna(desc_planilha) or str(desc_planilha).strip() == ""
                        is_standard = (normalizar_texto(desc_planilha) == normalizar_texto(desc_base))
                        
                        descricao_final = desc_base if (is_empty or is_standard) else desc_planilha
                        
                        lista_completa.append({
                            'Escola': nome_escola_limpo,
                            'Série': nome_aba,
                            'COD': cod,
                            'Descrição': descricao_final,
                            'QTD': int(row['QTD']),
                            'TIPO': normalizar_texto(row['TIPO'])
                        })
        except Exception as e:
            print(f"ALERTA: Falha ao ler {nome_arquivo_excel} do Drive: {e}")
            
    progress_bar.empty()
    return pd.DataFrame(lista_completa)

# --- FUNÇÃO ATUALIZADA (LÊ E ESCREVE NO DRIVE) ---
def run_batch_update(base_dados, sheets_service, drive_service):
    planilha_mestra = None
    aba_mestra = None
    try:
        planilha_mestra = sheets_service.open_by_key(PLANILHA_MESTRE_ID)
        aba_mestra = planilha_mestra.sheet1
        st.success("Planilha Mestra do Make.com aberta!")
    except Exception as e:
        st.error(f"Não consegui abrir a Planilha Mestra (ID: {PLANILHA_MESTRE_ID}). Erro: {e}")
        return

    try:
        dados_mestra = aba_mestra.get_all_records() 
        mapa_planilha = {}
        for i, linha in enumerate(dados_mestra):
            escola = str(linha.get("Escola")).strip()
            serie = str(linha.get("Série")).strip()
            if escola and serie:
                mapa_planilha[(escola, serie)] = i + 2 
    except Exception as e:
        st.error(f"Não consegui ler os dados da Planilha Mestra. Erro: {e}")
        return

    st.info(f"Encontrei {len(mapa_planilha)} linhas de Escola/Série na Planilha Mestra.")
    
    logo_base64 = converter_imagem_base64(LOGO_PATH)
    data_validade = extrair_data_validade(drive_service, PASTA_DRIVE_PLANILHAS_ID)
    
    try:
        query = f"'{PASTA_DRIVE_PLANILHAS_ID}' in parents and mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' and trashed=false"
        response = drive_service.files().list(q=query, spaces='drive', fields='files(id, name)').execute()
        arquivos_escola = response.get('files', [])
    except Exception as e:
        st.error(f"Erro ao listar arquivos do Google Drive: {e}")
        return

    progress_bar = st.progress(0.0)
    status_text = st.empty()
    arquivos_processados = 0
    
    for i, file in enumerate(arquivos_escola):
        nome_arquivo_excel = file.get('name')
        if nome_arquivo_excel == ARQUIVO_BASE or nome_arquivo_excel.startswith("~$"):
            continue

        nome_escola_limpo = nome_arquivo_excel.split('.')[0].split('_')[0]
        status_text.text(f"Processando: {nome_escola_limpo} ({i+1}/{len(arquivos_escola)})...")
        
        try:
            file_buffer = download_excel_bytes(drive_service, PASTA_DRIVE_PLANILHAS_ID, nome_arquivo_excel)
            if file_buffer is None:
                continue

            xl = pd.ExcelFile(file_buffer, engine='openpyxl')
            abas = xl.sheet_names
            
            for nome_aba in abas:
                # Recarrega os bytes, pois o pandas/openpyxl consome o buffer
                file_buffer_aba = download_excel_bytes(drive_service, PASTA_DRIVE_PLANILHAS_ID, nome_arquivo_excel)
                df_itens = pd.read_excel(file_buffer_aba, sheet_name=nome_aba, header=HEADER_DA_PLANILHA, dtype=str, engine='openpyxl')
                
                # Validação simplificada
                df_itens.columns = [normalizar_texto(col) for col in df_itens.columns.astype(str)]
                if "COD" not in df_itens.columns or "TIPO" not in df_itens.columns:
                    continue
                
                col_tipo = [c for c in df_itens.columns if "TIPO" in c][0]
                df_itens = df_itens.rename(columns={
                    [c for c in df_itens.columns if "COD" in c][0]: "COD",
                    col_tipo: "TIPO"
                })
                if "QTD" not in df_itens.columns: df_itens["QTD"] = 1
                df_itens["COD"] = df_itens["COD"].astype(str).str.strip() 
                df_itens[col_tipo] = df_itens[col_tipo].astype(str).apply(normalizar_texto)
                
                tipos_livro = ["LIVRO", "DICIONARIO", "LIVROS", "DICIONÁRIO"]
                tipos_integral = ["INTEGRAL"]
                tipos_bilingue = ["BILINGUE", "BILINGÜE"]

                df_livro = df_itens[df_itens[col_tipo].isin(tipos_livro)]
                df_vale = df_itens[df_itens[col_tipo] == "VALE"]
                df_integral = df_itens[df_itens[col_tipo].isin(tipos_integral)]
                df_bilingue = df_itens[df_itens[col_tipo].isin(tipos_bilingue)]
                df_material = df_itens[
                    (~df_itens[col_tipo].isin(tipos_livro)) &
                    (df_itens[col_tipo] != "VALE") &
                    (~df_itens[col_tipo].isin(tipos_integral)) &
                    (~df_itens[col_tipo].isin(tipos_bilingue))
                ]
                
                df_mat_final, _, total_material, _ = configurar_e_calcular_tabela(df_material, base_de_dados)
                df_vale_final, _, total_vale, _ = configurar_e_calcular_tabela(df_vale, base_de_dados)
                df_livro_final, _, total_livro, _ = configurar_e_calcular_tabela(df_livro, base_de_dados)
                df_integral_final, _, total_integral, _ = configurar_e_calcular_tabela(df_integral, base_de_dados)
                df_bilingue_final, _, total_bilingue, _ = configurar_e_calcular_tabela(df_bilingue, base_de_dados)
                
                totais = {
                    "material": total_material, "vale": total_vale, "livro": total_livro,
                    "integral": total_integral, "bilingue": total_bilingue,
                    "geral": total_material + total_vale + total_livro + total_integral + total_bilingue
                }
                
                # Re-download para o openpyxl
                file_buffer_obs = download_excel_bytes(drive_service, PASTA_DRIVE_PLANILHAS_ID, nome_arquivo_excel)
                wb_escola = openpyxl.load_workbook(file_buffer_obs, data_only=True)
                ws_escola = wb_escola[nome_aba]
                nt_str, pe_str = extrair_observacoes_iniciais(drive_service, PASTA_DRIVE_PLANILHAS_ID, nome_arquivo_excel, nome_aba)
                
                html_string = gerar_html_para_pdf(
                    logo_base64, nome_escola_limpo, nome_aba, "Orçamento Padrão", data_validade,
                    df_mat_final, df_vale_final, df_livro_final, df_integral_final, df_bilingue_final, 
                    nt_str, pe_str, "", totais
                )
                pdf_bytes = converter_html_para_pdf(html_string)
                
                if not pdf_bytes:
                    st.warning(f"  ↪ Falha ao gerar PDF para {nome_escola_limpo} - {nome_aba}. Pulando.")
                    continue
                    
                linha_para_atualizar = mapa_planilha.get((nome_escola_limpo, nome_aba))
                if not linha_para_atualizar:
                    continue
                    
                nome_pdf_final = f"Orcamento {nome_escola_limpo} {nome_aba}.pdf"
                
                arquivo_antigo_id = find_file_in_drive(drive_service, PASTA_DRIVE_PRINCIPAL_ID, nome_pdf_final)
                
                if arquivo_antigo_id:
                    drive_service.files().update(
                        fileId=arquivo_antigo_id,
                        addParents=PASTA_DRIVE_ARQUIVO_MORTO_ID, 
                        removeParents=PASTA_DRIVE_PRINCIPAL_ID 
                    ).execute()
                
                pdf_stream = BytesIO(pdf_bytes)
                media_body = MediaIoBaseUpload(pdf_stream, mimetype='application/pdf', resumable=True)
                file_metadata = {'name': nome_pdf_final, 'parents': [PASTA_DRIVE_PRINCIPAL_ID]}
                
                novo_arquivo = drive_service.files().create(
                    body=file_metadata, media_body=media_body, fields='id, webViewLink'
                ).execute()
                
                novo_link = novo_arquivo.get('webViewLink')
                
                aba_mestra.update_cell(linha_para_atualizar, 4, novo_link) 
                aba_mestra.update_cell(linha_para_atualizar, 5, "Atualizado")
                
                arquivos_processados += 1
                
        except Exception as e:
            st.error(f"Erro INESPERADO ao processar o arquivo {nome_arquivo_excel}: {e}")
            
        progress_bar.progress((i + 1) / len(arquivos_escola))
    
    status_text.success(f"🎉 Processo Concluído! {arquivos_processados} PDFs foram gerados e atualizados no Google Drive.")
# --- FIM DAS FUNÇÕES DO GOOGLE ---


# --- Interface Streamlit ---
st.set_page_config(page_title="Orçamento Escolar", layout="wide")
try:
    st.image(LOGO_PATH, width=150)
except Exception as e:
    st.error(f"Não foi possível carregar o Logo.jpg. Verifique se ele está no repositório GitHub. Erro: {e}")

st.title("Editor de Orçamento Escolar 📚")

def limpar_state_para_novo_modo():
    if st.session_state.get("carregando_rascunho", False):
        st.session_state.carregando_rascunho = False 
        return

    st.session_state.df_material = pd.DataFrame()
    st.session_state.df_vale = pd.DataFrame()
    st.session_state.df_livro = pd.DataFrame() 
    st.session_state.df_integral = pd.DataFrame() 
    st.session_state.df_bilingue = pd.DataFrame() 
    st.session_state.obs_nao_trabalhamos = ""
    st.session_state.obs_para_escolher = ""
    st.session_state.obs_outras = ""
    st.session_state.escola_manual = ""
    st.session_state.serie_manual = ""
    st.session_state.aba_anterior = None
    st.session_state.escola_anterior = None
    st.session_state.vale_aluno = ""
    st.session_state.vale_responsavel = ""
    st.session_state.vale_telefone = ""
    st.session_state.nome_cliente = "" 
    st.session_state.livro_cliente = "" 
    st.session_state.livro_telefone = "" 
    st.session_state.livro_obs = "" 


# --- NOVA INICIALIZAÇÃO (LENDO DO GOOGLE DRIVE) ---
# SUBSTITUA A FUNÇÃO ANTIGA "autenticar_google" POR ESTA VERSÃO:

@st.cache_resource
def autenticar_google():
    """Autentica com Google usando o fluxo OAuth do Streamlit (client_secret.json)."""
    try:
        if "google_creds" not in st.secrets:
            st.error("Erro Crítico: 'google_creds' (client_secret.json) não encontrado nos Segredos do Streamlit.")
            st.info("Por favor, cole o conteúdo do seu client_secret.json (o primeiro que você criou) para o st.secrets.")
            return None, None

        # O Streamlit Cloud gerencia o token.json automaticamente.
        # Esta é a forma correta de usar o InstalledAppFlow
        flow = InstalledAppFlow.from_client_secrets_info(
            st.secrets["google_creds"]["web"],
            SCOPES,
            redirect_uri='urn:ietf:wg:oauth:2.0:oob' # Essencial para o Streamlit Cloud
        )

        # O Streamlit Cloud lida com este fluxo magicamente
        # Ele vai pausar o app e te dar um link para autenticar no log
        creds = flow.run_console()
        
        drive_service = build('drive', 'v3', credentials=creds)
        sheets_service = gspread.authorize(creds)

        return drive_service, sheets_service
        
    except Exception as e:
        st.error(f"Erro na autenticação do Google: {e}")
        st.info("O Streamlit tentará abrir uma aba de autenticação. Por favor, autorize e recarregue a página.")
        st.exception(e)
        return None, None
# --- FIM DA SUBSTITUIÇÃO ---

        return drive_service, sheets_service
        
    except Exception as e:
        st.error(f"Erro na autenticação do Google: {e}")
        st.exception(e)
        return None, None

drive_service, sheets_service = autenticar_google()

if drive_service:
    # Carrega a base de dados principal
    base_de_dados, lista_para_busca = carregar_base_dados(drive_service, PASTA_DRIVE_PLANILHAS_ID)
    
    # Carrega a lista de escolas do Drive
    @st.cache_data(ttl=600)
    def get_school_list(drive_service, folder_id):
        try:
            query = f"'{folder_id}' in parents and mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' and trashed=false"
            response = drive_service.files().list(q=query, spaces='drive', fields='files(id, name)').execute()
            files = response.get('files', [])
            
            school_files = [f.get('name') for f in files if f.get('name') != ARQUIVO_BASE and not f.get('name').startswith("~$")]
            return sorted([f.replace(".xlsx", "") for f in school_files])
        except Exception as e:
            st.error(f"Não foi possível carregar a lista de escolas do Drive: {e}")
            return []
            
    opcoes_escolas = get_school_list(drive_service, PASTA_DRIVE_PLANILHAS_ID)
    
else:
    st.error("Autenticação do Google falhou. O aplicativo não pode carregar os dados das planilhas.")
    base_de_dados, lista_para_busca = None, []
    opcoes_escolas = []
# --- FIM DA NOVA INICIALIZAÇÃO ---


if 'orcamento_mode' not in st.session_state:
    st.session_state.orcamento_mode = "Novo Orçamento"
if 'df_material' not in st.session_state: st.session_state.df_material = pd.DataFrame()
if 'df_vale' not in st.session_state: st.session_state.df_vale = pd.DataFrame()
if 'df_livro' not in st.session_state: st.session_state.df_livro = pd.DataFrame() 
if 'df_integral' not in st.session_state: st.session_state.df_integral = pd.DataFrame() 
if 'df_bilingue' not in st.session_state: st.session_state.df_bilingue = pd.DataFrame() 
if 'aba_anterior' not in st.session_state: st.session_state.aba_anterior = None
if 'escola_anterior' not in st.session_state: st.session_state.escola_anterior = None
if 'obs_nao_trabalhamos' not in st.session_state: st.session_state.obs_nao_trabalhamos = ""
if 'obs_para_escolher' not in st.session_state: st.session_state.obs_para_escolher = ""
if 'obs_outras' not in st.session_state: st.session_state.obs_outras = ""
if 'escola_manual' not in st.session_state: st.session_state.escola_manual = ""
if 'serie_manual' not in st.session_state: st.session_state.serie_manual = ""
if 'vale_aluno' not in st.session_state: st.session_state.vale_aluno = ""
if 'vale_responsavel' not in st.session_state: st.session_state.vale_responsavel = ""
if 'vale_telefone' not in st.session_state: st.session_state.vale_telefone = ""
if 'nome_cliente' not in st.session_state: st.session_state.nome_cliente = ""
if 'carregando_rascunho' not in st.session_state: st.session_state.carregando_rascunho = False
if 'livro_cliente' not in st.session_state: st.session_state.livro_cliente = ""
if 'livro_telefone' not in st.session_state: st.session_state.livro_telefone = ""
if 'livro_obs' not in st.session_state: st.session_state.livro_obs = ""


if "next_mode" in st.session_state:
    st.session_state.orcamento_mode = st.session_state.pop("next_mode")

st.radio(
    "Selecione o modo de operação:",
    ("Novo Orçamento", "Orçamento Escola Pronto", "Carregar Rascunho", "Gerador de Vale", "Pedido de Livro", "Buscador Itens", "Atualizador PDF"), 
    key="orcamento_mode",
    horizontal=True,
    on_change=limpar_state_para_novo_modo 
)

escola_final = None
serie_final = None
pode_carregar = False 

# --- RASCUNHOS NÃO SÃO SUPORTADOS NA NUVEM DESTA FORMA ---
if st.session_state.orcamento_mode == "Carregar Rascunho":
    st.header("📂 Carregar Rascunho Salvo")
    st.error("O carregamento de rascunhos salvos localmente não é suportado na versão em nuvem.")
    st.info("Para salvar um orçamento, use o modo 'Novo Orçamento' ou 'Orçamento Escola Pronto' e gere o PDF final.")
    pode_carregar = False
# --- FIM DA MUDANÇA ---
    
elif st.session_state.orcamento_mode == "Orçamento Escola Pronto":
    if not drive_service or not opcoes_escolas:
        st.error("Serviço do Google Drive não está disponível ou nenhuma planilha de escola foi encontrada.")
        st.stop()
        
    escola_selecionada = st.selectbox(
        "Digite o nome da escola", 
        options=opcoes_escolas, 
        placeholder="Comece a digitar...", 
        index=None,
        key="escola_selecionada_select" 
    )

    if escola_selecionada:
        nome_arquivo = escola_selecionada + ".xlsx"
        try:
            # --- MUDANÇA (LÊ DO DRIVE) ---
            file_buffer = download_excel_bytes(drive_service, PASTA_DRIVE_PLANILHAS_ID, nome_arquivo)
            if file_buffer is None:
                st.error(f"Não foi possível carregar o arquivo {nome_arquivo} do Google Drive.")
                st.stop()
            xl = pd.ExcelFile(file_buffer, engine='openpyxl')
            abas = xl.sheet_names
            # --- FIM DA MUDANÇA ---
        except Exception as e:
            st.error(f"Não foi possível ler as abas do arquivo '{nome_arquivo}': {e}"); st.stop()
            
        aba_selecionada = st.selectbox("Escolha a série", abas)
        
        if aba_selecionada:
            pode_carregar = True
            escola_final = escola_selecionada
            serie_final = aba_selecionada

            if (escola_selecionada != st.session_state.escola_anterior) or (aba_selecionada != st.session_state.aba_anterior):
                st.toast(f"Carregando dados para {escola_selecionada} - {aba_selecionada}...")
                
                # --- MUDANÇA (LÊ DO DRIVE) ---
                df_itens = carregar_itens(drive_service, PASTA_DRIVE_PLANILHAS_ID, nome_arquivo, aba_selecionada)
                
                if not df_itens.empty:
                    col_tipo = [c for c in df_itens.columns if "TIPO" in c][0]
                    df_itens[col_tipo] = df_itens[col_tipo].astype(str).str.upper()
                    
                    tipos_livro = ["LIVRO", "DICIONARIO", "LIVROS", "DICIONÁRIO"]
                    tipos_integral = ["INTEGRAL"]
                    tipos_bilingue = ["BILINGUE", "BILINGÜE"]
                    
                    st.session_state.df_livro = df_itens[df_itens[col_tipo].isin(tipos_livro)]
                    st.session_state.df_vale = df_itens[df_itens[col_tipo] == "VALE"]
                    st.session_state.df_integral = df_itens[df_itens[col_tipo].isin(tipos_integral)]
                    st.session_state.df_bilingue = df_itens[df_itens[col_tipo].isin(tipos_bilingue)]
                    st.session_state.df_material = df_itens[
                        (~df_itens[col_tipo].isin(tipos_livro)) &
                        (df_itens[col_tipo] != "VALE") &
                        (~df_itens[col_tipo].isin(tipos_integral)) &
                        (~df_itens[col_tipo].isin(tipos_bilingue))
                    ]

                    st.session_state.df_material, _, _, _ = configurar_e_calcular_tabela(st.session_state.df_material, base_de_dados)
                    st.session_state.df_vale, _, _, _ = configurar_e_calcular_tabela(st.session_state.df_vale, base_de_dados)
                    st.session_state.df_livro, _, _, _ = configurar_e_calcular_tabela(st.session_state.df_livro, base_de_dados)
                    st.session_state.df_integral, _, _, _ = configurar_e_calcular_tabela(st.session_state.df_integral, base_de_dados)
                    st.session_state.df_bilingue, _, _, _ = configurar_e_calcular_tabela(st.session_state.df_bilingue, base_de_dados)
                else:
                    st.session_state.df_material = pd.DataFrame(); st.session_state.df_vale = pd.DataFrame(); st.session_state.df_livro = pd.DataFrame(); st.session_state.df_integral = pd.DataFrame(); st.session_state.df_bilingue = pd.DataFrame()
                
                # --- MUDANÇA (LÊ DO DRIVE) ---
                nt, pe = extrair_observacoes_iniciais(drive_service, PASTA_DRIVE_PLANILHAS_ID, nome_arquivo, aba_selecionada)
                st.session_state.obs_nao_trabalhamos = nt; st.session_state.obs_para_escolher = pe; st.session_state.obs_outras = "" 

                st.session_state.escola_anterior = escola_selecionada
                st.session_state.aba_anterior = aba_selecionada
                st.rerun() 
    
elif st.session_state.orcamento_mode == "Novo Orçamento":
    col1, col2 = st.columns(2)
    with col1:
        st.text_input("Nome da Escola:", key="escola_manual")
    with col2:
        st.text_input("Série:", key="serie_manual")
    
    if st.session_state.escola_manual and st.session_state.serie_manual:
        pode_carregar = True
        escola_final = st.session_state.escola_manual
        serie_final = st.session_state.serie_manual

elif st.session_state.orcamento_mode == "Gerador de Vale":
    st.header("📄 Gerador de Vale Avulso")
    
    if not drive_service or not opcoes_escolas:
        st.error("Serviço do Google Drive não está disponível ou nenhuma planilha de escola foi encontrada.")
        st.stop()
        
    escola_selecionada = st.selectbox(
        "Selecione a escola", 
        options=opcoes_escolas, 
        placeholder="Comece a digitar...", 
        index=None,
        key="escola_selecionada_vale" 
    )
    
    aba_selecionada = None
    if escola_selecionada:
        nome_arquivo = escola_selecionada + ".xlsx"
        try:
            # --- MUDANÇA (LÊ DO DRIVE) ---
            file_buffer = download_excel_bytes(drive_service, PASTA_DRIVE_PLANILHAS_ID, nome_arquivo)
            if file_buffer is None:
                st.error(f"Não foi possível carregar o arquivo {nome_arquivo} do Google Drive.")
                st.stop()
            xl = pd.ExcelFile(file_buffer, engine='openpyxl')
            abas = xl.sheet_names
            # --- FIM DA MUDANÇA ---
        except Exception as e:
            st.error(f"Não foi possível ler as abas do arquivo '{nome_arquivo}': {e}"); st.stop()
            
        aba_selecionada = st.selectbox("Escolha a série", abas)
        
        if aba_selecionada:
            pode_carregar = True
            escola_final = escola_selecionada
            serie_final = aba_selecionada

            if (escola_selecionada != st.session_state.escola_anterior) or (aba_selecionada != st.session_state.aba_anterior):
                st.toast(f"Carregando dados para {escola_selecionada} - {aba_selecionada}...")
                df_itens = carregar_itens(drive_service, PASTA_DRIVE_PLANILHAS_ID, nome_arquivo, aba_selecionada)
                
                if not df_itens.empty:
                    col_tipo = [c for c in df_itens.columns if "TIPO" in c][0]
                    df_itens[col_tipo] = df_itens[col_tipo].astype(str).str.upper()
                    st.session_state.df_vale = df_itens[df_itens[col_tipo] == "VALE"]
                    st.session_state.df_vale, _, _, _ = configurar_e_calcular_tabela(st.session_state.df_vale, base_de_dados)
                    st.session_state.df_material = pd.DataFrame() 
                    st.session_state.df_livro = pd.DataFrame() 
                    st.session_state.df_integral = pd.DataFrame()
                    st.session_state.df_bilingue = pd.DataFrame()
                else:
                    st.session_state.df_material = pd.DataFrame(); st.session_state.df_vale = pd.DataFrame(); st.session_state.df_livro = pd.DataFrame(); st.session_state.df_integral = pd.DataFrame(); st.session_state.df_bilingue = pd.DataFrame()
                
                st.session_state.obs_nao_trabalhamos = ""; st.session_state.obs_para_escolher = ""; st.session_state.obs_outras = "" 
                st.session_state.escola_anterior = escola_selecionada
                st.session_state.aba_anterior = aba_selecionada
                st.rerun()
    
    if not escola_selecionada:
        pode_carregar = True 
        escola_final = "Não selecionada" 
        serie_final = "Não selecionada" 
            
elif st.session_state.orcamento_mode == "Pedido de Livro":
    st.header("📚 Gerador de Pedido de Livro")
    
    if not drive_service or not opcoes_escolas:
        st.error("Serviço do Google Drive não está disponível ou nenhuma planilha de escola foi encontrada.")
        st.stop()
        
    escola_selecionada = st.selectbox(
        "Selecione a escola (Opcional, para carregar livros)", 
        options=opcoes_escolas, 
        placeholder="Comece a digitar...", 
        index=None,
        key="escola_selecionada_livro" 
    )
    
    aba_selecionada = None 
    if escola_selecionada:
        nome_arquivo = escola_selecionada + ".xlsx"
        try:
            file_buffer = download_excel_bytes(drive_service, PASTA_DRIVE_PLANILHAS_ID, nome_arquivo)
            if file_buffer is None:
                st.error(f"Não foi possível carregar o arquivo {nome_arquivo} do Google Drive.")
                st.stop()
            xl = pd.ExcelFile(file_buffer, engine='openpyxl')
            abas = xl.sheet_names
        except Exception as e:
            st.error(f"Não foi possível ler as abas do arquivo '{nome_arquivo}': {e}"); st.stop()
            
        aba_selecionada = st.selectbox("Escolha a série", abas, key="aba_selecionada_livro") 
    
    col1, col2 = st.columns(2)
    with col1:
        st.text_input("Nome do Cliente:", key="livro_cliente")
    with col2:
        st.text_input(
            "Telefone:", 
            key="livro_telefone", 
            placeholder="(xx) xxxxx.xxxx",
            max_chars=15
        )
    st.text_area("Observação (Opcional):", key="livro_obs", height=100)
    
    
    if escola_selecionada and aba_selecionada:
        pode_carregar = True
        escola_final = escola_selecionada
        serie_final = aba_selecionada

        if (escola_selecionada != st.session_state.escola_anterior) or (aba_selecionada != st.session_state.aba_anterior):
            st.toast(f"Carregando livros para {escola_selecionada} - {aba_selecionada}...")
            df_itens = carregar_itens(drive_service, PASTA_DRIVE_PLANILHAS_ID, nome_arquivo, aba_selecionada)
            
            if not df_itens.empty:
                col_tipo = [c for c in df_itens.columns if "TIPO" in c][0]
                df_itens[col_tipo] = df_itens[col_tipo].astype(str).str.upper()
                
                tipos_livro = ["LIVRO", "DICIONARIO", "LIVROS", "DICIONÁRIO"]
                st.session_state.df_livro = df_itens[df_itens[col_tipo].isin(tipos_livro)]
                
                st.session_state.df_material = pd.DataFrame() 
                st.session_state.df_vale = pd.DataFrame()
                st.session_state.df_integral = pd.DataFrame()
                st.session_state.df_bilingue = pd.DataFrame()
                
                st.session_state.df_livro, _, _, _ = configurar_e_calcular_tabela(st.session_state.df_livro, base_de_dados)
            else:
                st.session_state.df_material = pd.DataFrame(); st.session_state.df_vale = pd.DataFrame(); st.session_state.df_livro = pd.DataFrame(); st.session_state.df_integral = pd.DataFrame(); st.session_state.df_bilingue = pd.DataFrame()
            
            st.session_state.escola_anterior = escola_selecionada
            st.session_state.aba_anterior = aba_selecionada
            st.rerun()
    
    elif not escola_selecionada:
        pode_carregar = True 
        escola_final = "Pedido de Livro" 
        serie_final = "Avulso"

elif st.session_state.orcamento_mode == "Buscador Itens":
    st.header("🔎 Busca Global de Produtos")
    st.info("Esta ferramenta varre todas as planilhas de escola no Google Drive para encontrar onde cada item é usado.")
    
    if base_de_dados is None:
        st.error("A Base de Dados (do Drive) não pôde ser carregada. A busca não pode funcionar.")
    else:
        if st.button("Iniciar Busca Global (Pode ser lento)"):
            with st.spinner("Lendo todas as planilhas do Google Drive... (Isso pode levar alguns minutos)"):
                df_global_produtos = build_full_database(drive_service, base_de_dados)
            st.success(f"Banco de dados carregado! {len(df_global_produtos)} itens encontrados em todas as escolas.")
            st.session_state.df_global_produtos = df_global_produtos
        
        if "df_global_produtos" in st.session_state:
            search_term = st.text_input("Digite o nome do produto ou COD (ex: TINTA PVA ou 80023):")
            
            if search_term:
                search_term_upper = normalizar_texto(search_term)
                
                df_filtrado = st.session_state.df_global_produtos[
                    (st.session_state.df_global_produtos['Descrição'].str.contains(search_term_upper, case=False, na=False)) |
                    (st.session_state.df_global_produtos['COD'] == search_term_upper)
                ]
                
                if df_filtrado.empty:
                    st.warning("Nenhum item encontrado com esse nome ou código.")
                else:
                    st.markdown("---")
                    st.subheader(f"Resultados para: '{search_term}'")
                    st.dataframe(df_filtrado[['Escola', 'Série', 'Descrição', 'QTD', 'TIPO']])
                    
                    st.subheader("Resumo de Quantidade Total por Escola")
                    df_resumo = df_filtrado.groupby('Escola')['QTD'].sum().reset_index().rename(columns={'QTD': 'QTD Total'})
                    st.dataframe(df_resumo)
    
    pode_carregar = False 

elif st.session_state.orcamento_mode == "Atualizador PDF":
    st.header("🚀 Atualização em Lote para o Google Drive")
    st.warning("Atenção: Este processo irá ler todas as planilhas de escola no Google Drive, gerar novos PDFs, arquivar os antigos no Google Drive e atualizar sua Planilha Mestre. Isso pode levar alguns minutos.", icon="⚠️")
    
    if not PLANILHA_MESTRE_ID or "COLE_O_ID" in PLANILHA_MESTRE_ID:
        st.error("Erro de Configuração: O ID da Planilha Mestre (`PLANILHA_MESTRE_ID`) não foi definido no topo do script.")
    # ... (outras verificações de ID) ...
    elif base_de_dados is None:
        st.error("A Base de Dados (do Drive) não pôde ser carregada. Verifique o arquivo.")
    else:
        if st.button("INICIAR ATUALIZAÇÃO EM LOTE", type="primary"):
            st.info("Iniciando... O processo rodará no servidor.")
            with st.spinner("Autenticando e iniciando o processo... (Isso pode demorar)"):
                try:
                    # Serviços já devem estar autenticados
                    if drive_service and sheets_service:
                        st.success("Autenticação com Google bem-sucedida!")
                        run_batch_update(base_de_dados, sheets_service, drive_service)
                    else:
                        st.error("Falha ao inicializar os serviços do Google.")
                except Exception as e:
                    st.error(f"Falha na autenticação do Google. Erro: {e}")
    
    pode_carregar = False 
# --- FIM DOS MODOS ---


# --- O RESTO DO APP SÓ RODA NOS MODOS 1, 2, 4 e 5 ---
if base_de_dados is not None and pode_carregar:
    
    nome_cliente = None
    
    if st.session_state.orcamento_mode == "Gerador de Vale":
        st.subheader("Informações do Aluno (Vale)")
        col1, col2, col3 = st.columns(3)
        with col1:
            st.text_input("Nome do Responsável:", key="vale_responsavel")
        with col2:
            st.text_input("Nome do Aluno:", key="vale_aluno")
        with col3:
            st.text_input("Telefone:", key="vale_telefone", placeholder="(xx) xxxxx.xxxx", max_chars=15)
        nome_cliente = None 
    
    elif st.session_state.orcamento_mode == "Pedido de Livro":
        # Informações já foram pedidas acima
        nome_cliente = None 
        
    else:
        st.subheader("Informações do Cliente (Orçamento)")
        nome_cliente = st.text_input("Nome do Cliente:", placeholder="Digite o nome do cliente...", key="nome_cliente")
    
    st.markdown("---")
    
    with st.expander("🔎 **Assistente de Busca / Adicionar Item Rápido**", expanded=True):
        item_selecionado = st.selectbox("Comece a digitar o nome ou código do item:", options=lista_para_busca, index=None, placeholder="Digite aqui para buscar...", key="busca_item")
        cod_para_copiar = ""
        if item_selecionado:
            cod_para_copiar = item_selecionado.split(']')[0].replace('[', '')
        st.session_state.codigo_para_copiar = cod_para_copiar
        st.write("Código (clique no ícone 📋 no canto para copiar):")
        st.code(cod_para_copiar, language=None) 
        st.markdown("---")
        st.markdown("**Para ADICIONAR o item selecionado ao final da lista:**")
        
        if st.session_state.orcamento_mode == "Gerador de Vale":
            col1, col2 = st.columns([1, 2], vertical_alignment="bottom")
            with col1:
                st.number_input("Quantidade:", min_value=1, value=1, step=1, key="qtd_adicionar")
            with col2:
                st.button("Adicionar Item", type="primary", on_click=set_add_flag, use_container_width=True)
        
        elif st.session_state.orcamento_mode == "Pedido de Livro":
            col1, col2 = st.columns([1, 2], vertical_alignment="bottom")
            with col1:
                st.number_input("Quantidade:", min_value=1, value=1, step=1, key="qtd_adicionar")
            with col2:
                st.button("Adicionar Item", type="primary", on_click=set_add_flag, use_container_width=True)
        
        else: 
            col1, col2, col3 = st.columns([1, 2, 2], vertical_alignment="bottom") 
            with col1:
                st.number_input("Quantidade:", min_value=1, value=1, step=1, key="qtd_adicionar")
            with col2:
                st.radio("Adicionar em:", ("Material", "Vale", "Livro", "Integral", "Bilingue"), key="tipo_adicionar", horizontal=True, label_visibility="collapsed")
            with col3:
                st.button("Adicionar Item", type="primary", on_click=set_add_flag, use_container_width=True)

    if 'item_para_adicionar' in st.session_state and st.session_state.item_para_adicionar:
        flag_val = st.session_state.item_para_adicionar
        st.session_state.item_para_adicionar = None
        if flag_val == "ERRO":
            st.error("Por favor, selecione um item na busca primeiro.")
        elif isinstance(flag_val, dict):
            novo_df = pd.DataFrame([flag_val])
            
            tipo_adicionado = flag_val["TIPO"].upper() 
            
            if tipo_adicionado == "VALE":
                st.session_state.df_vale = pd.concat([st.session_state.df_vale, novo_df], ignore_index=True)
            elif tipo_adicionado == "LIVRO":
                st.session_state.df_livro = pd.concat([st.session_state.df_livro, novo_df], ignore_index=True)
            elif tipo_adicionado == "INTEGRAL":
                st.session_state.df_integral = pd.concat([st.session_state.df_integral, novo_df], ignore_index=True)
            elif tipo_adicionado == "BILINGUE":
                st.session_state.df_bilingue = pd.concat([st.session_state.df_bilingue, novo_df], ignore_index=True)
            else: 
                st.session_state.df_material = pd.concat([st.session_state.df_material, novo_df], ignore_index=True)
                
            st.success(f"Item '{flag_val['ITEM_STR']}' adicionado!")
            st.rerun()

    # ---- SEÇÃO 1: MATERIAL INDIVIDUAL ----
    if st.session_state.orcamento_mode not in ["Gerador de Vale", "Pedido de Livro"]:
        st.subheader("🛒 Itens de Material Individual")
        total_material = 0
        
        st.session_state.df_material, config_mat, total_material, ordem_mat = configurar_e_calcular_tabela(st.session_state.df_material, base_de_dados)
        df_material_editado = st.data_editor(st.session_state.df_material, num_rows="dynamic", key="editor_material", use_container_width=True, column_config=config_mat, column_order=ordem_mat)
        
        if not df_material_editado.equals(st.session_state.df_material):
            st.session_state.df_material = df_material_editado
            st.rerun()
            
        st.markdown(f"### Subtotal do Material: <span style='color: {COR_PRINCIPAL_PAPELARIA};'>R$ {total_material:,.2f}</span>", unsafe_allow_html=True)
        st.markdown("---") 
    else:
        total_material = 0

    # ---- SEÇÃO 2: VALE COLETIVO ----
    if st.session_state.orcamento_mode not in ["Pedido de Livro"]:
        st.subheader("🎁 Vale de Material Coletivo")
        total_vale = 0
        
        st.session_state.df_vale, config_vale, total_vale, ordem_vale = configurar_e_calcular_tabela(st.session_state.df_vale, base_de_dados)
        df_vale_editado = st.data_editor(st.session_state.df_vale, num_rows="dynamic", key="editor_vale", use_container_width=True, column_config=config_vale, column_order=ordem_vale)
        
        if not df_vale_editado.equals(st.session_state.df_vale):
            st.session_state.df_vale = df_vale_editado
            st.rerun()
            
        st.markdown(f"### Subtotal do Vale Coletivo: <span style='color: #FF4B4B;'>R$ {total_vale:,.2f}</span>", unsafe_allow_html=True)
        st.markdown("---")
    else:
        total_vale = 0
    
    # ---- SEÇÃO 3: LIVROS ----
    if st.session_state.orcamento_mode != "Gerador de Vale":
        if st.session_state.orcamento_mode == "Pedido de Livro":
            st.subheader("📚 Itens do Pedido de Livro")
        else:
            st.subheader("📚 Livros sob Encomenda")
            
        total_livro = 0
        
        st.session_state.df_livro, config_livro, total_livro, ordem_livro = configurar_e_calcular_tabela(st.session_state.df_livro, base_de_dados)
        
        df_livro_editado = st.data_editor(st.session_state.df_livro, num_rows="dynamic", key="editor_livro", use_container_width=True, column_config=config_livro, column_order=ordem_livro)
        
        if not df_livro_editado.equals(st.session_state.df_livro):
            st.session_state.df_livro = df_livro_editado
            st.rerun()
            
        st.markdown(f"### Subtotal dos Livros: <span style='color: {COR_PRINCIPAL_PAPELARIA};'>R$ {total_livro:,.2f}</span>", unsafe_allow_html=True)
        st.markdown("---")
    else:
        total_livro = 0 

    # --- SEÇÃO 4: INTEGRAL ---
    if st.session_state.orcamento_mode not in ["Gerador de Vale", "Pedido de Livro"]:
        st.subheader("🎨 Itens do Período Integral")
        total_integral = 0
        
        st.session_state.df_integral, config_integral, total_integral, ordem_integral = configurar_e_calcular_tabela(st.session_state.df_integral, base_de_dados)
        
        df_integral_editado = st.data_editor(st.session_state.df_integral, num_rows="dynamic", key="editor_integral", use_container_width=True, column_config=config_integral, column_order=ordem_integral)
        
        if not df_integral_editado.equals(st.session_state.df_integral):
            st.session_state.df_integral = df_integral_editado
            st.rerun()
            
        st.markdown(f"### Subtotal do Integral: <span style='color: {COR_PRINCIPAL_PAPELARIA};'>R$ {total_integral:,.2f}</span>", unsafe_allow_html=True)
        st.markdown("---")
    else:
        total_integral = 0 

    # --- SEÇÃO 5: BILINGUE ---
    if st.session_state.orcamento_mode not in ["Gerador de Vale", "Pedido de Livro"]:
        st.subheader("🌎 Itens do Programa Bilíngue")
        total_bilingue = 0
        
        st.session_state.df_bilingue, config_bilingue, total_bilingue, ordem_bilingue = configurar_e_calcular_tabela(st.session_state.df_bilingue, base_de_dados)
        
        df_bilingue_editado = st.data_editor(st.session_state.df_bilingue, num_rows="dynamic", key="editor_bilingue", use_container_width=True, column_config=config_bilingue, column_order=ordem_bilingue)
        
        if not df_bilingue_editado.equals(st.session_state.df_bilingue):
            st.session_state.df_bilingue = df_bilingue_editado
            st.rerun()
            
        st.markdown(f"### Subtotal do Bilíngue: <span style='color: {COR_PRINCIPAL_PAPELARIA};'>R$ {total_bilingue:,.2f}</span>", unsafe_allow_html=True)
        st.markdown("---")
    else:
        total_bilingue = 0 
    
    
    # ---- Seção de Total e Observações (Não mostrar nos Modos Vale/Livro) ---
    if st.session_state.orcamento_mode not in ["Gerador de Vale", "Pedido de Livro"]:
        valor_total_orcamento = total_material + total_vale + total_livro + total_integral + total_bilingue
        st.markdown(f"## Valor Total do Orçamento: <span style='color: green;'>R$ {valor_total_orcamento:,.2f}</span>", unsafe_allow_html=True)
        
        st.markdown("---")
        data_validade = extrair_data_validade(drive_service, PASTA_DRIVE_PLANILHAS_ID)
        st.markdown(f"**📅 Data de validade:** {data_validade}")

        st.subheader("Observações do Orçamento")
        col1, col2 = st.columns(2)
        with col1:
            st.text_area("🔴 Não trabalhamos", key="obs_nao_trabalhamos", height=150)
        with col2:
            st.text_area("🟡 Para escolher", key="obs_para_escolher", height=150)
        st.text_area("📝 Outras Observações (Novo)", key="obs_outras", height=100)
            
        st.markdown("---")
        
        # --- RASCUNHOS DESABILITADOS NA NUVEM ---
        if st.session_state.orcamento_mode in ("Novo Orçamento", "Orçamento Escola Pronto"):
            st.header("💾 Salvar Rascunho")
            st.warning("Salvar rascunhos no servidor não é suportado nesta versão.")
            st.markdown("---")
        # --- FIM DA MUDANÇA ---

        # --- MUDANÇA (BOTÃO DE DOWNLOAD) ---
        st.header("🖨️ Gerar PDF para Download")
        if nome_cliente:
            if st.button("Gerar Orçamento em PDF", type="primary"):
                
                df_mat_final = st.session_state.df_material
                df_vale_final = st.session_state.df_vale
                df_livro_final = st.session_state.df_livro 
                df_integral_final = st.session_state.df_integral
                df_bilingue_final = st.session_state.df_bilingue
                
                tem_zero = False
                col_valor_unit_padrao = "VALOR UNITARIO"
                
                def checar_zeros(df, col_nome_padrao):
                    if not df.empty:
                        col_lista = [c for c in df.columns if "UNIT" in c or ("VALOR" in c and "TOTAL" not in c) or "PRECO" in c]
                        col_real = col_lista[0] if col_lista else col_nome_padrao
                        # [CORREÇÃO PREÇO LIVRO] Não validar zeros para livros, integral ou bilingue
                        if col_real in df.columns and (df[col_real] == 0).any():
                            # Se a coluna TIPO existir, verificamos se é um item especial
                            if "TIPO" in df.columns:
                                tipos_especiais = ["LIVRO", "DICIONARIO", "LIVROS", "DICIONÁRIO", "INTEGRAL", "BILINGUE", "BILINGÜE"]
                                if df[df[col_real] == 0]['TIPO'].str.upper().isin(tipos_especiais).all():
                                    return False # É um livro/especial com preço 0, está OK
                            return True # É material normal com preço 0
                    return False
                
                if checar_zeros(df_mat_final, col_valor_unit_padrao) or \
                   checar_zeros(df_vale_final, col_valor_unit_padrao):
                   # Não checamos mais livro, integral ou bilingue
                    tem_zero = True
                    
                if tem_zero:
                    st.error("Erro: Um ou mais itens de MATERIAL ou VALE estão com valor R$ 0,00. Corrija os códigos ou a planilha base antes de gerar o PDF.")
                else:
                    with st.spinner("Gerando seu PDF, aguarde..."):
                        if st.session_state.orcamento_mode == "Orçamento Escola Pronto":
                            nome_escola_pdf = escola_final.split('_')[0]
                        else:
                            nome_escola_pdf = escola_final

                        logo_base64 = converter_imagem_base64(LOGO_PATH)
                        
                        totais = {
                            "material": total_material, "vale": total_vale, "livro": total_livro, 
                            "integral": total_integral, "bilingue": total_bilingue, "geral": valor_total_orcamento
                        }
                        
                        html_string = gerar_html_para_pdf(
                            logo_base64, nome_escola_pdf, serie_final, nome_cliente, data_validade, 
                            df_mat_final, df_vale_final, df_livro_final, df_integral_final, df_bilingue_final, 
                            st.session_state.obs_nao_trabalhamos, 
                            st.session_state.obs_para_escolher, st.session_state.obs_outras, totais
                        )
                        
                        pdf_bytes = converter_html_para_pdf(html_string)
                        
                        if pdf_bytes:
                            nome_cliente_sanitizado = sanitizar_nome_arquivo(nome_cliente)
                            nome_arquivo = f"Orcamento {nome_escola_pdf} {serie_final} - {nome_cliente_sanitizado}.pdf"
                            
                            st.success(f"✅ Orçamento gerado com sucesso!")
                            st.download_button(
                                label="Clique aqui para baixar o PDF",
                                data=pdf_bytes,
                                file_name=nome_arquivo,
                                mime="application/pdf"
                            )
                        else:
                            st.error("Não foi possível gerar o PDF.")
        else:
            st.warning("Por favor, digite o nome do cliente acima para poder gerar o PDF.")

    # --- BOTÕES (VALE E LIVRO) ---
    else: 
        if st.session_state.orcamento_mode == "Gerador de Vale":
            st.header("🖨️ Gerar PDF do Vale")
            
            if st.session_state.vale_aluno and st.session_state.vale_responsavel and st.session_state.vale_telefone:
                if st.button("Gerar Vale em PDF", type="primary"):

                    df_vale_final = st.session_state.df_vale
                    total_vale_final = total_vale 
                    
                    tem_zero = False
                    col_valor_unit_padrao = "VALOR UNITARIO"
                    if not df_vale_final.empty:
                        col_valor_unit_vale_lista = [c for c in df_vale_final.columns if "UNIT" in c or ("VALOR" in c and "TOTAL" not in c) or "PRECO" in c]
                        col_valor_unit_vale = col_valor_unit_vale_lista[0] if col_valor_unit_vale_lista else col_valor_unit_padrao
                        
                        if col_valor_unit_vale in df_vale_final.columns and (df_vale_final[col_valor_unit_vale] == 0).any():
                            tem_zero = True
                    
                    if tem_zero:
                        st.error("Erro: Um ou mais itens estão com valor R$ 0,00. Corrija os códigos ou a planilha base antes de gerar o PDF.")
                    else:
                        with st.spinner("Gerando PDF do Vale..."):
                            nome_escola_pdf = escola_final.split('_')[0]
                            telefone_formatado = formatar_telefone(st.session_state.vale_telefone)
                            
                            try:
                                pdf_bytes = gerar_vale_pdf_reportlab(
                                    LOGO_PATH, 
                                    nome_escola_pdf, 
                                    serie_final,
                                    st.session_state.vale_aluno, 
                                    st.session_state.vale_responsavel, 
                                    telefone_formatado, 
                                    df_vale_final,
                                    total_vale_final
                                )
                                
                                nome_aluno_sanitizado = sanitizar_nome_arquivo(st.session_state.vale_aluno)
                                nome_arquivo = f"Vale {nome_escola_pdf} {serie_final} - {nome_aluno_sanitizado}.pdf"
                                
                                st.success(f"✅ Vale gerado com sucesso!")
                                st.download_button(
                                    label="Clique aqui para baixar o PDF",
                                    data=pdf_bytes,
                                    file_name=nome_arquivo,
                                    mime="application/pdf"
                                )
                            except Exception as e:
                                st.error(f"Erro ao gerar PDF do Vale com ReportLab: {e}")
                                st.exception(e) 
            else:
                st.warning("Por favor, preencha o Nome do Aluno, Responsável e Telefone para salvar o Vale.")
        
        elif st.session_state.orcamento_mode == "Pedido de Livro":
            st.header("🖨️ Gerar Pedido de Livro (2 Vias)")
            
            if st.session_state.livro_cliente and st.session_state.livro_telefone:
                if st.button("Gerar Pedido (2 vias) em PDF", type="primary"):
                    
                    df_livro_final = st.session_state.df_livro
                    
                    with st.spinner("Gerando PDF (2 vias)..."):
                        try:
                            pdf_bytes = gerar_pedido_livro_pdf_reportlab(
                                LOGO_PATH,
                                st.session_state.livro_cliente,
                                st.session_state.livro_telefone,
                                df_livro_final,
                                total_livro, 
                                st.session_state.livro_obs 
                            )
                            
                            nome_cliente_sanitizado = sanitizar_nome_arquivo(st.session_state.livro_cliente)
                            nome_arquivo = f"Pedido Livro - {nome_cliente_sanitizado}.pdf"
                            
                            st.success(f"✅ Pedido gerado com sucesso!")
                            st.download_button(
                                label="Clique aqui para baixar o PDF",
                                data=pdf_bytes,
                                file_name=nome_arquivo,
                                mime="application/pdf"
                            )
                        except Exception as e:
                            st.error(f"Erro ao gerar PDF do Pedido de Livro: {e}")
                            st.exception(e)
            else:
                st.warning("Por favor, preencha o Nome do Cliente e o Telefone para gerar o Pedido.")
        

elif not st.session_state.get("escola_selecionada_select") and st.session_state.orcamento_mode == "Orçamento Escola Pronto":
    st.info("Por favor, selecione uma escola para começar.")
elif st.session_state.orcamento_mode == "Carregar Rascunho":
    pass 
elif not (st.session_state.get("escola_manual") and st.session_state.get("serie_manual")) and st.session_state.orcamento_mode == "Novo Orçamento":
    st.info("Por favor, digite a escola e a série para começar.")
elif st.session_state.orcamento_mode == "Gerador de Vale":
    pass 
elif st.session_state.orcamento_mode == "Pedido de Livro":
    pass 
elif st.session_state.orcamento_mode == "Buscador Itens":
    pass 
elif st.session_state.orcamento_mode == "Atualizador PDF":
    pass 
elif base_de_dados is None:
    st.error("A base de dados (do Drive) não pôde ser carregada. O aplicativo não pode continuar.")


