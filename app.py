import streamlit as st
import pandas as pd
import pyodbc
from datetime import datetime
from io import BytesIO

# ==============================================================================
# 1. CONFIGURAÇÕES DA PÁGINA
# ==============================================================================
st.set_page_config(
    page_title="Conferência Fiscal",
    layout="wide",
    initial_sidebar_state="collapsed",
    page_icon="📊"
)

# ==============================================================================
# 2. FUNÇÕES (Lógica de Negócio)
# ==============================================================================

def to_excel(df):
    """Converte DataFrame para Excel em memória usando openpyxl."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Relatorio')
    processed_data = output.getvalue()
    return processed_data

def extrair_numero_nota_sql(valor):
    try:
        s_val = str(valor).strip()
        if len(s_val) > 4 and s_val.isdigit():
            return int(s_val[4:])
        return int(s_val)
    except:
        return 0

def extrair_numero_nota_excel(valor):
    try:
        if pd.isna(valor): return 0
        s_val = str(valor).strip()
        if '/' in s_val:
            return int(s_val.split('/')[1])
        return int(float(s_val))
    except:
        return 0

def verificar_cancelamento_excel(valor):
    if pd.isna(valor): return False
    return "CANCEL" in str(valor).upper()

@st.cache_data(show_spinner=False, ttl=300)
def buscar_dados_sql(data_inicio, data_fim, empresa_id, puxar_todas):
    try:
        db = st.secrets["uau_db"]
        raw_query = st.secrets["sql_queries"]["query_conferencia"]
        conn_str = f"DRIVER={db['DRIVER']};SERVER={db['SERVER']};DATABASE={db['DATABASE']};UID={db['UID']};PWD={db['PWD']}"
    except Exception as e:
        st.error("Erro ao carregar configurações de segurança (Secrets).")
        return pd.DataFrame()

    d_ini = data_inicio.strftime('%Y-%m-%d')
    d_fim = data_fim.strftime('%Y-%m-%d')

    if puxar_todas:
        filtro_empresa_nf = "1=1"
        filtro_empresa_end = "1=1"
    else:
        filtro_empresa_nf = f"NotasFiscais.Empresa_nf = {empresa_id}"
        filtro_empresa_end = f"NotaFiscalEndereco.Empresa_NfEnd = {empresa_id}"

    try:
        query = raw_query.format(
            d_ini=d_ini, d_fim=d_fim,
            filtro_empresa_nf=filtro_empresa_nf,
            filtro_empresa_end=filtro_empresa_end
        )
    except Exception as e:
        st.error(f"Erro ao formatar query: {e}")
        return pd.DataFrame()
    
    try:
        conn = pyodbc.connect(conn_str)
        df = pd.read_sql(query, conn)
        conn.close()
        return df
    except Exception as e:
        st.error(f"❌ Erro SQL: {e}")
        return pd.DataFrame()

# ==============================================================================
# 3. LAYOUT E INTERFACE
# ==============================================================================

st.markdown("<h1 style='text-align: center; color: #FFFFFF;'>Conferência de Notas Fiscais 2025</h1>", unsafe_allow_html=True)
st.markdown("<p style='text-align: center; color: #BDC3C7; margin-bottom: 30px;'>Sistema automático de comparação: <b>UAU</b> vs <b>Planilha de Controle</b>.</p>", unsafe_allow_html=True)

# Container 1: Filtros
with st.container(border=True):
    st.markdown("##### 🛠️ Configuração da Busca")
    col_dates, col_company = st.columns([2, 1])
    with col_dates:
        c1, c2 = st.columns(2)
        with c1: dt_inicio = st.date_input("📅 Data Inicial", value=datetime(2025, 1, 1))
        with c2: dt_fim = st.date_input("📅 Data Final", value=datetime.today())
    with col_company:
        empresa_id = st.number_input("🏢 Cód. Empresa", min_value=1, value=1)
        todas_empresas = st.checkbox("Todas Empresas", help="Ignora o código")

# Container 2: Upload
with st.container(border=True):
    st.markdown("##### 📂 Arquivo de Comparação")
    uploaded_file = st.file_uploader("Arraste sua planilha aqui", type=["xlsx", "csv"])

# Botão de Ação (Reseta o estado se clicado novamente)
st.write("")
col_vazia_esq, col_btn, col_vazia_dir = st.columns([1, 1, 1])
with col_btn:
    btn_run = st.button("🚀 INICIAR CONFERÊNCIA", type="primary", use_container_width=True)

# ==============================================================================
# 4. PROCESSAMENTO (COM SESSION STATE)
# ==============================================================================

# Se o usuário clicar no botão, rodamos o processamento e salvamos no session_state
if btn_run:
    if not uploaded_file:
        st.warning("⚠️ Por favor, anexe a planilha antes de clicar no botão.")
    else:
        status_bar = st.status("🔍 Iniciando conferência...", expanded=True)
        
        # 1. Busca SQL
        status_bar.write("📡 Conectando ao Sistema UAU...")
        df_sql = buscar_dados_sql(dt_inicio, dt_fim, empresa_id, todas_empresas)
        
        # 2. Leitura Excel
        status_bar.write("📊 Lendo planilha anexada...")
        try:
            if uploaded_file.name.endswith('.csv'):
                df_excel = pd.read_csv(uploaded_file, header=2)
            else:
                df_excel = pd.read_excel(uploaded_file, header=2)
            
            df_excel.columns = [str(c).strip() for c in df_excel.columns]
            
            if 'DATA NF' in df_excel.columns:
                df_excel['DATA NF'] = pd.to_datetime(df_excel['DATA NF'], errors='coerce')
                df_excel = df_excel[
                    (df_excel['DATA NF'].dt.date >= dt_inicio) & 
                    (df_excel['DATA NF'].dt.date <= dt_fim)
                ]
        except Exception as e:
            status_bar.update(label="Erro na leitura do arquivo", state="error")
            st.error(f"Erro: {e}")
            st.stop()

        status_bar.write("⚙️ Cruzando informações...")
        
        # 3. Tratamento
        if not df_sql.empty:
            df_sql['CHAVE'] = df_sql['NumNfAux_nf'].apply(extrair_numero_nota_sql)
            df_sql['CANCELADO_SISTEMA'] = df_sql['Status_nf'] == 1
        else:
            df_sql = pd.DataFrame(columns=['CHAVE', 'NumNfAux_nf', 'Status_nf', 'CANCELADO_SISTEMA'])

        if 'Nº NF' not in df_excel.columns:
            status_bar.update(label="Erro: Coluna não encontrada", state="error")
            st.error("A coluna 'Nº NF' não existe na planilha.")
            st.stop()
            
        df_excel['CHAVE'] = df_excel['Nº NF'].apply(extrair_numero_nota_excel)
        col_receb = 'DATA DE RECEBIMENTO'
        df_excel['CANCELADO_PLANILHA'] = df_excel[col_receb].apply(verificar_cancelamento_excel) if col_receb in df_excel.columns else False

        # 4. Comparação
        chaves_sistema = set(df_sql['CHAVE']) if not df_sql.empty else set()
        chaves_planilha = set(df_excel['CHAVE'])
        
        so_no_sistema = chaves_sistema - chaves_planilha
        so_na_planilha = chaves_planilha - chaves_sistema
        
        df_comum = pd.merge(
            df_sql[['CHAVE', 'NumNfAux_nf', 'CANCELADO_SISTEMA', 'Nome_pes', 'ValorTotNota_nf']], 
            df_excel[['CHAVE', 'Nº NF', 'CANCELADO_PLANILHA', 'VALOR NF']], 
            on='CHAVE', how='inner'
        )
        df_status_errado = df_comum[df_comum['CANCELADO_SISTEMA'] != df_comum['CANCELADO_PLANILHA']]
        
        # SALVANDO NO SESSION STATE PARA NÃO PERDER AO CLICAR NO DOWNLOAD
        st.session_state['resultado_pronto'] = True
        st.session_state['df_status_errado'] = df_status_errado
        
        # Preparando DataFrames das faltas para salvar também
        st.session_state['df_falta_uau'] = df_excel[df_excel['CHAVE'].isin(so_na_planilha)] if len(so_na_planilha) > 0 else pd.DataFrame()
        st.session_state['df_falta_planilha'] = df_sql[df_sql['CHAVE'].isin(so_no_sistema)] if len(so_no_sistema) > 0 else pd.DataFrame()
        
        # Salvando contagens
        st.session_state['count_sistema'] = len(so_no_sistema)
        st.session_state['count_planilha'] = len(so_na_planilha)
        st.session_state['count_errado'] = len(df_status_errado)

        status_bar.update(label="Concluído com sucesso!", state="complete", expanded=False)

# ==============================================================================
# 5. DASHBOARD (Lê do Session State)
# ==============================================================================

# Verificamos se existe resultado pronto na memória
if st.session_state.get('resultado_pronto'):
    
    st.markdown("### 📊 Resultado da Análise")
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.metric("Faltam na Planilha", f"{st.session_state['count_sistema']} notas", 
                  delta="Estão no UAU mas não no Excel", delta_color="inverse")
    with col2:
        st.metric("Faltam no Sistema", f"{st.session_state['count_planilha']} notas", 
                  delta="Estão no Excel mas não no UAU", delta_color="inverse")
    with col3:
        cor_status = "normal" if st.session_state['count_errado'] == 0 else "inverse"
        st.metric("Status Diferente", f"{st.session_state['count_errado']} notas", 
                  delta="Cancelado em um, Ativo no outro", delta_color=cor_status)

    st.markdown("---")
    
    tab1, tab2, tab3 = st.tabs(["📝 Relatório: Status Incorreto", "📂 Falta Lançar no UAU", "📉 Falta na Planilha"])
    
    # --- TAB 1 ---
    with tab1:
        df1 = st.session_state['df_status_errado']
        if not df1.empty:
            st.error("Atenção: Status divergentes.")
            c_info, c_dl = st.columns([4,1])
            with c_dl:
                st.download_button("📥 Baixar Excel", to_excel(df1), 'status_incorreto.xlsx', key='dl_1')
            st.dataframe(df1, use_container_width=True)
        else:
            st.success("✅ Tudo certo com os status.")

    # --- TAB 2 ---
    with tab2:
        df2 = st.session_state['df_falta_uau']
        if not df2.empty:
            st.warning("Notas na planilha, mas não no UAU.")
            c_info, c_dl = st.columns([4,1])
            with c_dl:
                st.download_button("📥 Baixar Excel", to_excel(df2), 'falta_lancar_no_uau.xlsx', key='dl_2')
            st.dataframe(df2, use_container_width=True)
        else:
            st.success("✅ Todas as notas da planilha estão no sistema.")

    # --- TAB 3 ---
    with tab3:
        df3 = st.session_state['df_falta_planilha']
        if not df3.empty:
            st.info("Notas no UAU, mas não na planilha.")
            c_info, c_dl = st.columns([4,1])
            with c_dl:
                st.download_button("📥 Baixar Excel", to_excel(df3), 'falta_na_planilha.xlsx', key='dl_3')
            st.dataframe(df3, use_container_width=True)
        else:
            st.success("✅ Não há notas sobrando no sistema.")