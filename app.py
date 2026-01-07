import streamlit as st
import pandas as pd
import io
import difflib
import unicodedata

# Configuração da Página
st.set_page_config(page_title="Agente de Automação de Vendas", layout="wide")

st.title("🤖 Agente de Automação de Vendas (Excel -> Excel)")
st.markdown("""
Este agente processa uma **Planilha de Vendas**, cruza com a **Base de Vendedores** e gera um relatório consolidado.
""")

# --- Etapa 1: Base de Conhecimento ---
st.header("1. Base de Conhecimento (Mapeamento)")

uploaded_base = st.file_uploader("Upload da Base de Vendedores (Excel)", type=["xlsx", "xls"], key="base_uploader")

vendedor_map = {}
df_base = None

if uploaded_base:
    try:
        df_base = pd.read_excel(uploaded_base)
        
        # Estrutura esperada: Colunas pares "Vendedor" e ímpares "Cliente" (ou repetidas)
        # Ex: Vendedor | Cliente | Vendedor | Cliente ...
        
        pairs_found = 0
        
        # Vamos iterar pelas colunas e procurar pares
        cols = df_base.columns
        for i in range(len(cols)):
            col_name = str(cols[i]).strip()
            col_lower = col_name.lower()
            
            # Identifica coluna de Cliente
            if 'cliente' in col_lower:
                # Verifica se a coluna ANTERIOR é Vendedor
                if i > 0:
                    prev_col_name = str(cols[i-1]).strip()
                    prev_col_lower = prev_col_name.lower()
                    
                    if 'vended' in prev_col_lower:
                        # Achamos um par (Vendedor [i-1], Cliente [i])
                        
                        # Extrai dados desse par
                        # Normaliza
                        clientes_serie = df_base.iloc[:, i].astype(str).str.strip().str.upper()
                        vendedores_serie = df_base.iloc[:, i-1].astype(str).str.strip()
                        
                        # Filtra vazios (nan ou 'nan')
                        mask = (clientes_serie != 'NAN') & (clientes_serie != '')
                        
                        # Update no dicionário global
                        par_map = dict(zip(clientes_serie[mask], vendedores_serie[mask]))
                        vendedor_map.update(par_map)
                        
                        pairs_found += 1
        
        if pairs_found > 0:
             st.success(f"✅ Base carregada! {len(vendedor_map)} clientes mapeados a partir de {pairs_found} pares de colunas.")
             with st.expander("Ver Base Consolidada"):
                 # Mostra o dicionário como dataframe para conferencia
                 df_debug = pd.DataFrame(list(vendedor_map.items()), columns=['Cliente', 'Vendedor'])
                 st.dataframe(df_debug)
        else:
            st.error("❌ Não encontrei pares de colunas 'Vendedor' e 'Cliente' lado a lado (ex: Col A=Vendedor, Col B=Cliente).")
            # Mostra as colunas lidas para ajudar no debug
            st.write("Colunas identificadas no arquivo:", cols.tolist())
            vendedor_map = {} # Reset
            
    except Exception as e:
        st.error(f"Erro ao ler planilha de base: {e}")

# --- Helper: Normalização ---
def normalize_string(s):
    """Remove acentos e coloca em maiúsculo para comparação"""
    if not isinstance(s, str):
        s = str(s)
    return "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn").upper().strip()

# --- Etapa 2: Processamento de Vendas ---
st.header("2. Processamento de Vendas (Excel)")

uploaded_sales = st.file_uploader("Upload da Planilha de Vendas (Input)", type=["xlsx", "xls"], key="sales_uploader", accept_multiple_files=False)

if uploaded_sales:
    if not vendedor_map:
        st.warning("⚠️ Por favor, faça o upload da Base de Vendedores na Etapa 1 primeiro.")
    else:
        try:
            df_vendas = pd.read_excel(uploaded_sales)
            st.info(f"📄 Planilha de Vendas carregada com {len(df_vendas)} linhas.")
            
            # Identificar colunas alvo na planilha de vendas
            col_data = None
            col_cliente_venda = None
            col_valor = None
            
            # Procura colunas por palavras-chave ou nomes exatos
            # Usuario pediu: "Data Aprovação", "Clientes", "Valor Total"
            
            for col in df_vendas.columns:
                c_clean = col.strip()
                c_lower = c_clean.lower()
                
                # Identificação inteligente
                if c_clean == "Data Aprovação" or ('data' in c_lower and 'aprov' in c_lower):
                    col_data = col
                elif c_clean == "Clientes" or c_clean == "Cliente" or ('cliente' in c_lower):
                    col_cliente_venda = col
                elif c_clean == "Valor Total" or ('valor' in c_lower and 'total' in c_lower):
                    col_valor = col
            
            missing_cols = []
            if not col_data: missing_cols.append("Data Aprovação")
            if not col_cliente_venda: missing_cols.append("Clientes")
            if not col_valor: missing_cols.append("Valor Total")
            
            if missing_cols:
                st.error(f"❌ Não encontrei as colunas esperadas: {', '.join(missing_cols)}. Verifique os nomes na planilha.")
                st.write("Colunas encontradas:", df_vendas.columns.tolist())
            else:
                # --- Processamento de Cruzamento ---
                st.write("Processando cruzamento de dados...")
                
                vendedores_encontrados = []
                logs_processamento = []
                
                clientes_keys_upper = [k.upper() for k in vendedor_map.keys()]
                
                for idx, row in df_vendas.iterrows():
                    cliente_input = row[col_cliente_venda]
                    
                    if pd.isna(cliente_input):
                        vendedores_encontrados.append("N/A")
                        continue
                        
                    cliente_input_norm = normalize_string(cliente_input)
                    vendedor_match = None
                    
                    # 1. Match Exato (Normalizado)
                    if cliente_input_norm in vendedor_map:
                         vendedor_match = vendedor_map[cliente_input_norm]
                    
                    # 2. Fuzzy Match
                    if not vendedor_match:
                        # Tenta achar o mais proximo
                        matches = difflib.get_close_matches(cliente_input_norm, clientes_keys_upper, n=1, cutoff=0.7)
                        if matches:
                            match_upper = matches[0]
                            # Como normalizamos as chaves do map, elas estão em upper
                            if match_upper in vendedor_map:
                                vendedor_match = vendedor_map[match_upper]
                    
                    if vendedor_match:
                        vendedores_encontrados.append(vendedor_match)
                    else:
                        vendedores_encontrados.append("Não Encontrado")
                        # logs_processamento.append(f"Linha {idx+2}: Cliente '{cliente_input}' não encontrado.")
                
                # Adiciona Coluna no DataFrame
                df_vendas['Vendedor'] = vendedores_encontrados
                
                # Reordenar: Data, Cliente, Valor, Vendedor...
                cols_priority = [col_data, col_cliente_venda, col_valor, 'Vendedor']
                other_cols = [c for c in df_vendas.columns if c not in cols_priority]
                df_final = df_vendas[cols_priority + other_cols]
                
                st.success("✅ Processamento Concluído!")
                
                st.subheader("Prévia do Relatório")
                st.dataframe(df_final)
                
                # Download
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_final.to_excel(writer, index=False, sheet_name='Consolidado')
                
                st.download_button(
                    label="💾 Baixar Relatório Consolidado",
                    data=output.getvalue(),
                    file_name="Relatorio_Vendas_Consolidado.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        except Exception as e:
            st.error(f"Erro ao processar planilha de vendas: {e}")
