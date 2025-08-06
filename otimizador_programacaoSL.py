import streamlit as st
import pandas as pd
import re
from datetime import datetime
import io # Necessário para manipulação do arquivo em memória

# --- LÓGICA DE NEGÓCIO (Otimização) ---
# Esta seção permanece idêntica à versão original. É o núcleo do sistema.

DATA_ATUAL = datetime.now()
PESO_META_BATCH = 120

def parse_dimensoes(descricao_produto):
    if not isinstance(descricao_produto, str) or "REBAIXAD" in descricao_produto.upper():
        return None, None
    match = re.search(r'(\d+[\.,]\d+|\d+)\s*X\s*(\d+)', descricao_produto)
    if match:
        try:
            espessura = float(match.group(1).replace(',', '.'))
            largura = int(match.group(2))
            return espessura, largura
        except (ValueError, IndexError):
            return None, None
    return None, None

def determinar_setup(espessura, largura):
    if espessura is None or largura is None:
        return "SETUP_INDEFINIDO"
    plaina = "PLAINA_FINA" if espessura <= 4.75 else "PLAINA_GROSSA"
    balancim = f"{largura}mm"
    return f"{plaina}_{balancim}"

def calcular_urgencia(data_entrega, data_atual):
    if pd.isna(data_entrega):
        return 5, "5 - COM FOLGA (SEM DATA)"
    dias_atraso = (data_atual - data_entrega).days
    if dias_atraso > 10: return 1, "1 - URGENTÍSSIMO"
    if 5 <= dias_atraso <= 10: return 2, "2 - URGENTE"
    if 0 <= dias_atraso <= 4: return 3, "3 - ATRASADO"
    if -10 <= dias_atraso < 0: return 4, "4 - NO TEMPO"
    return 5, "5 - COM FOLGA"

def processar_dados_lote(df_original):
    lotes = []
    for lote_id, grupo in df_original.groupby('LOTE'):
        qtde_numerica = pd.to_numeric(grupo['QTDE'], errors='coerce').fillna(0)
        peso_total_lote = qtde_numerica.sum()
        
        data_entrega_lote = grupo['DATA DE ENTREGA'].min()
        espessura_lote, largura_lote = None, None
        for produto in grupo['PRODUTO']:
            espessura, largura = parse_dimensoes(produto)
            if espessura is not None:
                espessura_lote, largura_lote = espessura, largura
                break
        setup_lote = determinar_setup(espessura_lote, largura_lote)
        urgencia_num, urgencia_nome = calcular_urgencia(data_entrega_lote, DATA_ATUAL)
        lotes.append({
            'LOTE': lote_id, 'PESO_TOTAL': peso_total_lote, 'DATA_ENTREGA': data_entrega_lote,
            'SETUP': setup_lote, 'URGENCIA_NIVEL': urgencia_num, 'URGENCIA_NOME': urgencia_nome,
            'ESPESSURA': espessura_lote if espessura_lote is not None else 999
        })
    return pd.DataFrame(lotes)

def otimizar_sequencia(df_lotes, priorities):
    p1, p2, p3 = priorities
    df_lotes_sorted = df_lotes.sort_values(by=[p1, p2, p3, 'DATA_ENTREGA'])
    sequencia_otimizada_lotes = df_lotes_sorted.to_dict('records')
    
    batches_info = []
    if not sequencia_otimizada_lotes:
        return [], []

    current_batch_weight = 0
    current_batch_setup = sequencia_otimizada_lotes[0]['SETUP']
    
    for lote in sequencia_otimizada_lotes:
        if lote['SETUP'] != current_batch_setup:
            batches_info.append({
                'setup': current_batch_setup,
                'peso': current_batch_weight,
                'atingiu_meta': current_batch_weight >= PESO_META_BATCH
            })
            current_batch_setup = lote['SETUP']
            current_batch_weight = 0
        current_batch_weight += lote['PESO_TOTAL']

    batches_info.append({
        'setup': current_batch_setup,
        'peso': current_batch_weight,
        'atingiu_meta': current_batch_weight >= PESO_META_BATCH
    })
        
    return sequencia_otimizada_lotes, batches_info

def gerar_relatorio_final(sequencia_lotes, df_original):
    if not sequencia_lotes: return pd.DataFrame()
    df_sequencia = pd.DataFrame(sequencia_lotes)
    lotes_ordenados = df_sequencia['LOTE'].unique()
    mapa_posicao = {lote_id: i + 1 for i, lote_id in enumerate(lotes_ordenados)}
    mapa_info = df_sequencia.drop_duplicates(subset=['LOTE']).set_index('LOTE')
    
    df_final = df_original.copy()
    
    df_final['Posição na Sequência'] = df_final['LOTE'].map(mapa_posicao)
    df_final['SETUP'] = df_final['LOTE'].map(mapa_info['SETUP'])
    df_final['URGENCIA'] = df_final['LOTE'].map(mapa_info['URGENCIA_NOME'])
    
    df_final.dropna(subset=['Posição na Sequência'], inplace=True)
    df_final['Posição na Sequência'] = df_final['Posição na Sequência'].astype(int)
    df_final = df_final.sort_values(by='Posição na Sequência')

    df_final['DATA DE ENTREGA'] = df_final['DATA DE ENTREGA'].dt.strftime('%d/%m/%Y').fillna('')
    if 'DT PRODUÇÃO' in df_final.columns:
        df_final['DT PRODUÇÃO'] = pd.to_datetime(df_final['DT PRODUÇÃO'], errors='coerce').dt.strftime('%d/%m/%Y').fillna('')
    if 'PREVISÃO' in df_final.columns:
        df_final['PREVISÃO'] = pd.to_datetime(df_final['PREVISÃO'], errors='coerce').dt.strftime('%d/%m/%Y').fillna('')

    ordem_final_colunas = [
        'Posição na Sequência', 'LOTE', 'PC', 'PEDIDO', 'PRODUTO', 'QTDE', 
        'PREVISÃO', 'OBS.:', 'DT PRODUÇÃO', 'TURNO', 'PESO BOB', 'DATA DE ENTREGA',
        'SETUP', 'URGENCIA'
    ]
    
    for col in ordem_final_colunas:
        if col not in df_final.columns:
            df_final[col] = ''
    df_relatorio = df_final[ordem_final_colunas]
    
    return df_relatorio

def calcular_metricas(sequencia_lotes, batches_info):
    if not batches_info: return pd.DataFrame(), pd.DataFrame()
    total_mudancas_setup = len(batches_info) - 1 if len(batches_info) > 0 else 0
    num_batches = len(batches_info)
    batches_na_meta = sum(1 for b in batches_info if b['atingiu_meta'])
    perc_batches_na_meta = (batches_na_meta / num_batches * 100) if num_batches > 0 else 0
    peso_total_processado = sum(b['peso'] for b in batches_info)
    peso_medio_batch = peso_total_processado / num_batches if num_batches > 0 else 0
    
    df_sequencia = pd.DataFrame(sequencia_lotes)
    df_lotes_unicos = df_sequencia.drop_duplicates(subset=['LOTE'])
    dist_urgencia = df_lotes_unicos['URGENCIA_NOME'].value_counts().reset_index()
    dist_urgencia.columns = ['Nível de Urgência', 'Qtd. Lotes']
    
    metricas = {
        "Total de Lotes Processados": len(df_lotes_unicos),
        "Peso Total Processado (t)": f"{peso_total_processado:.2f}",
        "Total de Batches Criados": num_batches,
        "Total de Mudanças de Setup": total_mudancas_setup,
        "Batches que Atingiram a Meta de 120t": batches_na_meta,
        "% de Batches na Meta": f"{perc_batches_na_meta:.2f}%",
        "Peso Médio por Batch (t)": f"{peso_medio_batch:.2f}",
    }
    df_metricas = pd.DataFrame(list(metricas.items()), columns=['Métrica', 'Valor'])
    return df_metricas, dist_urgencia


# --- INTERFACE GRÁFICA (Streamlit) ---

st.set_page_config(layout="wide", page_title="Otimizador Fagor")

st.title("Otimizador de Programação de Corte Fagor")
st.markdown("---")

# Mapa de prioridades (Interface -> Nome da Coluna no DataFrame)
priority_map = {
    "Urgência (Data)": "URGENCIA_NIVEL",
    "Setup (Plaina/Largura)": "SETUP",
    "Espessura": "ESPESSURA"
}
priority_options = list(priority_map.keys())

# --- 1. Upload do Arquivo ---
st.header("1. Anexar Planilha")
uploaded_file = st.file_uploader("Selecione o arquivo Excel com os dados de produção (.xlsx)", type=["xlsx"])

if uploaded_file:
    st.success(f"Arquivo '{uploaded_file.name}' carregado com sucesso!")
    st.markdown("---")
    
    # --- 2. Definição de Prioridades ---
    st.header("2. Definir Ordem de Prioridade")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        p1_selection = st.selectbox("Prioridade 1:", priority_options, index=0)
    
    with col2:
        # Filtra as opções para não repetir a Prioridade 1
        options_p2 = [opt for opt in priority_options if opt != p1_selection]
        p2_selection = st.selectbox("Prioridade 2:", options_p2, index=0)
        
    with col3:
        # Filtra as opções para não repetir a P1 e P2
        options_p3 = [opt for opt in priority_options if opt not in [p1_selection, p2_selection]]
        p3_selection = st.selectbox("Prioridade 3:", options_p3, index=0)

    st.markdown("---")

    # --- 3. Execução ---
    st.header("3. Gerar Programação")
    
    if st.button("Gerar Programação Otimizada", type="primary"):
        with st.spinner("Processando... Por favor, aguarde. Isso pode levar alguns segundos."):
            try:
                # Obter prioridades selecionadas e mapeá-las para os nomes das colunas
                priorities_keys = [p1_selection, p2_selection, p3_selection]
                priorities = [priority_map[p] for p in priorities_keys]

                # Leitura e processamento dos dados (mesma lógica robusta)
                df = pd.read_excel(uploaded_file, sheet_name='Fagor', dtype=str)
                df.columns = [str(col).strip() for col in df.columns]

                essential_cols = ['LOTE', 'PRODUTO', 'QTDE']
                for col in essential_cols:
                    if col not in df.columns:
                        raise ValueError(f"A coluna essencial '{col}' não foi encontrada na planilha.")
                df.dropna(subset=essential_cols, inplace=True)
                
                date_col_name = 'DATA DE ENTREGA'
                if date_col_name in df.columns:
                    def robust_date_parser(val):
                        if pd.isna(val) or val == '': return pd.NaT
                        try:
                            excel_date_num = float(val)
                            return pd.to_datetime('1899-12-30') + pd.to_timedelta(excel_date_num, 'D')
                        except (ValueError, TypeError):
                            return pd.to_datetime(val, dayfirst=True, errors='coerce')
                    df[date_col_name] = df[date_col_name].apply(robust_date_parser)
                else:
                    df[date_col_name] = pd.NaT

                df['LOTE'] = df['LOTE'].astype(str)
                df['QTDE'] = pd.to_numeric(df['QTDE'], errors='coerce').fillna(0)
                
                # Execução da lógica de negócio
                df_lotes = processar_dados_lote(df)
                sequencia, batches = otimizar_sequencia(df_lotes, priorities)
                
                if not sequencia:
                    st.error("Nenhum lote válido para processar foi encontrado.")
                else:
                    df_relatorio = gerar_relatorio_final(sequencia, df)
                    df_metricas, df_dist_urgencia = calcular_metricas(sequencia, batches)

                    # --- Preparação do arquivo Excel para download ---
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        df_relatorio.to_excel(writer, sheet_name='Sequencia_Otimizada', index=False)
                        df_metricas.to_excel(writer, sheet_name='Metricas_Performance', index=False)
                        df_dist_urgencia.to_excel(writer, sheet_name='Distribuicao_Urgencia', index=False)
                        
                        # Aplica a mesma formatação profissional do script original
                        workbook = writer.book
                        header_format = workbook.add_format({
                            'bold': True, 'text_wrap': False, 'valign': 'vcenter',
                            'fg_color': '#4F81BD', 'font_color': 'white', 'border': 1
                        })
                        
                        for sheet_name in writer.sheets:
                            worksheet = writer.sheets[sheet_name]
                            worksheet.set_zoom(63)
                            worksheet.freeze_panes(1, 0)
                            
                            df_to_format = None
                            if sheet_name == 'Sequencia_Otimizada': df_to_format = df_relatorio
                            elif sheet_name == 'Metricas_Performance': df_to_format = df_metricas
                            else: df_to_format = df_dist_urgencia
                            
                            for col_num, value in enumerate(df_to_format.columns.values):
                                worksheet.write(0, col_num, value, header_format)
                            
                            for i, col in enumerate(df_to_format.columns):
                                max_len = df_to_format[col].astype(str).map(len).max()
                                col_width = max(max_len, len(col)) + 2
                                worksheet.set_column(i, i, col_width)
                        
                        # Oculta colunas M e N
                        worksheet_seq = writer.sheets['Sequencia_Otimizada']
                        worksheet_seq.set_column('M:N', None, None, {'hidden': True})

                    # Disponibiliza o arquivo para download
                    st.success("Programação otimizada gerada com sucesso!")
                    
                    # --- 4. Exibição dos Resultados ---
                    st.header("4. Resultados da Otimização")
                    
                    st.download_button(
                        label="Clique aqui para baixar o relatório completo em Excel",
                        data=output.getvalue(),
                        file_name="programacao_otimizada.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

                    st.subheader("Métricas de Performance")
                    st.dataframe(df_metricas)
                    
                    st.subheader("Distribuição de Lotes por Urgência")
                    st.dataframe(df_dist_urgencia)
                    
                    st.subheader("Prévia da Sequência Otimizada")
                    # Exibe o relatório final sem as colunas ocultas para uma prévia limpa
                    st.dataframe(df_relatorio.drop(columns=['SETUP', 'URGENCIA']))

            except Exception as e:
                st.error(f"Ocorreu um erro durante o processo: {e}")
else:
    st.info("Aguardando o upload da planilha Excel para começar.")