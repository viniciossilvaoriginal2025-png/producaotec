import streamlit as st
import pandas as pd
import plotly.express as px

# Configuração de Layout
st.set_page_config(page_title="Gestão de Produção Técnica", layout="wide")

# Função de carregamento com cache e otimização de memória
@st.cache_data(show_spinner="Processando planilha...", max_entries=2)
def load_data(file):
    try:
        # Lê apenas os nomes das colunas para mapear os índices
        preview = pd.read_excel(file, header=1, nrows=0)
        all_cols = preview.columns.tolist()
        
        # Coluna A (Técnico) = Índice 0 | Coluna P (Tipo) = Índice 15
        idx_tec = 0
        idx_tipo = 15 if len(all_cols) > 15 else len(all_cols) - 1
        
        # Carrega APENAS as duas colunas necessárias
        df = pd.read_excel(
            file, 
            header=1, 
            usecols=[idx_tec, idx_tipo],
            engine='openpyxl'
        )
        
        c_tec = df.columns[0]
        c_tipo = df.columns[1]
        
        # Limpeza e conversão para categorias (economiza 90% de RAM)
        df = df.dropna().copy()
        df[c_tec] = df[c_tec].astype(str).str.strip().str.upper().astype('category')
        df[c_tipo] = df[c_tipo].astype(str).str.strip().astype('category')
        
        return df, c_tec, c_tipo
    except Exception as e:
        st.error(f"Erro ao processar: {e}")
        return None, None, None

st.title("📊 Monitor de Produção Técnica")
st.sidebar.markdown("### Configurações")

uploaded_file = st.sidebar.file_uploader("Upload da Planilha (.xlsx)", type=["xlsx"])

if uploaded_file:
    df, col_tec, col_tipo = load_data(uploaded_file)
    
    if df is not None:
        # Métricas Principais
        m1, m2, m3 = st.columns(3)
        m1.metric("Total Atendimentos", len(df))
        m2.metric("Qtd Técnicos", df[col_tec].nunique())
        m3.metric("Produtividade Média", round(len(df)/df[col_tec].nunique(), 1))

        st.markdown("---")
        
        c1, c2 = st.columns([6, 4])

        with c1:
            st.subheader("🏆 Ranking por Técnico")
            # Agrupamento para o gráfico
            ranking = df[col_tec].value_counts().reset_index()
            ranking.columns = ['Técnico', 'Quantidade']
            
            fig_bar = px.bar(
                ranking, x='Quantidade', y='Técnico', 
                orientation='h', color='Quantidade',
                color_continuous_scale='Blues', text_auto=True
            )
            fig_bar.update_layout(yaxis={'categoryorder':'total ascending'})
            st.plotly_chart(fig_bar, use_container_width=True)

        with c2:
            st.subheader("🎯 Mix de Atendimentos")
            mix = df[col_tipo].value_counts().reset_index()
            mix.columns = ['Tipo', 'Quantidade']
            
            # --- ÚNICA ALTERAÇÃO AQUI ---
            mix = mix[mix['Quantidade'] >= 5]
            # ----------------------------
            
            st.plotly_chart(px.pie(mix, values='Quantidade', names='Tipo', hole=0.5), use_container_width=True)

        st.subheader("📋 Matriz Técnico x Tipo")
        # Tabela dinâmica com totais
        pivot = pd.crosstab(df[col_tec], df[col_tipo], margins=True, margins_name="TOTAL")
        st.dataframe(pivot, use_container_width=True)
        
        if st.sidebar.button("Limpar Cache"):
            st.cache_data.clear()
            st.rerun()
else:
    st.info("Aguardando o upload da planilha Excel...")
