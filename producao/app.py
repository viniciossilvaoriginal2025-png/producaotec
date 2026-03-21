import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO
from datetime import datetime

st.set_page_config(layout="wide")

st.title("📊 Produção Técnica — Análise")

# Inicializa o armazenamento das rotas na memória da sessão
if 'rotas_personalizadas' not in st.session_state:
    st.session_state['rotas_personalizadas'] = {}

arquivo = st.file_uploader("Enviar arquivo Excel", type=["xlsx"])

if arquivo:

    df = pd.read_excel(arquivo, header=1)

    # =========================
    # COLUNAS
    # =========================

    COL_BAIRRO = df.columns[8]     # I
    COL_TECNICO = df.columns[18]   # S
    COL_SERVICO = df.columns[20]   # U
    COL_ENCAM = df.columns[23]     # X
    COL_FECH = df.columns[31]      # AF

    # =========================
    # LIMPEZA
    # =========================

    df[COL_BAIRRO] = (
        df[COL_BAIRRO]
        .astype(str)
        .str.strip()
        .replace(["nan", "None", ""], "Sem bairro")
    )

    # =========================
    # DATAS
    # =========================

    df[COL_ENCAM] = pd.to_datetime(
        df[COL_ENCAM].astype(str).str.replace(".", ":", regex=False),
        dayfirst=True,
        errors="coerce"
    )

    df[COL_FECH] = pd.to_datetime(
        df[COL_FECH].astype(str).str.replace(".", ":", regex=False),
        dayfirst=True,
        errors="coerce"
    )

    df["TEMPO_DELTA"] = df[COL_FECH] - df[COL_ENCAM]

    def formatar_tempo(td):
        if pd.isna(td):
            return ""
        total = int(td.total_seconds())
        d = total // 86400
        h = (total % 86400) // 3600
        m = (total % 3600) // 60
        s = total % 60
        return f"{d}d {h}h {m}m {s}s"

    df["TEMPO_DHMS"] = df["TEMPO_DELTA"].apply(formatar_tempo)

    # =========================
    # FILTROS
    # =========================

    st.sidebar.header("🔎 Filtros")

    tecnicos = sorted(df[COL_TECNICO].dropna().unique())
    servicos = sorted(df[COL_SERVICO].dropna().unique())

    with st.sidebar.expander("👷 Técnicos"):

        marcar_todos_tec = st.checkbox("Selecionar todos técnicos", True)

        tecnicos_sel = []
        for t in tecnicos:
            if st.checkbox(t, value=marcar_todos_tec, key=f"tec_{t}"):
                tecnicos_sel.append(t)

    with st.sidebar.expander("🛠️ Serviços"):

        marcar_todos_serv = st.checkbox("Selecionar todos serviços", True)

        servicos_sel = []
        for s in servicos:
            if st.checkbox(s, value=marcar_todos_serv, key=f"serv_{s}"):
                servicos_sel.append(s)

    # --- NOVA SEÇÃO: CRIAR ROTAS ---
    st.sidebar.header("🗺️ Criar Rotas")
    with st.sidebar.expander("➕ Nova Rota de Atendimento"):
        nome_nova_rota = st.text_input("Nome da Rota (ex: Rota Leste)")
        
        # Novo campo para definir a quantidade de técnicos
        qtd_tecnicos = st.number_input("Quantidade de Técnicos", min_value=1, value=1, step=1)
        
        bairros_unicos = sorted(df[COL_BAIRRO].unique())
        bairros_selecionados_rota = st.multiselect("Selecione os Bairros da Rota", bairros_unicos)
        
        if st.button("Salvar Rota"):
            if nome_nova_rota and bairros_selecionados_rota:
                # Salva os bairros e a quantidade de técnicos na estrutura da rota
                st.session_state['rotas_personalizadas'][nome_nova_rota] = {
                    "bairros": bairros_selecionados_rota,
                    "qtd_tecnicos": qtd_tecnicos
                }
                st.success(f"Rota '{nome_nova_rota}' salva com sucesso!")
            else:
                st.warning("Preencha o nome e selecione pelo menos um bairro.")
        
        if st.session_state['rotas_personalizadas']:
            st.markdown("---")
            st.markdown("**Rotas Ativas:**")
            for r_nome, r_dados in st.session_state['rotas_personalizadas'].items():
                # Tratamento para rotas criadas antes desta atualização
                if isinstance(r_dados, list):
                    qtd_t = 1
                    len_b = len(r_dados)
                else:
                    qtd_t = r_dados.get("qtd_tecnicos", 1)
                    len_b = len(r_dados.get("bairros", []))
                    
                st.markdown(f"- **{r_nome}**: {len_b} bairros | 👷 {qtd_t} técnico(s)")
            if st.button("Limpar Rotas"):
                st.session_state['rotas_personalizadas'] = {}
                st.rerun()

    df_filtrado = df[
        df[COL_TECNICO].isin(tecnicos_sel) &
        df[COL_SERVICO].isin(servicos_sel)
    ].copy()

    # --- MAPEAMENTO DA ROTA NO DATAFRAME ---
    def obter_rota(bairro):
        for nome_rota, dados_rota in st.session_state['rotas_personalizadas'].items():
            # Verifica se é uma lista (versão antiga) ou dicionário (versão nova)
            if isinstance(dados_rota, list):
                if bairro in dados_rota:
                    return nome_rota
            else:
                if bairro in dados_rota.get("bairros", []):
                    return nome_rota
        return "Sem Rota Definida"

    df_filtrado["ROTA_PERSONALIZADA"] = df_filtrado[COL_BAIRRO].apply(obter_rota)

    st.success(f"Registros filtrados: {len(df_filtrado)}")

    # =========================
    # PRODUÇÃO POR TÉCNICO
    # =========================

    st.subheader("👷 Produção por Técnico")

    prod_tecnico = df_filtrado.groupby(COL_TECNICO).size().sort_values(ascending=False)

    st.dataframe(prod_tecnico)
    
    # Converte para DataFrame para usar no Plotly
    df_tecnicos = prod_tecnico.reset_index()
    df_tecnicos.columns = ["Técnico", "Quantidade"]

    # Calcula o total para exibir no gráfico
    total_somado_tec = df_tecnicos["Quantidade"].sum()

    # Cria o gráfico de colunas com o total acima
    fig_tecnicos = px.bar(
        df_tecnicos,
        x="Técnico",
        y="Quantidade",
        text="Quantidade",
        title=f"Total de procedimentos exibidos: {total_somado_tec}"
    )
    
    # Posiciona o texto do lado de fora (acima) da coluna
    fig_tecnicos.update_traces(textposition='outside')
    
    # Dá uma margem extra no topo para o número não cortar e inclina os textos
    max_y_tec = df_tecnicos["Quantidade"].max() if not df_tecnicos.empty else 10
    fig_tecnicos.update_layout(
        yaxis_range=[0, max_y_tec * 1.15],
        xaxis_tickangle=-45,
        margin=dict(t=40)
    )

    st.plotly_chart(fig_tecnicos, use_container_width=True)

    # =========================
    # PRODUÇÃO POR SERVIÇO
    # =========================

    st.subheader("🛠️ Produção por Serviço")

    prod_servico = df_filtrado.groupby(COL_SERVICO).size().sort_values(ascending=False)

    st.dataframe(prod_servico)

    # =========================
    # BAIRROS
    # =========================

    st.subheader("🏘️ Atendimentos por Bairro")

    bairro_counts = df_filtrado[COL_BAIRRO].value_counts()
    bairro_counts = bairro_counts[bairro_counts >= 1]

    st.dataframe(bairro_counts)
    
    # Converte para DataFrame para usar no Plotly
    df_bairros = bairro_counts.reset_index()
    df_bairros.columns = ["Bairro", "Quantidade"]

    # Calcula o total para exibir no gráfico
    total_somado = df_bairros["Quantidade"].sum()

    # Cria o gráfico de colunas com o total acima
    fig_bairros = px.bar(
        df_bairros,
        x="Bairro",
        y="Quantidade",
        text="Quantidade",
        title=f"Total de procedimentos exibidos: {total_somado}"
    )
    
    # Posiciona o texto do lado de fora (acima) da coluna
    fig_bairros.update_traces(textposition='outside')
    
    # Dá uma margem extra no topo para o número não cortar e inclina os textos
    max_y = df_bairros["Quantidade"].max() if not df_bairros.empty else 10
    fig_bairros.update_layout(
        yaxis_range=[0, max_y * 1.15],
        xaxis_tickangle=-45,
        margin=dict(t=40)
    )

    st.plotly_chart(fig_bairros, use_container_width=True)

    # =========================
    # NOVA SEÇÃO — PRODUÇÃO POR ROTA PERSONALIZADA
    # =========================
    
    if st.session_state['rotas_personalizadas']:
        st.subheader("🗺️ Atendimentos por Rota (Personalizada)")

        rota_counts = df_filtrado["ROTA_PERSONALIZADA"].value_counts().reset_index()
        rota_counts.columns = ["Rota", "Quantidade"]

        total_somado_rotas = rota_counts["Quantidade"].sum()

        fig_rotas = px.bar(
            rota_counts,
            x="Rota",
            y="Quantidade",
            text="Quantidade",
            title=f"Total de procedimentos por Rota: {total_somado_rotas}"
        )
        
        fig_rotas.update_traces(textposition='outside')
        
        max_y_rotas = rota_counts["Quantidade"].max() if not rota_counts.empty else 10
        fig_rotas.update_layout(
            yaxis_range=[0, max_y_rotas * 1.15],
            xaxis_tickangle=-45,
            margin=dict(t=40)
        )

        st.plotly_chart(fig_rotas, use_container_width=True)

        # Matriz Técnico x Rota
        st.markdown("**Matriz Técnico × Rota**")
        matriz_rota = pd.crosstab(
            df_filtrado[COL_TECNICO],
            df_filtrado["ROTA_PERSONALIZADA"],
            margins=True, margins_name="TOTAL"
        )
        st.dataframe(matriz_rota, use_container_width=True)

    # =========================
    # NOVO — PROCEDIMENTOS POR BAIRRO
    # =========================

    st.subheader("🛠️ Procedimentos por Bairro")

    proc_bairro = pd.crosstab(
        df_filtrado[COL_BAIRRO],
        df_filtrado[COL_SERVICO]
    )

    st.dataframe(proc_bairro, use_container_width=True)

    bairro_sel_proc = st.selectbox(
        "Selecionar bairro para ver procedimentos",
        sorted(df_filtrado[COL_BAIRRO].unique())
    )

    df_bairro_proc = df_filtrado[df_filtrado[COL_BAIRRO] == bairro_sel_proc]

    ranking_proc = (
        df_bairro_proc[COL_SERVICO]
        .value_counts()
        .reset_index()
    )

    ranking_proc.columns = ["Procedimento", "Quantidade"]

    st.dataframe(ranking_proc, use_container_width=True)
    st.bar_chart(ranking_proc.set_index("Procedimento"))

    # =========================
    # MATRIZ TÉCNICO × BAIRRO
    # =========================

    st.subheader("🧭 Técnicos por Bairro (matriz de atuação)")

    matriz = pd.crosstab(
        df_filtrado[COL_TECNICO],
        df_filtrado[COL_BAIRRO]
    )

    st.dataframe(matriz, use_container_width=True)

    # =========================
    # ATUAÇÃO POR ROTA (BAIRROS DO TÉCNICO)
    # =========================

    st.subheader("📍 Atuação por Rota (Bairros por Técnico)")

    if not df_filtrado.empty:
        tec_rota_sel = st.selectbox(
            "Selecione o Técnico para ver suas áreas de maior atuação:",
            sorted(df_filtrado[COL_TECNICO].unique())
        )

        # Filtra apenas os dados do técnico selecionado
        df_tec_rota = df_filtrado[df_filtrado[COL_TECNICO] == tec_rota_sel]
        
        # Conta a quantidade por bairro
        rota_ranking = df_tec_rota[COL_BAIRRO].value_counts().reset_index()
        rota_ranking.columns = ["Bairro", "Quantidade"]

        # Calcula o total para o título do gráfico
        total_rota = rota_ranking["Quantidade"].sum()

        st.dataframe(rota_ranking, use_container_width=True)

        # Cria o gráfico interativo se houver dados
        if not rota_ranking.empty:
            fig_rota = px.bar(
                rota_ranking,
                x="Bairro",
                y="Quantidade",
                text="Quantidade",
                title=f"Total de atendimentos de {tec_rota_sel} exibidos: {total_rota}"
            )
            fig_rota.update_traces(textposition='outside')
            
            # Ajusta a escala e margens
            max_y_rota = rota_ranking["Quantidade"].max()
            fig_rota.update_layout(
                yaxis_range=[0, max_y_rota * 1.15],
                xaxis_tickangle=-45,
                margin=dict(t=40)
            )
            st.plotly_chart(fig_rota, use_container_width=True)

    # =========================
    # TEMPO MÉDIO
    # =========================

    st.subheader("⏱️ Tempo Médio por Técnico")

    df_tempo = df_filtrado.dropna(subset=["TEMPO_DELTA"])

    if len(df_tempo):

        tempo_medio = (
            df_tempo.groupby(COL_TECNICO)["TEMPO_DELTA"]
            .mean()
            .sort_values(ascending=False)
        )

        st.dataframe(tempo_medio.apply(formatar_tempo))

    # =========================
    # LISTA DETALHADA
    # =========================

    st.subheader("📋 Lista de Atendimentos")

    ordem = st.selectbox("Ordenar por tempo", ["Maior tempo", "Menor tempo"])

    detalhe = df_filtrado.sort_values(
        "TEMPO_DELTA",
        ascending=(ordem == "Menor tempo")
    )

    mostrar = [
        "ROTA_PERSONALIZADA",
        COL_BAIRRO,
        COL_TECNICO,
        COL_SERVICO,
        COL_ENCAM,
        COL_FECH,
        "TEMPO_DHMS"
    ]

    st.dataframe(detalhe[mostrar], height=450)

    # =========================
    # DOWNLOAD LISTA
    # =========================

    buffer = BytesIO()

    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        detalhe[mostrar].to_excel(writer, index=False)

    buffer.seek(0)

    st.download_button(
        "⬇️ Baixar lista filtrada",
        buffer,
        "atendimentos_filtrados.xlsx"
    )

    # =========================
    # RELATÓRIO FORMATADO
    # =========================

    st.subheader("📥 Relatório")

    if st.button("Gerar relatório Excel formatado"):

        buffer = BytesIO()

        resumo = pd.DataFrame({
            "Indicador": [
                "Total de atendimentos",
                "Total de técnicos",
                "Total de bairros",
                "Gerado em"
            ],
            "Valor": [
                len(df_filtrado),
                df_filtrado[COL_TECNICO].nunique(),
                df_filtrado[COL_BAIRRO].nunique(),
                datetime.now().strftime("%d/%m/%Y %H:%M")
            ]
        })

        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            resumo.to_excel(writer, sheet_name="Resumo", index=False)
            prod_tecnico.to_excel(writer, sheet_name="Produção Técnico")
            bairro_counts.to_excel(writer, sheet_name="Produção Bairro")
            matriz.to_excel(writer, sheet_name="Tecnico x Bairro")

        buffer.seek(0)

        st.download_button(
            "⬇️ Baixar relatório",
            buffer,
            "relatorio_producao.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

else:
    st.info("Envie o Excel para iniciar.")
