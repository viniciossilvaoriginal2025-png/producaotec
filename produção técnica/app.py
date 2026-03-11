import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime

st.set_page_config(layout="wide")

st.title("📊 Produção Técnica — Análise")

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

    df_filtrado = df[
        df[COL_TECNICO].isin(tecnicos_sel) &
        df[COL_SERVICO].isin(servicos_sel)
    ]

    st.success(f"Registros filtrados: {len(df_filtrado)}")

    # =========================
    # PRODUÇÃO POR TÉCNICO
    # =========================

    st.subheader("👷 Produção por Técnico")

    prod_tecnico = df_filtrado.groupby(COL_TECNICO).size().sort_values(ascending=False)

    st.dataframe(prod_tecnico)
    st.bar_chart(prod_tecnico)

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

    bairro_counts = df_filtrado[COL_BAIRRO].value_counts().head(15)

    st.dataframe(bairro_counts)
    st.bar_chart(bairro_counts)

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
