import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO
from datetime import datetime
import json
import os

st.set_page_config(layout="wide")

st.title("📊 Produção Técnica — Análise")

# =========================
# GERENCIAMENTO DE ROTAS (PERSISTÊNCIA EM ARQUIVO)
# =========================
ARQUIVO_ROTAS = 'rotas_personalizadas.json'

def carregar_rotas():
    if os.path.exists(ARQUIVO_ROTAS):
        try:
            with open(ARQUIVO_ROTAS, 'r', encoding='utf-8') as f:
                return json.load(f)
        except:
            return {}
    return {}

def salvar_rotas(rotas):
    with open(ARQUIVO_ROTAS, 'w', encoding='utf-8') as f:
        json.dump(rotas, f, ensure_ascii=False, indent=4)

# Inicializa o armazenamento das rotas carregando do arquivo
if 'rotas_personalizadas' not in st.session_state:
    st.session_state['rotas_personalizadas'] = carregar_rotas()

# Mapeamento de Cores para o Status
CORES_STATUS = {
    "SOLUCIONADO": "#2ecc71", # Verde
    "PENDENTE": "#f39c12" # Laranja
}

df = None
coluna_status = None

arquivo = st.file_uploader("Enviar arquivo Excel", type=["xlsx"])

if arquivo:
    df = pd.read_excel(arquivo, header=1)

    if not df.empty:
        # --- 🛠️ 1. MAPEAMENTO DE COLUNAS E PREVENÇÃO DE PERDA DE DADOS (O BUG DOS 8 PENDENTES) ---
        # Muitas Visitas Agendadas não possuem um Técnico ou Bairro atribuído.
        # Ao deixar como nulo/vazio, os filtros do Pandas apagavam esses itens. Agora preenchemos com texto.
        try:
            COL_BAIRRO = df.columns[8]     # I
            COL_TECNICO = df.columns[18]   # S
            COL_SERVICO = df.columns[20]   # U
            COL_ENCAM = df.columns[23]     # X
            COL_FECH = df.columns[31]      # AF
        except IndexError:
            st.error("O Excel não tem colunas suficientes. Verifique o formato do arquivo (o cabeçalho deve estar na linha 2).")
            st.stop()

        df[COL_BAIRRO] = df[COL_BAIRRO].astype(str).str.strip().replace(["nan", "None", ""], "Sem Bairro")
        df[COL_TECNICO] = df[COL_TECNICO].astype(str).str.strip().replace(["nan", "None", ""], "Sem Técnico")
        df[COL_SERVICO] = df[COL_SERVICO].astype(str).str.strip().replace(["nan", "None", ""], "Sem Serviço")

        # --- 🚀 2. DETEÇÃO EXATA DA COLUNA DE STATUS ("STATUS ATENDIMENTO" ou "Q") ---
        col_status_otim = None
        for col in df.columns:
            if str(col).strip().upper() in ["STATUS ATENDIMENTO", "Q"]:
                col_status_otim = col
                break
                
        # Fallback de segurança
        if not col_status_otim:
            for col in df.columns:
                amostra = df[col].dropna().astype(str).str.strip().str.upper()
                if amostra.isin(["SOLUCIONADO", "VISITA_AGENDADA", "CONTATO_CLIENTE"]).any():
                    col_status_otim = col
                    break

        col_visita_otim = None
        for col in df.columns:
            if "GEROU_VISITA" in str(col).upper() or "GEROU VISITA" in str(col).upper():
                col_visita_otim = col
                break
                
        # --- 🎯 3. APLICAÇÃO DA REGRA DE NEGÓCIO E FILTRO ---
        if col_status_otim:
            # Substituímos os underlines por espaços para garantir que qualquer variação de escrita seja lida
            status_norm = df[col_status_otim].astype(str).str.strip().str.upper().str.replace("_", " ")
            
            # Apanha TUDO o que tiver "VISITA" ou "CONTATO" de forma tolerante a erros de digitação
            mask_pendentes = status_norm.str.contains('VISITA|CONTATO|AGENDAD', regex=True, na=False)
            mask_solucionado = status_norm.str.contains('SOLUCIONADO', regex=True, na=False)
            
            if col_visita_otim:
                mask_gerou_visita = df[col_visita_otim].astype(str).str.strip().str.upper().isin(['SIM', 'S', 'TRUE', '1', 'VERDADEIRO'])
                mask_solucionado_valido = mask_solucionado & mask_gerou_visita
            else:
                mask_solucionado_valido = mask_solucionado
                
            # Sobrescreve o dataframe apenas com os dados úteis
            df = df[mask_pendentes | mask_solucionado_valido].copy()
            
            if df.empty:
                st.warning("⚠️ Nenhum registro restou após aplicar o filtro de otimização (Apenas Solucionados que geraram visita, Visita Agendada ou Contato Cliente).")
            else:
                # Padronizamos os nomes finais para o Dashboard não se confundir
                # Acedemos ao DataFrame filtrado
                status_filtrado = df[col_status_otim].astype(str).str.strip().str.upper().str.replace("_", " ")
                
                mask_final_pendente = status_filtrado.str.contains('VISITA|CONTATO|AGENDAD', regex=True, na=False)
                mask_final_solucionado = status_filtrado.str.contains('SOLUCIONADO', regex=True, na=False)
                
                df.loc[mask_final_pendente, col_status_otim] = 'PENDENTE'
                df.loc[mask_final_solucionado, col_status_otim] = 'SOLUCIONADO'
                
                # Definimos a variável global que será usada para agrupar os gráficos
                coluna_status = col_status_otim

        # =========================
        # FORMATAÇÃO DE DATAS
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
        # FILTROS (BARRA LATERAL)
        # =========================

        st.sidebar.header("🔎 Filtros")

        # Filtro de Data (Calendário)
        with st.sidebar.expander("📅 Período (Data de Fechamento)", expanded=True):
            valid_dates = df[COL_FECH].dropna()
            if not valid_dates.empty:
                min_date = valid_dates.min().date()
                max_date = valid_dates.max().date()
            else:
                min_date = datetime.today().date()
                max_date = datetime.today().date()

            start_date = st.date_input(
                "Data Inicial:",
                value=min_date,
                min_value=min_date,
                max_value=max_date
            )
            
            end_date = st.date_input(
                "Data Final:",
                value=max_date,
                min_value=min_date,
                max_value=max_date
            )

            if start_date <= end_date:
                # Gera todos os dias dentro do intervalo selecionado
                dias_no_intervalo = pd.date_range(start=start_date, end=end_date).date.tolist()
                dias_formatados = [d.strftime("%d/%m/%Y") for d in dias_no_intervalo]
                
                dias_selecionados_str = st.multiselect(
                    "Dias a incluir (remova para ignorar):",
                    options=dias_formatados,
                    default=dias_formatados,
                    help="Todos os dias do intervalo começam marcados. Clique no 'X' para remover dias específicos (ex: finais de semana)."
                )
                
                dias_validos = [datetime.strptime(d, "%d/%m/%Y").date() for d in dias_selecionados_str]
            else:
                dias_validos = []
                st.error("⚠️ A Data Inicial não pode ser maior que a Data Final.")
            
            incluir_vazios = st.checkbox("Incluir registros sem data de fechamento (Pendentes)", value=True)

        # Filtro de Status visível
        with st.sidebar.expander("📌 Status / Situação", expanded=True):
            if coluna_status:
                st.text_input("Coluna Identificada:", value=coluna_status, disabled=True)
                status_unicos = sorted(df[coluna_status].dropna().astype(str).unique())
                status_sel = st.multiselect(
                    "Selecione os Status para análise:", 
                    options=status_unicos, 
                    default=status_unicos
                )
            else:
                st.error("Coluna de Status não encontrada.")
                status_sel = []

        # Filtros de Técnicos e Serviços
        tecnicos = sorted(df[COL_TECNICO].dropna().unique()) if not df.empty else []
        servicos = sorted(df[COL_SERVICO].dropna().unique()) if not df.empty else []

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

        # --- LÓGICA DE APLICAÇÃO DOS FILTROS FINAIS ---
        if not df.empty and coluna_status:
            start_dt = pd.to_datetime(start_date)
            end_dt = pd.to_datetime(end_date) + pd.Timedelta(days=1, seconds=-1)

            mask_tecnico = df[COL_TECNICO].isin(tecnicos_sel)
            mask_servico = df[COL_SERVICO].isin(servicos_sel)
            mask_status = df[coluna_status].isin(status_sel)
            
            if incluir_vazios:
                mask_data = df[COL_FECH].dt.date.isin(dias_validos) | df[COL_FECH].isna()
            else:
                mask_data = df[COL_FECH].dt.date.isin(dias_validos)

            df_filtrado = df[mask_tecnico & mask_servico & mask_data & mask_status].copy()
        else:
            df_filtrado = pd.DataFrame()

# =========================
# GERAÇÃO DO DASHBOARD
# =========================
if df is not None and not df.empty and not df_filtrado.empty and coluna_status:

    # --- SEÇÃO: GERENCIAR ROTAS ---
    st.sidebar.header("🗺️ Gerenciar Rotas")
    
    bairros_unicos = sorted(df[COL_BAIRRO].unique())
    
    with st.sidebar.expander("➕ Nova Rota"):
        nome_nova_rota = st.text_input("Nome da Rota (ex: Rota Leste)", key="novo_nome")
        qtd_tecnicos = st.number_input("Quantidade de Técnicos", min_value=1, value=1, step=1, key="nova_qtd")
        bairros_selecionados_rota = st.multiselect("Selecione os Bairros da Rota", bairros_unicos, key="novos_bairros")
        
        if st.button("Salvar Nova Rota"):
            if nome_nova_rota and bairros_selecionados_rota:
                # Salva os bairros e a quantidade de técnicos na estrutura da rota
                st.session_state['rotas_personalizadas'][nome_nova_rota] = {
                    "bairros": bairros_selecionados_rota,
                    "qtd_tecnicos": qtd_tecnicos
                }
                salvar_rotas(st.session_state['rotas_personalizadas']) # Salva no arquivo
                st.success(f"Rota '{nome_nova_rota}' salva com sucesso!")
                st.rerun()
            else:
                st.warning("Preencha o nome e selecione pelo menos um bairro.")
    
    # Se houver rotas, mostra o menu de edição
    if st.session_state['rotas_personalizadas']:
        with st.sidebar.expander("✏️ Editar / Excluir Rota"):
            rota_para_editar = st.selectbox(
                "Selecione a Rota para editar", 
                list(st.session_state['rotas_personalizadas'].keys())
            )
            
            if rota_para_editar:
                dados_atuais = st.session_state['rotas_personalizadas'][rota_para_editar]
                
                # Tratamento para compatibilidade retroativa
                if isinstance(dados_atuais, list):
                    qtd_atual = 1
                    bairros_atuais = dados_atuais
                else:
                    qtd_atual = dados_atuais.get("qtd_tecnicos", 1)
                    bairros_atuais = dados_atuais.get("bairros", [])

                # Adicionando a rota_para_editar nas chaves para forçar a atualização dos campos quando a rota selecionada mudar
                edit_nome = st.text_input("Renomear Rota", value=rota_para_editar, key=f"edit_nome_{rota_para_editar}")
                edit_qtd = st.number_input("Editar Quantidade de Técnicos", min_value=1, value=qtd_atual, step=1, key=f"edit_qtd_{rota_para_editar}")
                
                # Garante que os bairros já selecionados estejam nas opções para evitar erros no Streamlit
                opcoes_bairros_edit = sorted(list(set(bairros_unicos + bairros_atuais)))
                edit_bairros = st.multiselect("Editar Bairros", opcoes_bairros_edit, default=bairros_atuais, key=f"edit_bairros_{rota_para_editar}")

                col_salvar, col_excluir = st.columns(2)
                
                with col_salvar:
                    if st.button("Salvar"):
                        if edit_nome and edit_bairros:
                            # Se mudou o nome, exclui a chave antiga
                            if edit_nome != rota_para_editar:
                                del st.session_state['rotas_personalizadas'][rota_para_editar]
                            
                            st.session_state['rotas_personalizadas'][edit_nome] = {
                                "bairros": edit_bairros,
                                "qtd_tecnicos": edit_qtd
                            }
                            salvar_rotas(st.session_state['rotas_personalizadas']) # Salva no arquivo
                            st.success("Rota atualizada!")
                            st.rerun()
                        else:
                            st.warning("Preencha o nome e os bairros.")
                
                with col_excluir:
                    if st.button("Excluir"):
                        del st.session_state['rotas_personalizadas'][rota_para_editar]
                        salvar_rotas(st.session_state['rotas_personalizadas']) # Salva no arquivo
                        st.success("Rota excluída!")
                        st.rerun()

        st.sidebar.markdown("---")
        st.sidebar.markdown("**Rotas Ativas:**")
        for r_nome, r_dados in st.session_state['rotas_personalizadas'].items():
            if isinstance(r_dados, list):
                qtd_t = 1
                len_b = len(r_dados)
            else:
                qtd_t = r_dados.get("qtd_tecnicos", 1)
                len_b = len(r_dados.get("bairros", []))
                
            st.sidebar.markdown(f"- **{r_nome}**: {len_b} bairros | 👷 {qtd_t} técnico(s)")
            
        if st.sidebar.button("Limpar Todas as Rotas"):
            st.session_state['rotas_personalizadas'] = {}
            salvar_rotas(st.session_state['rotas_personalizadas']) # Salva no arquivo (limpa ele)
            st.rerun()

    # --- MAPEAMENTO DA ROTA NO DATAFRAME ---
    def obter_rota(bairro):
        for nome_rota, dados_rota in st.session_state['rotas_personalizadas'].items():
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

    if not df_filtrado.empty:
        # Agrupamento separando por Status
        df_tecnicos_status = df_filtrado.groupby([COL_TECNICO, coluna_status]).size().reset_index(name="Quantidade")
        ordem_tec = df_filtrado[COL_TECNICO].value_counts().index.tolist()

        # Tabela cruzada com Status
        tab_tec = pd.crosstab(df_filtrado[COL_TECNICO], df_filtrado[coluna_status], margins=True, margins_name="TOTAL")
        tab_tec = tab_tec.reindex(index=ordem_tec + ["TOTAL"]).fillna(0).astype(int)
        st.dataframe(tab_tec, use_container_width=True)

        # Gráfico colorido por Status (Lado a Lado)
        total_somado_tec = df_tecnicos_status["Quantidade"].sum()
        fig_tecnicos = px.bar(
            df_tecnicos_status,
            x=COL_TECNICO,
            y="Quantidade",
            color=coluna_status,
            text="Quantidade",
            title=f"Total de procedimentos exibidos: {total_somado_tec}",
            category_orders={COL_TECNICO: ordem_tec},
            color_discrete_map=CORES_STATUS
        )
        
        fig_tecnicos.update_traces(textposition='outside')
        fig_tecnicos.update_layout(xaxis_tickangle=-45, margin=dict(t=40), barmode='group')
        st.plotly_chart(fig_tecnicos, use_container_width=True)

    # =========================
    # PRODUÇÃO POR SERVIÇO
    # =========================

    st.subheader("🛠️ Produção por Serviço")

    if not df_filtrado.empty:
        ordem_serv = df_filtrado[COL_SERVICO].value_counts().index.tolist()
        tab_serv = pd.crosstab(df_filtrado[COL_SERVICO], df_filtrado[coluna_status], margins=True, margins_name="TOTAL")
        tab_serv = tab_serv.reindex(index=ordem_serv + ["TOTAL"]).fillna(0).astype(int)
        st.dataframe(tab_serv, use_container_width=True)

    # =========================
    # BAIRROS
    # =========================

    st.subheader("🏘️ Atendimentos por Bairro")

    if not df_filtrado.empty:
        # Filtrar bairros com >= 5 atendimentos no total
        bairro_totals = df_filtrado[COL_BAIRRO].value_counts()
        bairros_validos = bairro_totals[bairro_totals >= 5].index.tolist()
        df_bairros_filtrado = df_filtrado[df_filtrado[COL_BAIRRO].isin(bairros_validos)]

        if not df_bairros_filtrado.empty:
            df_bairros_status = df_bairros_filtrado.groupby([COL_BAIRRO, coluna_status]).size().reset_index(name="Quantidade")
            
            # Tabela cruzada
            tab_bairros = pd.crosstab(df_bairros_filtrado[COL_BAIRRO], df_bairros_filtrado[coluna_status], margins=True, margins_name="TOTAL")
            tab_bairros = tab_bairros.reindex(index=bairros_validos + ["TOTAL"]).fillna(0).astype(int)
            st.dataframe(tab_bairros, use_container_width=True)

            # Gráfico colorido (Lado a Lado)
            total_somado_bairros = df_bairros_status["Quantidade"].sum()
            fig_bairros = px.bar(
                df_bairros_status,
                x=COL_BAIRRO,
                y="Quantidade",
                color=coluna_status,
                text="Quantidade",
                title=f"Total de procedimentos exibidos: {total_somado_bairros}",
                category_orders={COL_BAIRRO: bairros_validos},
                color_discrete_map=CORES_STATUS
            )
            
            fig_bairros.update_traces(textposition='outside')
            fig_bairros.update_layout(xaxis_tickangle=-45, margin=dict(t=40), barmode='group')
            st.plotly_chart(fig_bairros, use_container_width=True)

    # =========================
    # NOVA SEÇÃO — PRODUÇÃO POR ROTA PERSONALIZADA
    # =========================
    
    if st.session_state['rotas_personalizadas'] and not df_filtrado.empty:
        st.subheader("🗺️ Atendimentos por Rota (Personalizada)")

        df_rotas_status = df_filtrado.groupby(["ROTA_PERSONALIZADA", coluna_status]).size().reset_index(name="Quantidade")
        ordem_rotas = df_filtrado["ROTA_PERSONALIZADA"].value_counts().index.tolist()

        if not df_rotas_status.empty:
            total_somado_rotas = df_rotas_status["Quantidade"].sum()

            fig_rotas = px.bar(
                df_rotas_status,
                x="ROTA_PERSONALIZADA",
                y="Quantidade",
                color=coluna_status,
                text="Quantidade",
                title=f"Total de procedimentos por Rota: {total_somado_rotas}",
                category_orders={"ROTA_PERSONALIZADA": ordem_rotas},
                color_discrete_map=CORES_STATUS,
                labels={"ROTA_PERSONALIZADA": "Rota"} # Adiciona o rótulo legível
            )
            
            fig_rotas.update_traces(textposition='outside')
            fig_rotas.update_layout(xaxis_tickangle=-45, margin=dict(t=40), barmode='group')
            st.plotly_chart(fig_rotas, use_container_width=True)

        # Matriz Técnico x Rota
        st.markdown("**Matriz Técnico × Rota (Total Geral)**")
        matriz_rota = pd.crosstab(
            df_filtrado[COL_TECNICO],
            df_filtrado["ROTA_PERSONALIZADA"],
            margins=True, margins_name="TOTAL"
        )
        st.dataframe(matriz_rota, use_container_width=True)

    # =========================
    # PROCEDIMENTOS POR BAIRRO
    # =========================

    st.subheader("🛠️ Procedimentos por Bairro")

    if not df_filtrado.empty:
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
        
        if not df_bairro_proc.empty:
            df_ranking_proc = df_bairro_proc.groupby([COL_SERVICO, coluna_status]).size().reset_index(name="Quantidade")
            ordem_proc = df_bairro_proc[COL_SERVICO].value_counts().index.tolist()

            tab_proc = pd.crosstab(df_bairro_proc[COL_SERVICO], df_bairro_proc[coluna_status], margins=True, margins_name="TOTAL")
            tab_proc = tab_proc.reindex(index=ordem_proc + ["TOTAL"]).fillna(0).astype(int)
            st.dataframe(tab_proc, use_container_width=True)

            fig_proc = px.bar(
                df_ranking_proc, 
                x=COL_SERVICO, 
                y="Quantidade", 
                color=coluna_status, 
                text="Quantidade", 
                category_orders={COL_SERVICO: ordem_proc}, 
                color_discrete_map=CORES_STATUS
            )
            fig_proc.update_traces(textposition='outside')
            fig_proc.update_layout(barmode='group')
            st.plotly_chart(fig_proc, use_container_width=True)

    # =========================
    # MATRIZ TÉCNICO × BAIRRO
    # =========================

    st.subheader("🧭 Técnicos por Bairro (matriz de atuação)")

    if not df_filtrado.empty:
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

        df_tec_rota = df_filtrado[df_filtrado[COL_TECNICO] == tec_rota_sel]
        
        if not df_tec_rota.empty:
            df_rota_ranking = df_tec_rota.groupby([COL_BAIRRO, coluna_status]).size().reset_index(name="Quantidade")
            ordem_bairros_tec = df_tec_rota[COL_BAIRRO].value_counts().index.tolist()

            tab_rota_tec = pd.crosstab(df_tec_rota[COL_BAIRRO], df_tec_rota[coluna_status], margins=True, margins_name="TOTAL")
            tab_rota_tec = tab_rota_tec.reindex(index=ordem_bairros_tec + ["TOTAL"]).fillna(0).astype(int)
            st.dataframe(tab_rota_tec, use_container_width=True)

            total_rota = df_rota_ranking["Quantidade"].sum()

            fig_rota = px.bar(
                df_rota_ranking,
                x=COL_BAIRRO,
                y="Quantidade",
                color=coluna_status,
                text="Quantidade",
                title=f"Total de atendimentos de {tec_rota_sel} exibidos: {total_rota}",
                category_orders={COL_BAIRRO: ordem_bairros_tec},
                color_discrete_map=CORES_STATUS
            )
            fig_rota.update_traces(textposition='outside')
            fig_rota.update_layout(xaxis_tickangle=-45, margin=dict(t=40), barmode='group')
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
        coluna_status,
        COL_BAIRRO,
        COL_TECNICO,
        COL_SERVICO,
        COL_ENCAM,
        COL_FECH,
        "TEMPO_DHMS"
    ]

    # Função para colorir a coluna de Status
    def colorir_status(val):
        if val == 'SOLUCIONADO':
            return 'background-color: #d4edda; color: #155724' # Verde claro
        elif val == 'PENDENTE':
            return 'background-color: #fff3cd; color: #856404' # Amarelo claro
        return ''

    # Aplica o estilo na tabela
    try:
        df_estilizado = detalhe[mostrar].style.map(colorir_status, subset=[coluna_status])
    except AttributeError:
        # Fallback para versões anteriores do Pandas
        df_estilizado = detalhe[mostrar].style.applymap(colorir_status, subset=[coluna_status])

    st.dataframe(df_estilizado, height=450)

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
            if 'tab_tec' in locals(): tab_tec.to_excel(writer, sheet_name="Produção Técnico")
            if 'tab_bairros' in locals(): tab_bairros.to_excel(writer, sheet_name="Produção Bairro")
            if 'matriz_rota' in locals(): matriz_rota.to_excel(writer, sheet_name="Tecnico x Bairro")

        buffer.seek(0)

        st.download_button(
            "⬇️ Baixar relatório",
            buffer,
            "relatorio_producao.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

elif arquivo is not None:
    pass # Os dados não passaram na validação do filtro

else:
    st.info("Envie o Excel para iniciar.")
