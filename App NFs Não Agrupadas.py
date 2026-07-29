import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from io import BytesIO
from pathlib import Path
from urllib.request import Request, urlopen
from urllib.error import HTTPError, URLError

st.set_page_config(
    page_title="NFs Faturadas - Protocolocadas - Agrupadas",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded",
)

ARQUIVO_PADRAO = "Base Tracking.xlsx"


def obter_configuracao(nome, padrao=""):
    """Lê uma configuração do secrets.toml sem exigir que o arquivo exista."""
    try:
        return st.secrets.get(nome, padrao)
    except Exception:
        return padrao


def localizar_base():
    """
    Prioridade de carregamento da base:
    1. URL definida em BASE_TRACKING_URL nos Secrets do Streamlit;
    2. arquivo Base Tracking.xlsx salvo junto com o app no repositório.

    Para repositório privado, GITHUB_TOKEN pode ser informado nos Secrets.
    """
    url = str(obter_configuracao("BASE_TRACKING_URL", "")).strip()
    if url:
        headers = {"User-Agent": "streamlit-nfs-agrupadas"}
        token = str(obter_configuracao("GITHUB_TOKEN", "")).strip()
        if token:
            headers["Authorization"] = f"Bearer {token}"

        try:
            requisicao = Request(url, headers=headers)
            with urlopen(requisicao, timeout=60) as resposta:
                return BytesIO(resposta.read())
        except HTTPError as erro:
            raise RuntimeError(
                f"Não foi possível baixar a base do GitHub (HTTP {erro.code}). "
                "Confira a URL e, se o repositório for privado, o GITHUB_TOKEN."
            ) from erro
        except URLError as erro:
            raise RuntimeError(f"Não foi possível acessar a base no GitHub: {erro.reason}") from erro

    caminho_local = Path(__file__).resolve().parent / ARQUIVO_PADRAO
    if caminho_local.exists():
        return caminho_local

    raise FileNotFoundError(
        f"Arquivo {ARQUIVO_PADRAO} não encontrado no repositório e "
        "BASE_TRACKING_URL não foi configurada nos Secrets do Streamlit."
    )

COLS_OBRIGATORIAS = [
    "Unidade", "Nome Cliente", "CNPJ/CPF", "Canal Venda", "SubCanal Venda",
    "Pedido", "Protocolo", "Data Nota Fiscal", "Data Protocolo", "Transportadora", "Nota Fiscal"
]

COLS_TABELA = [
    "Unidade", "Data do Protocolo", "Nome do cliente", "CNPJ", "Canal", "Sub Canal",
    "Pedido", "Protocolo", "Transportadora"
]

COLS_EXPORTAR = [
    "Unidade", "Data do Protocolo", "Nome do cliente", "CNPJ", "Canal", "Sub Canal",
    "Pedido", "Protocolo", "Transportadora", "Data", "Nota Fiscal", "NFs não agrupadas",
    "Qtd Protocolos Grupo", "Qtd NFs Grupo"
]

st.markdown(
    """
    <style>
        .main .block-container {padding-top: 1.2rem;}
        div[data-testid="stMetric"] {
            background-color: #ffffff;
            border: 1px solid #e6e6e6;
            padding: 16px 18px;
            border-radius: 14px;
            box-shadow: 0 1px 4px rgba(0,0,0,0.06);
        }
        div[data-testid="stMetric"] label {font-size: 0.90rem !important; color: #555 !important;}
        div[data-testid="stMetric"] div[data-testid="stMetricValue"] {font-size: 1.75rem !important;}
        .titulo-card {font-weight: 700; font-size: 1.05rem; margin: 0.8rem 0 0.2rem 0;}
        .descricao {color: #666; font-size: 0.95rem;}
    </style>
    """,
    unsafe_allow_html=True,
)

@st.cache_data(show_spinner="Carregando base...")
def carregar_base(caminho_ou_buffer):
    df = pd.read_excel(caminho_ou_buffer, engine="openpyxl", dtype=str)
    df.columns = df.columns.astype(str).str.strip()

    faltantes = [c for c in COLS_OBRIGATORIAS if c not in df.columns]
    if faltantes:
        raise ValueError("Colunas obrigatórias não localizadas na base: " + ", ".join(faltantes))

    for col in COLS_OBRIGATORIAS:
        df[col] = df[col].fillna("").astype(str).str.strip()

    df["Data NF"] = pd.to_datetime(df["Data Nota Fiscal"], errors="coerce").dt.date
    df["Data"] = pd.to_datetime(df["Data NF"], errors="coerce").dt.strftime("%d/%m/%Y")
    df["Data Protocolo Dt"] = pd.to_datetime(df["Data Protocolo"], errors="coerce")
    df["Data do Protocolo"] = df["Data Protocolo Dt"].dt.strftime("%d/%m/%Y %H:%M")

    df = df.rename(columns={
        "Nome Cliente": "Nome do cliente",
        "CNPJ/CPF": "CNPJ",
        "Canal Venda": "Canal",
        "SubCanal Venda": "Sub Canal",
    })

    df["Transportadora Normalizada"] = df["Transportadora"].str.upper().str.strip()
    df["Protocolo"] = df["Protocolo"].fillna("").astype(str).str.strip()
    df["Nota Fiscal"] = df["Nota Fiscal"].fillna("").astype(str).str.strip()
    df["CNPJ"] = df["CNPJ"].fillna("").astype(str).str.strip()
    df["Pedido"] = df["Pedido"].fillna("").astype(str).str.strip()
    return df


def filtrar_transportadoras_validas(df, transportadoras_desconsiderar=None):
    """Remove as transportadoras selecionadas pelo usuário.

    A regra usa correspondência por "contém" para cobrir variações do nome.
    Exemplo: selecionar CORREIOS remove CORREIOS, CORREIOS S/A etc.
    """
    transportadoras_desconsiderar = transportadoras_desconsiderar or []
    transportadora = df["Transportadora Normalizada"].fillna("").astype(str)
    mask_desconsiderar = transportadora.apply(
        lambda valor: any(
            str(item).strip().upper() in valor
            for item in transportadoras_desconsiderar
            if str(item).strip() != ""
        )
    )
    return df[~mask_desconsiderar].copy()


def filtrar_transportadoras_consideradas(df, transportadoras_considerar=None):
    """Mantém somente as transportadoras selecionadas pelo usuário.

    Se nada for selecionado, mantém todas as transportadoras já disponíveis após as demais regras.
    A regra usa correspondência por "contém" para cobrir variações do nome.
    """
    transportadoras_considerar = transportadoras_considerar or []
    if not transportadoras_considerar:
        return df.copy()

    transportadora = df["Transportadora Normalizada"].fillna("").astype(str)
    mask_considerar = transportadora.apply(
        lambda valor: any(
            str(item).strip().upper() in valor
            for item in transportadoras_considerar
            if str(item).strip() != ""
        )
    )
    return df[mask_considerar].copy()


def aplicar_regra_cenario(df, transportadoras_desconsiderar=None, transportadoras_considerar=None):
    base = filtrar_transportadoras_validas(df, transportadoras_desconsiderar)
    base = filtrar_transportadoras_consideradas(base, transportadoras_considerar)
    chaves = ["CNPJ", "Data NF", "Unidade"]

    base["Qtd Protocolos Grupo"] = base.groupby(chaves)["Protocolo"].transform("nunique")
    base["Qtd NFs Grupo"] = base.groupby(chaves)["Nota Fiscal"].transform("nunique")

    cenario = base[(base["Qtd Protocolos Grupo"] > 1) & (base["Qtd NFs Grupo"] > 1)].copy()

    cenario = cenario.sort_values(
        by=["CNPJ", "Data NF", "Unidade", "Data Protocolo Dt", "Protocolo", "Pedido", "Nota Fiscal"],
        ascending=True,
        na_position="last"
    ).copy()
    cenario["Ordem NF Grupo"] = cenario.groupby(chaves).cumcount()
    cenario["NFs não agrupadas"] = (cenario["Ordem NF Grupo"] > 0).astype(int)

    return cenario


def opcoes_ordenadas(serie):
    return sorted([x for x in serie.dropna().astype(str).unique() if x.strip() != ""])


def aplicar_filtros(df, unidades=None, datas=None, canais=None, subcanais=None):
    filtrado = df.copy()
    if unidades:
        filtrado = filtrado[filtrado["Unidade"].isin(unidades)]
    if datas:
        filtrado = filtrado[filtrado["Data"].isin(datas)]
    if canais:
        filtrado = filtrado[filtrado["Canal"].isin(canais)]
    if subcanais:
        filtrado = filtrado[filtrado["Sub Canal"].isin(subcanais)]
    return filtrado




def formatar_percentual(valor):
    return f"{valor:.2%}".replace(".", ",")

def card_dataframe(df_cenario, df_base_nf_sem_cnpj, grupo, titulo):
    st.markdown(f"<div class='titulo-card'>{titulo}</div>", unsafe_allow_html=True)

    # Tabela de resumo ajustada conforme solicitado:
    # 1) Total de NFs Agrupáveis
    # 2) Total de NFs não agrupadas
    # 3) % Ineficiência de Agrupamento = Total de NFs não agrupadas / Total de NFs Agrupáveis
    resumo = (
        df_cenario.groupby(grupo, dropna=False)
          .agg(**{
              "Total de NFs Agrupáveis": ("Nota Fiscal", "nunique"),
              "Total de NFs não agrupadas": ("NFs não agrupadas", "sum"),
          })
          .reset_index()
    ) if not df_cenario.empty else pd.DataFrame(columns=[grupo, "Total de NFs Agrupáveis", "Total de NFs não agrupadas"])

    if resumo.empty:
        st.info("Sem dados para exibir neste agrupamento.")
        return

    resumo["Total de NFs Agrupáveis"] = resumo["Total de NFs Agrupáveis"].astype(int)
    resumo["Total de NFs não agrupadas"] = resumo["Total de NFs não agrupadas"].astype(int)

    # Mantém o percentual como número para o Streamlit alinhar à direita.
    resumo["% Ineficiência de Agrupamento"] = resumo.apply(
        lambda row: (row["Total de NFs não agrupadas"] / row["Total de NFs Agrupáveis"] * 100)
        if row["Total de NFs Agrupáveis"] else 0,
        axis=1
    )

    resumo = resumo[[
        grupo,
        "Total de NFs Agrupáveis",
        "Total de NFs não agrupadas",
        "% Ineficiência de Agrupamento"
    ]].sort_values("Total de NFs não agrupadas", ascending=False)

    st.dataframe(
        resumo,
        use_container_width=True,
        hide_index=True,
        column_config={
            "% Ineficiência de Agrupamento": st.column_config.NumberColumn(
                "% Ineficiência de Agrupamento",
                format="%.2f%%"
            )
        }
    )


def converter_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Relacao")
    return output.getvalue()


def tabela_ocorrencias(df):
    st.subheader("Tabela de Ocorrências")

    coluna_ocorrencia = "Última Ocorrência"

    if coluna_ocorrencia not in df.columns:
        st.info("Coluna de ocorrência não localizada na base. Verifique se existe uma coluna chamada Última Ocorrência.")
        return

    if df.empty:
        st.info("Sem dados para exibir na tabela de ocorrências.")
        return

    base = df.copy()
    base[coluna_ocorrencia] = base[coluna_ocorrencia].fillna("").astype(str).str.strip()
    base[coluna_ocorrencia] = base[coluna_ocorrencia].replace("", "Sem ocorrência informada")

    tabela = (
        base.groupby(coluna_ocorrencia, dropna=False)
        .size()
        .reset_index(name="Quantidade")
        .rename(columns={coluna_ocorrencia: "Ocorrência"})
        .sort_values("Quantidade", ascending=False)
        .reset_index(drop=True)
    )

    total = int(tabela["Quantidade"].sum())
    tabela["%"] = tabela["Quantidade"].apply(lambda qtd: (qtd / total * 100) if total else 0)

    linha_total = pd.DataFrame({
        "Ocorrência": ["Total"],
        "Quantidade": [total],
        "%": [100.0 if total else 0.0]
    })

    tabela = pd.concat([tabela, linha_total], ignore_index=True)

    def destacar_total(row):
        if row["Ocorrência"] == "Total":
            return [
                "background-color: #e8f2ff; color: #0b3d91; font-weight: 700; border-top: 2px solid #0b3d91;"
            ] * len(row)
        return [""] * len(row)

    tabela_formatada = (
        tabela.style
        .apply(destacar_total, axis=1)
        .format({"%": "{:.2f}%", "Quantidade": "{:,.0f}"})
    )

    st.dataframe(
        tabela_formatada,
        use_container_width=True,
        hide_index=True
    )


def grafico_pizza_nfs_nao_agrupadas_por_unidade(df):
    st.subheader("Gráfico de pizza - NFs não agrupadas por Unidade")

    if df.empty or "NFs não agrupadas" not in df.columns:
        st.info("Sem dados para gerar o gráfico.")
        return

    pizza_base = (
        df.groupby("Unidade", dropna=False)["NFs não agrupadas"]
          .sum()
          .reset_index()
          .query("`NFs não agrupadas` > 0")
          .sort_values("NFs não agrupadas", ascending=False)
    )

    if pizza_base.empty:
        st.info("Não há NFs não agrupadas para exibir no gráfico.")
        return

    total = pizza_base["NFs não agrupadas"].sum()
    pizza_base["% Participação"] = pizza_base["NFs não agrupadas"] / total

    principais = pizza_base[pizza_base["% Participação"] >= 0.05].copy()
    outros = pizza_base[pizza_base["% Participação"] < 0.05].copy()

    if not outros.empty:
        linha_outros = pd.DataFrame({
            "Unidade": ["OUTROS"],
            "NFs não agrupadas": [outros["NFs não agrupadas"].sum()],
            "% Participação": [outros["NFs não agrupadas"].sum() / total]
        })
        pizza = pd.concat([principais, linha_outros], ignore_index=True)
    else:
        pizza = principais.copy()

    pizza = pizza.sort_values("NFs não agrupadas", ascending=False).reset_index(drop=True)

    # Redução solicitada: gráfico e fontes 30% menores.
    # Tamanho anterior: 5.6 x 5.6. Novo tamanho: 3.92 x 3.92.
    fig, ax = plt.subplots(figsize=(3.92, 3.92))
    wedges, texts, autotexts = ax.pie(
        pizza["NFs não agrupadas"],
        labels=pizza["Unidade"],
        autopct=lambda pct: f"{pct:.1f}%" if pct > 0 else "",
        startangle=90,
        textprops={"fontsize": 7},
        pctdistance=0.65,
        labeldistance=1.10
    )

    # Garante que rótulos externos e percentuais internos fiquem 30% menores.
    for texto in texts:
        texto.set_fontsize(7)
    for texto in autotexts:
        texto.set_fontsize(7)

    ax.axis("equal")
    ax.set_title("Quantidade de NFs não agrupadas por Unidade", fontsize=8.4, pad=8)
    fig.tight_layout()

    # Mantém o gráfico no tamanho real definido no Matplotlib,
    # evitando que o Streamlit estique novamente para a largura inteira da tela.
    st.pyplot(fig, use_container_width=False)

    with st.expander("Ver dados do gráfico"):
        tabela_grafico = pizza.copy()
        tabela_grafico["% Participação"] = (tabela_grafico["% Participação"] * 100).round(2).astype(str) + "%"
        st.dataframe(tabela_grafico, use_container_width=True, hide_index=True)

    with st.expander("Ver unidades agrupadas em OUTROS"):
        if outros.empty:
            st.info("Nenhuma unidade ficou abaixo de 5%.")
        else:
            outros_exibir = outros.copy()
            outros_exibir["% Participação"] = (outros_exibir["% Participação"] * 100).round(2).astype(str) + "%"
            st.dataframe(outros_exibir, use_container_width=True, hide_index=True)



st.title("📦 NFs (Faturadas - Protocolocadas - Agrupadas) ")
st.markdown(
    "<div class='descricao'>Cenário considera NF emitida no mesmo dia para o mesmo CNPJ, "
    "no mesmo local de origem, com protocolos distintos e desconsiderando as transportadoras selecionadas no filtro.</div>",
    unsafe_allow_html=True,
)

with st.sidebar:
    st.header("Filtros")

try:
    fonte_base = localizar_base()
    df_original = carregar_base(fonte_base)
except Exception as e:
    st.error(f"Erro ao carregar a base: {e}")
    st.stop()

opcoes_transportadoras = opcoes_ordenadas(
    pd.Series(["CLIENTE RETIRA", "CORREIOS"] + opcoes_ordenadas(df_original["Transportadora Normalizada"]))
)
default_transportadoras = [x for x in ["CLIENTE RETIRA", "CORREIOS"] if x in opcoes_transportadoras]

with st.sidebar:
    transportadoras_desconsiderar = st.multiselect(
        "Desconsiderar Transportadora",
        options=opcoes_transportadoras,
        default=default_transportadoras,
        help="Remove as transportadoras selecionadas. A regra considera quando o nome da transportadora contém o texto selecionado."
    )
    transportadoras_considerar = st.multiselect(
        "Considerar somente Transportadora",
        options=opcoes_transportadoras,
        default=[],
        help="Quando preenchido, mantém somente as transportadoras selecionadas. Se ficar vazio, considera todas as transportadoras não desconsideradas acima."
    )
    st.divider()

cenario = aplicar_regra_cenario(
    df_original,
    transportadoras_desconsiderar=transportadoras_desconsiderar,
    transportadoras_considerar=transportadoras_considerar,
)

with st.sidebar:
    unidades = st.multiselect("Unidade", opcoes_ordenadas(cenario["Unidade"]), default=[])
    datas = st.multiselect("Data NF", opcoes_ordenadas(cenario["Data"]), default=[])
    canais = st.multiselect("Canal", opcoes_ordenadas(cenario["Canal"]), default=[])
    subcanais = st.multiselect("Sub Canal", opcoes_ordenadas(cenario["Sub Canal"]), default=[])

filtrado = aplicar_filtros(cenario, unidades, datas, canais, subcanais)

# Base para o denominador "Total de NFs":
# considera as NFs da base filtrada, excluindo as transportadoras selecionadas pelo usuário,
# mas sem aplicar a regra de agrupamento por CNPJ.
base_nf_sem_cnpj = filtrar_transportadoras_validas(
    df_original,
    transportadoras_desconsiderar=transportadoras_desconsiderar,
)
base_nf_sem_cnpj = filtrar_transportadoras_consideradas(
    base_nf_sem_cnpj,
    transportadoras_considerar=transportadoras_considerar,
)
base_nf_sem_cnpj_filtrada = aplicar_filtros(base_nf_sem_cnpj, unidades, datas, canais, subcanais)

# Total de NFs geral usado apenas como referência da base filtrada.
total_nf_sem_considerar_cnpj = base_nf_sem_cnpj_filtrada["Nota Fiscal"].nunique()

# Total de NFs Agrupáveis: NFs que entraram no cenário de agrupamento.
total_nfs_agrupaveis = filtrado["Nota Fiscal"].nunique() if not filtrado.empty else 0
nfs_nao_agrupadas = int(filtrado["NFs não agrupadas"].sum()) if not filtrado.empty else 0

# % Ineficiência de Agrupamento:
# Total de NFs não agrupadas / Total de NFs Agrupáveis
perc_eficiencia_agrupamento = (
    nfs_nao_agrupadas / total_nfs_agrupaveis
) if total_nfs_agrupaveis else 0

# Cards principais ajustados:
# Excluídos os cards "Total de NFs Agrupáveis" e "Total NFs não agrupadas" da linha principal.
# Card 1: Total de Protocolos
# Card 2: Total Ideal de Protocolos
# Card 3: Total de Pedidos Gerados
c1, c2, c3 = st.columns(3)
c1.metric("Total de Protocolos", f"{filtrado['Protocolo'].nunique():,}".replace(",", "."))
c2.metric("Total Ideal de Protocolos", f"{filtrado['CNPJ'].nunique():,}".replace(",", "."))
c3.metric("Total de Pedidos Gerados", f"{filtrado['Pedido'].nunique():,}".replace(",", "."))

# Cards inferiores mantidos com nomes ajustados.
n1, n2, n3 = st.columns(3)
n1.metric("Total de NFs Agrupáveis", f"{total_nfs_agrupaveis:,}".replace(",", "."))
n2.metric("Total de NFs não agrupadas", f"{nfs_nao_agrupadas:,}".replace(",", "."))
n3.metric("% Ineficiência de Agrupamento", formatar_percentual(perc_eficiencia_agrupamento))

st.divider()

aba_unidade, aba_canal, aba_subcanal = st.tabs(["📍 Por Unidade", "🧭 Por Canal", "🏷️ Por Sub Canal"])
with aba_unidade:
    card_dataframe(filtrado, base_nf_sem_cnpj_filtrada, "Unidade", "Total de CNPJ por Unidade")
with aba_canal:
    card_dataframe(filtrado, base_nf_sem_cnpj_filtrada, "Canal", "Total de CNPJ por Canal")
with aba_subcanal:
    card_dataframe(filtrado, base_nf_sem_cnpj_filtrada, "Sub Canal", "Total de CNPJ por Sub Canal")

st.divider()
st.subheader("Relação de pedidos/protocolos no cenário")

relacao = filtrado[COLS_TABELA].drop_duplicates().sort_values(
    by=["Unidade", "Data do Protocolo", "CNPJ", "Pedido", "Protocolo"],
    ascending=True,
)
st.dataframe(relacao, use_container_width=True, hide_index=True)

exportar = filtrado[COLS_EXPORTAR].drop_duplicates().sort_values(
    by=["Unidade", "Data do Protocolo", "CNPJ", "Data", "Pedido", "Protocolo"]
)

st.download_button(
    label="⬇️ Baixar relação em Excel",
    data=converter_excel(exportar),
    file_name="relacao_nfs_mesmo_cnpj_protocolos_distintos.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True,
)

st.divider()
grafico_pizza_nfs_nao_agrupadas_por_unidade(filtrado)

st.divider()
tabela_ocorrencias(filtrado)

transportadoras_desconsiderar_txt = ", ".join(transportadoras_desconsiderar) if transportadoras_desconsiderar else "nenhuma"
transportadoras_considerar_txt = ", ".join(transportadoras_considerar) if transportadoras_considerar else "todas as transportadoras não desconsideradas"

with st.expander("Critério aplicado no cálculo"):
    st.markdown(
        f"""
        - Transportadoras desconsideradas: **{transportadoras_desconsiderar_txt}**;
        - Considerar somente transportadoras: **{transportadoras_considerar_txt}**;
        - primeiro o cálculo remove os registros selecionados em **Desconsiderar Transportadora**;
        - depois, se houver seleção em **Considerar somente Transportadora**, o cálculo mantém somente essas transportadoras;
        - mesmo **CNPJ**;
        - mesma **Data da Nota Fiscal**;
        - mesma **Unidade** como local de origem;
        - mais de um **Protocolo distinto** no agrupamento;
        - mais de uma **Nota Fiscal distinta** no agrupamento;
        - coluna **NFs não agrupadas**: dentro de cada grupo, a primeira NF localizada recebe **0** e as demais recebem **1**;
        - no gráfico de pizza, unidades com participação menor que **5%** são agrupadas como **OUTROS**;
        - card inferior **Total de NFs Agrupáveis**: conta as NFs distintas que entraram no cenário de agrupamento;
        - card inferior **% Ineficiência de Agrupamento**: divide o Total de NFs não agrupadas pelo Total de NFs Agrupáveis;
        - nas tabelas por Unidade, Canal e Sub Canal, as colunas exibidas são **Total de NFs Agrupáveis**, **Total de NFs não agrupadas** e **% Ineficiência de Agrupamento**;
        - o gráfico de pizza foi reduzido em **30%** no tamanho e nas fontes dos títulos/rótulos/percentuais.
        """
    )
