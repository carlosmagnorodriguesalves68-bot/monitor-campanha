import streamlit as st
import pandas as pd
import numpy as np
import unicodedata
import re

st.set_page_config(page_title="Monitor de Campanha", layout="wide")

# =====================================================
# CSS — VISUAL CAMPANHA PROTOP
# =====================================================

st.markdown("""
<style>
.block-container {
    padding-top: 2.8rem !important;
    padding-left: 1.2rem !important;
    padding-right: 1.2rem !important;
    max-width: 98% !important;
}

body, .stApp {
    background: linear-gradient(180deg, #0B2A55 0%, #103B73 180px, #F6F8FB 420px);
}

h1 {
    color: #FFFFFF !important;
    font-weight: 900 !important;
    margin-top: 0px !important;
    padding-top: 0px !important;
}

p, label {
    color: #111827;
}

div[data-testid="stCaptionContainer"] {
    color: #EAF6FF !important;
}

div[data-testid="stFileUploader"] {
    max-width: 430px !important;
    margin-top: -8px !important;
    margin-bottom: 10px !important;
}

div[data-testid="stFileUploaderDropzone"] {
    background: #FFFFFF !important;
    min-height: 55px !important;
    padding: 6px 10px !important;
    border-radius: 12px !important;
    border: 2px solid #22B8CF !important;
}

div[data-testid="stFileUploaderDropzone"] small {
    display: none !important;
}

[data-testid="stMetric"] {
    background: #FFFFFF;
    border: 1px solid #E5E7EB;
    border-top: 5px solid #FF7A00;
    border-radius: 16px;
    padding: 14px;
    box-shadow: 0 4px 12px rgba(0,0,0,0.12);
}

[data-testid="stMetricLabel"] {
    color: #0B2A55 !important;
    font-weight: 700 !important;
}

[data-testid="stMetricValue"] {
    color: #111827 !important;
    font-weight: 900 !important;
}

h2, h3 {
    color: #0B2A55 !important;
    font-weight: 900 !important;
    margin-top: 8px !important;
}

div.stButton > button {
    height: 38px;
    border-radius: 10px;
    font-weight: 700;
    background: #FF7A00;
    color: #FFFFFF;
    border: none;
}

div.stButton > button:hover {
    background: #E85F00;
    color: #FFFFFF;
}

[data-testid="stDataFrame"] {
    border-radius: 14px;
    border: 1px solid #DCE3EA;
    overflow: hidden;
}

hr {
    border-color: #22B8CF;
}

</style>
""", unsafe_allow_html=True)


# =====================================================
# FUNÇÕES BASE
# =====================================================

def normalizar(txt):
    if pd.isna(txt):
        txt = ""
    txt = str(txt).strip().lower()
    txt = unicodedata.normalize("NFKD", txt).encode("ASCII", "ignore").decode("utf-8")
    txt = re.sub(r"[^a-z0-9]+", "_", txt)
    return txt.strip("_")


def numero(v):
    try:
        if pd.isna(v):
            return 0.0
        if isinstance(v, (int, float, np.integer, np.floating)):
            return float(v)

        v = str(v).strip()
        v = v.replace("R$", "").replace("%", "").strip()

        if "," in v:
            v = v.replace(".", "").replace(",", ".")

        return float(v)
    except:
        return 0.0


def moeda(v):
    try:
        return f"R$ {float(v):,.0f}".replace(",", ".")
    except:
        return "R$ 0"


def pct(v):
    try:
        return f"{float(v):.1f}%"
    except:
        return "0.0%"


def achar_header(excel, aba, palavras):
    temp = pd.read_excel(excel, sheet_name=aba, header=None, nrows=80)

    for i in range(len(temp)):
        linha = " ".join(temp.iloc[i].dropna().astype(str).tolist()).upper()
        if all(p.upper() in linha for p in palavras):
            return i

    return 0


def ler_aba(excel, aba, palavras_header):
    linha_header = achar_header(excel, aba, palavras_header)

    df = pd.read_excel(
        excel,
        sheet_name=aba,
        header=linha_header
    )

    df = df.dropna(axis=0, how="all")
    df = df.dropna(axis=1, how="all")

    colunas = []
    contador = {}

    for c in df.columns:
        nome = normalizar(c)

        if nome == "":
            nome = "coluna"

        if nome not in contador:
            contador[nome] = 0
            colunas.append(nome)
        else:
            contador[nome] += 1
            colunas.append(f"{nome}_{contador[nome]}")

    df.columns = colunas

    return df


def localizar_aba(abas, termos):
    for aba in abas:
        nome = normalizar(aba)
        for termo in termos:
            if termo in nome:
                return aba
    return None


def pegar_coluna(df, opcoes):
    for op in opcoes:
        if op in df.columns:
            return op
    return None


def status_vb(falta_vb, perc_vb):
    if falta_vb <= 0:
        return "🟢 Fechada"
    elif perc_vb >= 85:
        return "🟡 Próxima"
    else:
        return "🔴 Crítica"


def status_pos(perc_pos):
    if perc_pos >= 100:
        return "🟢 Fechada"
    elif perc_pos >= 60:
        return "🟡 Próxima"
    else:
        return "🔴 Crítica"


def formatar_tabela(df):
    df = df.copy()

    for c in df.columns:
        if c in ["Meta VB", "Real VB", "Falta VB"]:
            df[c] = df[c].apply(moeda)

        if c in ["% VB", "% POS"]:
            df[c] = df[c].apply(pct)

        if c in ["Meta POS", "Real POS", "Falta POS"]:
            df[c] = df[c].astype(float).round(0).astype(int)

    return df


# =====================================================
# TOPO
# =====================================================

topo1, topo2 = st.columns([8, 1])

with topo1:
    st.title("🧭 Monitor de Campanha")
    st.caption("Central inteligente de decisão comercial")

with topo2:
    st.write("")
    if st.button("🧹 Limpar", use_container_width=True):
        for k in list(st.session_state.keys()):
            del st.session_state[k]
        st.rerun()


arquivo = st.file_uploader(
    "📁 Suba a planilha",
    type=["xlsx", "xls", "xlsb"]
)


# =====================================================
# APP
# =====================================================

if arquivo:

    try:
        engine = "pyxlsb" if arquivo.name.endswith(".xlsb") else None
        excel = pd.ExcelFile(arquivo, engine=engine)

        aba_st = "ST00841" if "ST00841" in excel.sheet_names else localizar_aba(excel.sheet_names, ["st00841"])
        aba_posit = localizar_aba(excel.sheet_names, ["posit"])

        if not aba_st or not aba_posit:
            st.error("❌ Não encontrei as abas ST00841 e Posit na planilha.")
            st.stop()

        df_st = ler_aba(excel, aba_st, ["LINHA", "NOME"])
        df_posit = ler_aba(excel, aba_posit, ["CODCLI", "CLIENTE"])

        col_linha = pegar_coluna(df_st, ["linha"])
        col_industria = pegar_coluna(df_st, ["industria"])
        col_familia = pegar_coluna(df_st, ["nome_linha", "familia", "linha_nome"])

        col_meta_vb = pegar_coluna(df_st, ["obj_total_vb", "meta_vb", "objetivo_vb"])
        col_real_vb = pegar_coluna(df_st, ["real_vb", "realizado_vb"])
        col_perc_vb = pegar_coluna(df_st, ["var", "var_vb", "perc_vb"])

        col_meta_pos = pegar_coluna(df_st, ["meta_pos_cs", "meta_pos"])
        col_real_pos = pegar_coluna(df_st, ["real_pos_cs", "real_pos"])
        col_perc_pos = pegar_coluna(df_st, ["var_pos", "perc_pos"])

        col_pont_vb_1 = pegar_coluna(df_st, ["98_5_a_99_9"])
        col_pont_vb_2 = pegar_coluna(df_st, ["100_acima"])
        col_pont_pos_1 = pegar_coluna(df_st, ["95_5_a_99_9"])
        col_pont_pos_2 = pegar_coluna(df_st, ["100_acima_1"])

        obrigatorias = [
            col_linha,
            col_industria,
            col_familia,
            col_meta_vb,
            col_real_vb,
            col_meta_pos,
            col_real_pos
        ]

        if any(c is None for c in obrigatorias):
            st.error("❌ Não encontrei todas as colunas principais da aba ST00841.")
            st.stop()

        df_st = df_st[df_st[col_linha].astype(str).str.startswith("A")].copy()

        for c in [
            col_meta_vb,
            col_real_vb,
            col_perc_vb,
            col_meta_pos,
            col_real_pos,
            col_perc_pos,
            col_pont_vb_1,
            col_pont_vb_2,
            col_pont_pos_1,
            col_pont_pos_2
        ]:
            if c:
                df_st[c] = df_st[c].apply(numero)

        if not col_perc_vb:
            df_st["perc_vb_calc"] = np.where(
                df_st[col_meta_vb] > 0,
                df_st[col_real_vb] / df_st[col_meta_vb] * 100,
                0
            )
            col_perc_vb = "perc_vb_calc"

        if not col_perc_pos:
            df_st["perc_pos_calc"] = np.where(
                df_st[col_meta_pos] > 0,
                df_st[col_real_pos] / df_st[col_meta_pos] * 100,
                0
            )
            col_perc_pos = "perc_pos_calc"

        df_st["Falta VB"] = (df_st[col_meta_vb] - df_st[col_real_vb]).clip(lower=0)
        df_st["Falta POS"] = (df_st[col_meta_pos] - df_st[col_real_pos]).clip(lower=0)

        df_st["Status VB"] = df_st.apply(
            lambda x: status_vb(x["Falta VB"], x[col_perc_vb]),
            axis=1
        )

        df_st["Status POS"] = df_st[col_perc_pos].apply(status_pos)

        base = pd.DataFrame({
            "Código Linha": df_st[col_linha].astype(str),
            "Indústria": df_st[col_industria].astype(str),
            "Família": df_st[col_familia].astype(str),
            "Meta VB": df_st[col_meta_vb],
            "Real VB": df_st[col_real_vb],
            "Falta VB": df_st["Falta VB"],
            "% VB": df_st[col_perc_vb],
            "Meta POS": df_st[col_meta_pos],
            "Real POS": df_st[col_real_pos],
            "Falta POS": df_st["Falta POS"],
            "% POS": df_st[col_perc_pos],
            "Status VB": df_st["Status VB"],
            "Status POS": df_st["Status POS"]
        })

        # =====================================================
        # CARDS
        # =====================================================

        meta_vb = base["Meta VB"].sum()
        real_vb = base["Real VB"].sum()
        falta_vb = base["Falta VB"].sum()

        pont_vb = 0
        pont_pos = 0

        if col_pont_vb_1:
            pont_vb += df_st[col_pont_vb_1].sum()
        if col_pont_vb_2:
            pont_vb += df_st[col_pont_vb_2].sum()
        if col_pont_pos_1:
            pont_pos += df_st[col_pont_pos_1].sum()
        if col_pont_pos_2:
            pont_pos += df_st[col_pont_pos_2].sum()

        c1, c2, c3, c4 = st.columns(4)

        with c1:
            st.metric("Meta VB", moeda(meta_vb))

        with c2:
            st.metric("Real VB", moeda(real_vb))

        with c3:
            st.metric("Falta VB", moeda(falta_vb))

        with c4:
            st.metric("Pontuação Total", int(pont_vb + pont_pos))

        # =====================================================
        # FILTRO GLOBAL DE FAMÍLIA
        # =====================================================

        familias = ["Todas"] + sorted(base["Família"].dropna().astype(str).unique())

        familia_filtro = st.selectbox(
            "Filtrar Família",
            familias
        )

        if familia_filtro != "Todas":
            base_filtrada = base[base["Família"] == familia_filtro].copy()
        else:
            base_filtrada = base.copy()

        # =====================================================
        # QUADROS PRINCIPAIS
        # =====================================================

        col1, col2 = st.columns(2)

        with col1:
            st.subheader("🎯 POS mais próximo")

            ranking_pos = base_filtrada[
                base_filtrada["Falta POS"] > 0
            ][
                [
                    "Indústria",
                    "Família",
                    "Meta POS",
                    "Real POS",
                    "Falta POS",
                    "% POS",
                    "Status POS",
                    "Status VB"
                ]
            ].sort_values(
                ["Falta POS", "% POS"],
                ascending=[True, False]
            )

            st.dataframe(
                formatar_tabela(ranking_pos),
                use_container_width=True,
                hide_index=True,
                height=360
            )

        with col2:
            st.subheader("💰 VB mais próximo")

            ranking_vb = base_filtrada[
                base_filtrada["Falta VB"] > 0
            ][
                [
                    "Indústria",
                    "Família",
                    "Meta VB",
                    "Real VB",
                    "Falta VB",
                    "% VB",
                    "Status VB",
                    "Status POS"
                ]
            ].sort_values(
                ["Falta VB", "% VB"],
                ascending=[True, False]
            )

            st.dataframe(
                formatar_tabela(ranking_vb),
                use_container_width=True,
                hide_index=True,
                height=360
            )

        # =====================================================
        # RAIO-X DO CLIENTE
        # =====================================================

        st.subheader("🔎 RAIO-X DO CLIENTE")
        st.caption("Selecione um cliente e veja se comprou ou não a família selecionada.")

        col_cliente = pegar_coluna(df_posit, ["cliente"])

        if not col_cliente:
            st.warning("Não encontrei a coluna CLIENTE na aba Posit.")
            st.stop()

        cliente_escolhido = st.selectbox(
            "Selecionar Cliente",
            sorted(df_posit[col_cliente].dropna().astype(str).unique())
        )

        filtro_status = st.selectbox(
            "Filtro Cliente",
            ["Todos", "Falta vender", "Já positivou"]
        )

        cliente_df = df_posit[
            df_posit[col_cliente].astype(str) == cliente_escolhido
        ]

        if not cliente_df.empty:

            cliente_row = cliente_df.iloc[0]

            base_raiox = base_filtrada.copy()

            linhas = []

            for _, fam in base_raiox.iterrows():

                codigo = normalizar(fam["Código Linha"])
                familia = normalizar(fam["Família"])

                coluna_posit = None

                for c in df_posit.columns:
                    c_norm = normalizar(c)

                    if c_norm == codigo or familia in c_norm or c_norm in familia:
                        coluna_posit = c
                        break

                positivou = 0

                if coluna_posit:
                    try:
                        positivou = int(pd.to_numeric(cliente_row[coluna_posit], errors="coerce"))
                    except:
                        positivou = 0

                status_cliente = "Já positivou" if positivou == 1 else "Falta vender"
                checklist = "✅" if positivou == 1 else "⬜"

                if positivou == 1:
                    acao = "Não priorizar agora"
                elif fam["Status POS"] == "🟡 Próxima" and fam["Status VB"] == "🟡 Próxima":
                    acao = "Alta prioridade: pode ajudar VB e POS"
                elif fam["Status POS"] == "🟡 Próxima":
                    acao = "Ofertar item de giro para fechar POS"
                elif fam["Status VB"] == "🟡 Próxima":
                    acao = "Ofertar pedido para fechar VB"
                else:
                    acao = "Baixa prioridade"

                linhas.append({
                    "Checklist": checklist,
                    "Família": fam["Família"],
                    "Status Cliente": status_cliente,
                    "Status VB": fam["Status VB"],
                    "Status POS": fam["Status POS"],
                    "Falta VB": fam["Falta VB"],
                    "Falta POS": fam["Falta POS"],
                    "Ação sugerida": acao
                })

            df_raiox = pd.DataFrame(linhas)

            if filtro_status != "Todos":
                df_raiox = df_raiox[df_raiox["Status Cliente"] == filtro_status]

            ordem = {
                "Alta prioridade: pode ajudar VB e POS": 1,
                "Ofertar item de giro para fechar POS": 2,
                "Ofertar pedido para fechar VB": 3,
                "Baixa prioridade": 4,
                "Não priorizar agora": 5
            }

            df_raiox["ordem"] = df_raiox["Ação sugerida"].map(ordem)

            df_raiox = df_raiox.sort_values(
                ["ordem", "Falta POS", "Falta VB"],
                ascending=[True, True, True]
            ).drop(columns=["ordem"])

            col_whats1, col_whats2 = st.columns([4, 1])

            with col_whats2:
                gerar_whats = st.button(
                    "📲 WhatsApp",
                    use_container_width=True
                )

            if gerar_whats:

                df_faltantes = df_raiox[
                    df_raiox["Status Cliente"] == "Falta vender"
                ]

                texto = f"""🚀 CHECKLIST DA VISITA

👤 Cliente:
{cliente_escolhido}

━━━━━━━━━━━━━━

🎯 FAMÍLIAS PARA OFERTAR
"""

                for _, row in df_faltantes.head(8).iterrows():
                    texto += f"""
☐ {row['Família']}
💰 VB: {moeda(row['Falta VB'])}
🎯 POS: {int(row['Falta POS'])}
📌 {row['Ação sugerida']}

"""

                texto += """━━━━━━━━━━━━━━

✅ Foco
"""

                st.code(texto, language=None)
                st.caption("📲 Clique na folhinha para copiar e colar no WhatsApp.")

            st.dataframe(
                formatar_tabela(df_raiox),
                use_container_width=True,
                hide_index=True,
                height=520
            )

    except Exception as e:
        st.error(f"Erro: {e}")

else:
    st.info("📁 Suba a planilha para iniciar.")