import streamlit as st
import pandas as pd
import numpy as np
import unicodedata
import re

st.set_page_config(page_title="Monitor de Campanha", layout="wide")

st.markdown("""
<style>
.block-container {padding-top: 0.5rem; padding-bottom: 1rem;}
body, .stApp {background-color: #F6F8FB;}
h1 {margin-bottom: 0px;}

div[data-testid="stFileUploader"] section {
    padding: 6px !important;
    min-height: 55px !important;
}

.card {
    background: #FFFFFF;
    padding: 12px;
    border-radius: 14px;
    border: 1px solid #E5E7EB;
    box-shadow: 0 2px 8px rgba(0,0,0,0.04);
}
.card-title {font-size: 12px; color: #6B7280; font-weight: 600;}
.card-value {font-size: 24px; font-weight: 800; color: #111827;}

.section-title {
    font-size: 20px;
    font-weight: 800;
    color: #111827;
    margin-top: 10px;
    margin-bottom: 8px;
}

.raiox-box {
    background: white;
    border-radius: 16px;
    padding: 14px;
    border: 1px solid #E5E7EB;
    box-shadow: 0 2px 8px rgba(0,0,0,0.04);
    margin-top: 16px;
    margin-bottom: 8px;
}
.raiox-title {font-size: 22px; font-weight: 800; color: #111827;}
.raiox-sub {font-size: 13px; color: #6B7280;}

div.stButton > button {
    height: 42px;
    font-size: 13px;
    border-radius: 12px;
}
</style>
""", unsafe_allow_html=True)


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
            return 0
        if isinstance(v, (int, float, np.integer, np.floating)):
            return float(v)
        v = str(v).strip()
        if "," in v:
            v = v.replace(".", "").replace(",", ".")
        return float(v)
    except:
        return 0


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


def ler_aba(excel, aba, palavras):
    linha = achar_header(excel, aba, palavras)
    df = pd.read_excel(excel, sheet_name=aba, header=linha)
    df = df.dropna(axis=0, how="all")
    df = df.dropna(axis=1, how="all")

    cols = [normalizar(c) for c in df.columns]
    novas = []
    contador = {}

    for c in cols:
        if c == "":
            c = "coluna"

        if c not in contador:
            contador[c] = 0
            novas.append(c)
        else:
            contador[c] += 1
            novas.append(f"{c}_{contador[c]}")

    df.columns = novas
    df = df.loc[:, ~df.columns.astype(str).str.contains("^unnamed")]
    return df


def localizar_aba(abas, palavras):
    for aba in abas:
        nome = normalizar(aba)
        for p in palavras:
            if p in nome:
                return aba
    return None


def col(df, nome):
    return nome if nome in df.columns else None


def status_vb(falta_vb, percentual_vb):
    if falta_vb <= 0:
        return "🟢 Fechada"
    elif percentual_vb >= 85:
        return "🟡 Próxima"
    else:
        return "🔴 Crítica"


def status_pos(p):
    if p >= 100:
        return "🟢 Fechada"
    elif p >= 60:
        return "🟡 Próxima"
    else:
        return "🔴 Crítica"


def card(titulo, valor):
    st.markdown(
        f"""
        <div class="card">
            <div class="card-title">{titulo}</div>
            <div class="card-value">{valor}</div>
        </div>
        """,
        unsafe_allow_html=True
    )


def formatar_tabela(df):
    df = df.copy()
    for c in df.columns:
        if "VB" in c and "%" not in c and df[c].dtype != object:
            df[c] = df[c].apply(moeda)
        if "%" in c and df[c].dtype != object:
            df[c] = df[c].apply(pct)
    return df


st.title("🧭 Monitor de Campanha")
st.caption("Central inteligente de decisão comercial")

arquivo = st.file_uploader(
    "📁 Suba a planilha",
    type=["xlsx", "xls", "xlsb"]
)

if arquivo:
    try:
        engine = "pyxlsb" if arquivo.name.endswith(".xlsb") else None
        excel = pd.ExcelFile(arquivo, engine=engine)

        aba_posit = localizar_aba(excel.sheet_names, ["posit"])
        aba_st = "ST00841" if "ST00841" in excel.sheet_names else localizar_aba(excel.sheet_names, ["st00841"])
        aba_cota = localizar_aba(excel.sheet_names, ["cota"])

        if not aba_posit or not aba_st:
            st.error("❌ Não encontrei as abas principais da planilha.")
            st.stop()

        df_posit = ler_aba(excel, aba_posit, ["CODCLI", "CLIENTE"])
        df_st = ler_aba(excel, aba_st, ["LINHA", "NOME"])

        if aba_cota:
            df_cota = ler_aba(excel, aba_cota, ["LINHA", "PRODUTO"])
        else:
            df_cota = pd.DataFrame()

        col_linha = col(df_st, "linha")
        col_ind = col(df_st, "industria")
        col_nome = col(df_st, "nome_linha")

        col_meta_vb = col(df_st, "obj_total_vb")
        col_real_vb = col(df_st, "real_vb")
        col_perc_vb = col(df_st, "var")

        col_meta_pos = col(df_st, "meta_pos_cs")
        col_real_pos = col(df_st, "real_pos_cs")
        col_perc_pos = col(df_st, "var_pos")

        col_pont_vb_1 = col(df_st, "98_5_a_99_9")
        col_pont_vb_2 = col(df_st, "100_acima")
        col_pont_pos_1 = col(df_st, "95_5_a_99_9")
        col_pont_pos_2 = col(df_st, "100_acima_1")

        obrigatorias = [
            col_linha, col_ind, col_nome,
            col_meta_vb, col_real_vb,
            col_meta_pos, col_real_pos
        ]

        if any(c is None for c in obrigatorias):
            st.error("❌ Não encontrei as colunas principais da aba ST00841.")
            st.stop()

        df_st = df_st[df_st[col_linha].astype(str).str.startswith("A")].copy()

        for c in [
            col_meta_vb, col_real_vb, col_perc_vb,
            col_meta_pos, col_real_pos, col_perc_pos,
            col_pont_vb_1, col_pont_vb_2, col_pont_pos_1, col_pont_pos_2
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

        df_st["falta_vb"] = (df_st[col_meta_vb] - df_st[col_real_vb]).clip(lower=0)
        df_st["falta_pos"] = (df_st[col_meta_pos] - df_st[col_real_pos]).clip(lower=0)

        df_st["status_vb"] = df_st.apply(
            lambda x: status_vb(x["falta_vb"], x[col_perc_vb]),
            axis=1
        )

        df_st["status_pos"] = df_st[col_perc_pos].apply(status_pos)

        meta_vb = df_st[col_meta_vb].sum()
        real_vb = df_st[col_real_vb].sum()
        falta_vb = df_st["falta_vb"].sum()

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

        k1, k2, k3, k4 = st.columns(4)

        with k1:
            card("Meta VB", moeda(meta_vb))
        with k2:
            card("Real VB", moeda(real_vb))
        with k3:
            card("Falta VB", moeda(falta_vb))
        with k4:
            card("Pontuação Total", int(pont_vb + pont_pos))

        raio_x = df_st[
            [
                col_nome,
                col_ind,
                col_meta_vb,
                col_real_vb,
                "falta_vb",
                col_perc_vb,
                col_meta_pos,
                col_real_pos,
                "falta_pos",
                col_perc_pos,
                "status_vb",
                "status_pos"
            ]
        ].copy()

        raio_x.columns = [
            "Família",
            "Indústria",
            "Meta VB",
            "Real VB",
            "Falta VB",
            "% VB",
            "Meta POS",
            "Real POS",
            "Falta POS",
            "% POS",
            "Status VB",
            "Status POS"
        ]

        colA, colB = st.columns(2)

        with colA:
            st.markdown('<div class="section-title">🎯 POS mais próximo</div>', unsafe_allow_html=True)

            ranking_pos = raio_x[
                raio_x["Falta POS"] > 0
            ][
                [
                    "Família",
                    "Indústria",
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

        with colB:
            st.markdown('<div class="section-title">💰 VB mais próximo</div>', unsafe_allow_html=True)

            ranking_vb = raio_x[
                raio_x["Falta VB"] > 0
            ][
                [
                    "Família",
                    "Indústria",
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

        col_cliente = "cliente" if "cliente" in df_posit.columns else None

        col_rx1, col_rx2 = st.columns([5, 1])

        with col_rx1:
            st.markdown("""
            <div class="raiox-box">
                <div class="raiox-title">🔎 RAIO-X DO CLIENTE</div>
                <div class="raiox-sub">
                    Selecione um cliente e veja o que já positivou e o que ainda falta vender.
                </div>
            </div>
            """, unsafe_allow_html=True)

        with col_rx2:
            st.write("")
            gerar_checklist = st.button(
                "📲 WhatsApp",
                use_container_width=True
            )

        if col_cliente:
            filtro1, filtro2 = st.columns([3, 1])

            with filtro1:
                cliente_escolhido = st.selectbox(
                    "Selecionar Cliente",
                    sorted(df_posit[col_cliente].dropna().astype(str).unique())
                )

            with filtro2:
                filtro_status = st.selectbox(
                    "Filtro",
                    ["Todos", "Falta vender", "Já positivou"]
                )

            cliente_df = df_posit[
                df_posit[col_cliente].astype(str) == cliente_escolhido
            ]

            if not cliente_df.empty:
                cliente_row = cliente_df.iloc[0]
                linhas_raiox = []

                for _, fam in raio_x.iterrows():
                    nome_familia = str(fam["Família"]).strip()
                    coluna_posit = None

                    for c in df_posit.columns:
                        if normalizar(nome_familia) in normalizar(c):
                            coluna_posit = c
                            break

                    positivou = 0

                    if coluna_posit:
                        try:
                            positivou = int(pd.to_numeric(cliente_row[coluna_posit], errors="coerce"))
                        except:
                            positivou = 0

                    checklist = "✅" if positivou == 1 else "⬜"
                    status_cliente = "Já positivou" if positivou == 1 else "Falta vender"

                    status_vb_atual = fam["Status VB"]
                    status_pos_atual = fam["Status POS"]

                    if positivou == 1:
                        acao = "Não priorizar agora"
                    elif status_pos_atual == "🟡 Próxima" and status_vb_atual == "🟡 Próxima":
                        acao = "Alta prioridade: pode ajudar VB e POS"
                    elif status_pos_atual == "🟡 Próxima":
                        acao = "Ofertar item de giro para fechar POS"
                    elif status_vb_atual == "🟡 Próxima":
                        acao = "Ofertar pedido para fechar VB"
                    else:
                        acao = "Baixa prioridade"

                    linhas_raiox.append({
                        "Checklist": checklist,
                        "Família": nome_familia,
                        "Status Cliente": status_cliente,
                        "Status VB": status_vb_atual,
                        "Status POS": status_pos_atual,
                        "Falta VB": fam["Falta VB"],
                        "Falta POS": fam["Falta POS"],
                        "Ação sugerida": acao
                    })

                df_raiox = pd.DataFrame(linhas_raiox)

                if filtro_status != "Todos":
                    df_raiox = df_raiox[
                        df_raiox["Status Cliente"] == filtro_status
                    ]

                prioridade = {
                    "Alta prioridade: pode ajudar VB e POS": 1,
                    "Ofertar item de giro para fechar POS": 2,
                    "Ofertar pedido para fechar VB": 3,
                    "Baixa prioridade": 4,
                    "Não priorizar agora": 5
                }

                df_raiox["ordem"] = df_raiox["Ação sugerida"].map(prioridade)

                df_raiox = df_raiox.sort_values(
                    ["ordem", "Falta POS", "Falta VB"],
                    ascending=[True, True, True]
                ).drop(columns=["ordem"])

                if gerar_checklist:
                    df_faltantes = df_raiox[
                        df_raiox["Status Cliente"] == "Falta vender"
                    ].copy()

                    texto_whats = f"""🚀 CHECKLIST DA VISITA

👤 Cliente:
{cliente_escolhido}

━━━━━━━━━━━━━━

🎯 FAMÍLIAS PARA OFERTAR
"""

                    for _, row in df_faltantes.head(8).iterrows():
                        texto_whats += f"""
☐ {row['Família']}
💰 VB: {moeda(row['Falta VB'])}
🎯 POS: {int(row['Falta POS'])}
📌 {row['Ação sugerida']}

"""

                    texto_whats += """━━━━━━━━━━━━━━

✅ Foco:
• itens de giro
• mix rápido
• pedido complementar
"""

                    st.code(
                        texto_whats,
                        language=None
                    )

                    st.caption(
                        "📲 Clique na folhinha no canto direito para copiar e colar no WhatsApp"
                    )

                df_raiox_formatado = formatar_tabela(df_raiox)

                st.dataframe(
                    df_raiox_formatado,
                    use_container_width=True,
                    hide_index=True,
                    height=520
                )

    except Exception as e:
        st.error(f"Erro: {e}")

else:
    st.info("📁 Suba a planilha para iniciar.")