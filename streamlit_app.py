import streamlit as st
import pandas as pd
import re

st.set_page_config(
    page_title="Conversor TXT → Lovable",
    layout="wide"
)

st.title("Conversor de Relatórios para Lovable")

arquivo = st.file_uploader(
    "Selecione o arquivo TXT",
    type=["txt"]
)

if arquivo:

    texto = arquivo.read().decode(
        "latin-1",
        errors="ignore"
    )

    # =====================================================
    # ESTABELECIMENTO
    # =====================================================

    estabelecimento = ""

    m_est = re.search(
        r'Estabelecimento:\s*\d+\s*-\s*([A-Z0-9 ]+)',
        texto,
        re.IGNORECASE
    )

    if m_est:
        estabelecimento = m_est.group(1).strip()

    # =====================================================
    # PDV
    # =====================================================

    pdv_codigo = ""
    pdv_nome = ""

    m_pdv = re.search(
        r'(\d+)-PDV\d+-([A-Z0-9]+)',
        texto
    )

    if m_pdv:
        pdv_codigo = m_pdv.group(1).strip()
        pdv_nome = m_pdv.group(2).strip()

    # =====================================================
    # PROCESSAMENTO
    # =====================================================

    linhas = texto.splitlines()

    venda_atual = None
    nf_atual = None
    data_atual = None

    registros = []

    for linha in linhas:

        linha = linha.strip()

        # ==========================================
        # VENDA
        # ==========================================

        venda_match = re.search(
            r'Venda\/Sa(\d+)',
            linha
        )

        if venda_match:
            venda_atual = venda_match.group(1)

        # ==========================================
        # NF
        # ==========================================

        nf_match = re.search(
            r'(NC\d+\/\d+)',
            linha
        )

        if nf_match:
            nf_atual = nf_match.group(1)

        # ==========================================
        # DATA
        # ==========================================

        data_match = re.search(
            r'(\d{2}/\d{2}/\d{4})',
            linha
        )

        if data_match:
            data_atual = data_match.group(1)

        # ==========================================
        # ITEM
        # ==========================================

        item_match = re.search(
            r'(\d{6})(.+?)(\d+,\d+)UN\s+(\d+,\d+)\s+\d+,\d+\s+(\d+,\d+)',
            linha
        )

        if item_match:

            item_codigo = item_match.group(1)

            item_nome = (
                item_match.group(2)
                .replace("  ", " ")
                .strip()
            )

            quantidade = (
                item_match.group(3)
                .replace(".", "")
                .replace(",", ".")
            )

            valor_unitario = (
                item_match.group(4)
                .replace(".", "")
                .replace(",", ".")
            )

            valor_total = (
                item_match.group(5)
                .replace(".", "")
                .replace(",", ".")
            )

            registros.append({
                "estabelecimento": estabelecimento,
                "pdv_codigo": pdv_codigo,
                "pdv_nome": pdv_nome,
                "data_venda": data_atual,
                "venda": venda_atual,
                "nf": nf_atual,
                "item_codigo": item_codigo,
                "item_nome": item_nome,
                "quantidade": float(quantidade),
                "valor_unitario": float(valor_unitario),
                "valor_total": float(valor_total)
            })

    # =====================================================
    # DATAFRAME
    # =====================================================

    df = pd.DataFrame(registros)

    if df.empty:
        st.error("Nenhum registro encontrado.")
        st.stop()

    # =====================================================
    # CONVERTER DATA
    # =====================================================

    df["data_venda"] = pd.to_datetime(
        df["data_venda"],
        format="%d/%m/%Y",
        errors="coerce"
    )

    df["data_venda"] = df["data_venda"].dt.strftime(
        "%Y-%m-%d"
    )

    # =====================================================
    # ESTATÍSTICAS
    # =====================================================

    col1, col2, col3, col4 = st.columns(4)

    col1.metric(
        "Registros",
        len(df)
    )

    col2.metric(
        "Vendas",
        df["venda"].nunique()
    )

    col3.metric(
        "Itens",
        df["item_codigo"].nunique()
    )

    col4.metric(
        "Faturamento",
        f"R$ {df['valor_total'].sum():,.2f}"
    )

    # =====================================================
    # VISUALIZAÇÃO
    # =====================================================

    st.subheader("Prévia")

    st.dataframe(
        df.head(100),
        use_container_width=True
    )

    # =====================================================
    # DOWNLOAD CSV
    # =====================================================

    csv = df.to_csv(
        index=False,
        sep=";"
    )

    st.download_button(
        "Baixar CSV para Lovable",
        csv,
        file_name=f"{pdv_codigo}_{pdv_nome}.csv",
        mime="text/csv"
    )

    # =====================================================
    # DOWNLOAD XLSX
    # =====================================================

    try:

        from io import BytesIO

        output = BytesIO()

        with pd.ExcelWriter(
            output,
            engine="openpyxl"
        ) as writer:

            df.to_excel(
                writer,
                index=False,
                sheet_name="Vendas"
            )

        st.download_button(
            "Baixar XLSX para Lovable",
            output.getvalue(),
            file_name=f"{pdv_codigo}_{pdv_nome}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception:
        st.warning(
            "openpyxl não instalado. O CSV continua disponível normalmente."
        )
