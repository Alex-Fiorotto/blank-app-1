import streamlit as st
import pandas as pd
import re
from io import BytesIO

st.set_page_config(
    page_title="Conversor TXT -> Lovable",
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

    estabelecimento = ""

    m_est = re.search(
        r'Estabelecimento:\s*\d+\s*-\s*([^\n\r]+)',
        texto
    )

    if m_est:
        estabelecimento = m_est.group(1).strip()

    m_pdv = re.search(
        r'PDV:\s*(\d+)-PDV\d+-([^\n\r]+)',
        texto
    )

    pdv_codigo = ""
    pdv_nome = ""

    if m_pdv:
        pdv_codigo = m_pdv.group(1).strip()
        pdv_nome = m_pdv.group(2).strip()

    linhas = texto.splitlines()

    registros = []

    venda_atual = None
    nf_atual = None
    data_atual = None

    for linha in linhas:

        linha = linha.strip()

        venda_match = re.search(
            r'Venda\/Sa(\d+)',
            linha
        )

        if venda_match:
            venda_atual = venda_match.group(1)

        nf_match = re.search(
            r'(NC\d+\/\d+)',
            linha
        )

        if nf_match:
            nf_atual = nf_match.group(1)

        data_match = re.search(
            r'(\d{2}/\d{2}/\d{4})',
            linha
        )

        if data_match:
            data_atual = data_match.group(1)

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

    df = pd.DataFrame(registros)

    if not df.empty:

        st.success(
            f"{len(df)} registros encontrados"
        )

        st.dataframe(
            df.head(50),
            use_container_width=True
        )

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
            "Baixar XLSX",
            output.getvalue(),
            file_name="vendas_lovable.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    else:

        st.error(
            "Nenhum registro encontrado."
        )
