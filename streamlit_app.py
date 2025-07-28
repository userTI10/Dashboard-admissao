import streamlit as st
import requests
import io
import xlsxwriter
from collections import defaultdict
import pandas as pd
from dotenv import load_dotenv
import os

# Configuração da página
st.set_page_config(page_title="📋 Processos Holmes", layout="wide")

# Cabeçalho com logo
col1, col2 = st.columns([1, 10])
with col1:
    st.image("logo-black.webp", width=300)
with col2:
    st.title("📄 Processos Abertos - Holmes")

# Token e URL
load_dotenv()
API_TOKEN = os.getenv("API_TOKEN")
API_URL = "https://app-api.holmesdoc.io/v2/search"

# Tamanho da página
tamanho_pagina = 100
pagina_atual = st.number_input("Página", min_value=1, value=1, step=1)

# Função genérica para buscar processos por status
def buscar_processos(status_nome):
    payload = {
        "query": {
            "from": (pagina_atual - 1) * tamanho_pagina,
            "size": tamanho_pagina,
            "context": "process",
            "sort": "updated_at",
            "order": "desc",
            "groups": [
                {
                    "match_all": True,
                    "terms": [
                        {
                            "value": "680b719537846a536ec8df4d",
                            "type": "is",
                            "field": "template_id"
                        },
                        {
                            "field": "status",
                            "type": "is",
                            "value": status_nome,
                            "nested": False
                        }
                    ]
                }
            ]
        },
        "trash": False,
        "deleted_by_me": False,
        "api_token": API_TOKEN
    }

    headers = {
        "Content-Type": "application/json"
    }

    response = requests.post(API_URL, headers=headers, json=payload)
    response.raise_for_status()
    data = response.json()

    registros = []
    total_vagas = 0

    for processo in data.get("docs", []):
        props = {p.get("identifier"): p for p in processo.get("props", [])}

        titulo = props.get("titulo", {}).get("value", "")
        solicitante = props.get("nome_do_solicitante", {}).get("value", "")
        tipo_vaga = props.get("tipo_de_vaga", {}).get("label", "")
        razao_social = props.get("razao_social", {}).get("value", "")
        vagas_str = props.get("numero_de_vagas", {}).get("value", "0")

        try:
            vagas = int(vagas_str)
        except ValueError:
            vagas = 0

        total_vagas += vagas

        registros.append({
            "Solicitação": processo.get("identifier", ""),
            "Título": titulo,
            "Solicitante": solicitante,
            "Tipo de Vaga": tipo_vaga,
            "Razão Social": razao_social,
            "Vagas": vagas
        })

    return registros, total_vagas

# === Processos Abertos ===
try:
    registros_abertos, total_vagas_abertas = buscar_processos("opened")

    st.subheader("✅ Processos Abertos")

    termo_busca = st.text_input("🔍 Buscar por Solicitação, Título, Solicitante ou Tipo de Vaga:")

    if termo_busca:
        termo_busca = termo_busca.lower()
        registros_abertos = [
            r for r in registros_abertos if
            termo_busca in r["Solicitação"].lower()
            or termo_busca in r["Título"].lower()
            or termo_busca in r["Solicitante"].lower()
            or termo_busca in r["Tipo de Vaga"].lower()
        ]

    total_filtrado = sum(r["Vagas"] for r in registros_abertos)

    st.markdown(f"<h4>🧮 Total de Vagas Encontradas: <span style='color:#2E86AB;'>{total_filtrado}</span></h4>", unsafe_allow_html=True)
    st.dataframe(registros_abertos, use_container_width=True)

    # Excel
    if registros_abertos:
        xlsx_buffer = io.BytesIO()
        workbook = xlsxwriter.Workbook(xlsx_buffer, {'in_memory': True})
        worksheet = workbook.add_worksheet("Abertos")
        headers = list(registros_abertos[0].keys())
        for col_num, header in enumerate(headers):
            worksheet.write(0, col_num, header)
        for row_num, row in enumerate(registros_abertos, start=1):
            for col_num, key in enumerate(headers):
                worksheet.write(row_num, col_num, row[key])
        workbook.close()
        xlsx_buffer.seek(0)
        st.download_button("📥 Baixar Excel (Abertos)", xlsx_buffer, "processos_abertos.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        # Ranking Abertos
        ranking = defaultdict(int)
        for r in registros_abertos:
            ranking[r["Solicitante"]] += r["Vagas"]
        ranking_ordenado = sorted(ranking.items(), key=lambda x: x[1], reverse=True)
        df_ranking = pd.DataFrame(ranking_ordenado, columns=["Solicitante", "Total de Vagas"])
        st.markdown("### 🏆 Ranking de Solicitantes por Vagas Abertas")
        st.dataframe(df_ranking, use_container_width=True)

except requests.exceptions.RequestException as e:
    st.error(f"❌ Erro ao buscar processos abertos: {e}")

# === Processos Cancelados ===
try:
    registros_cancelados, total_vagas_canceladas = buscar_processos("canceled")

    st.subheader("❌ Processos Cancelados")
    st.markdown(f"<h4>🧮 Total de Vagas Canceladas: <span style='color:#CB4335;'>{total_vagas_canceladas}</span></h4>", unsafe_allow_html=True)
    st.dataframe(registros_cancelados, use_container_width=True)

    if registros_cancelados:
        xlsx_buffer = io.BytesIO()
        workbook = xlsxwriter.Workbook(xlsx_buffer, {'in_memory': True})
        worksheet = workbook.add_worksheet("Cancelados")
        headers = list(registros_cancelados[0].keys())
        for col_num, header in enumerate(headers):
            worksheet.write(0, col_num, header)
        for row_num, row in enumerate(registros_cancelados, start=1):
            for col_num, key in enumerate(headers):
                worksheet.write(row_num, col_num, row[key])
        workbook.close()
        xlsx_buffer.seek(0)
        st.download_button("📥 Baixar Excel (Cancelados)", xlsx_buffer, "processos_cancelados.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        # Ranking Cancelados
        ranking_cancelados = defaultdict(int)
        for r in registros_cancelados:
            ranking_cancelados[r["Solicitante"]] += r["Vagas"]
        ranking_ordenado = sorted(ranking_cancelados.items(), key=lambda x: x[1], reverse=True)
        df_ranking_cancelados = pd.DataFrame(ranking_ordenado, columns=["Solicitante", "Total de Vagas Canceladas"])
        st.markdown("### 🏴 Ranking de Solicitantes com Mais Vagas Canceladas")
        st.dataframe(df_ranking_cancelados, use_container_width=True)

except requests.exceptions.RequestException as e:
    st.error(f"❌ Erro ao buscar processos cancelados: {e}")

