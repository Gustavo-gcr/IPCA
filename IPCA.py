import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import os
import subprocess
import sys

# Função para carregar dados do Excel
def carregar_dados_excel(nome_arquivo):
    try:
        caminho_arquivo = os.path.join(os.path.dirname(__file__), nome_arquivo)
        df = pd.read_excel(caminho_arquivo)
        df['dia'] = pd.to_datetime(df['dia'], format='%d/%m/%Y', errors='coerce')
        return df
    except Exception as e:
        st.write(f"Erro ao carregar a planilha: {e}")
        return None

# Função para obter o último mês preenchido
def obter_ultimo_mes(df):
    if df is not None and 'TAXA DIA' in df.columns:
        df_preenchido = df.dropna(subset=['TAXA DIA'])
        if not df_preenchido.empty:
            ultimo_mes = df_preenchido['dia'].max()
            return ultimo_mes.strftime('%m/%Y') if pd.notnull(ultimo_mes) else "Desconhecido"
    return "Desconhecido"

# Configuração para páginas no Streamlit
st.sidebar.title("Navegação")
pagina_selecionada = st.sidebar.radio("Selecione a Página", ["Calculadora", "Atualizar Planilha"])

if pagina_selecionada == "Calculadora":
    st.title("Calculadora de Juros IPCA")
    dados_ipca = carregar_dados_excel('IPCA-Teste.xlsx')
    ultimo_mes = obter_ultimo_mes(dados_ipca)
    st.write(f"Último mês preenchido na planilha: {ultimo_mes}")
    
    if dados_ipca is not None:
        with st.form("calculo_juros_form"):
            valor_input = st.number_input("Digite o valor", min_value=0.0)
            hoje = datetime.today()
            data_minima = datetime(2016, 6, 30)
            data_maxima = hoje.replace(day=1) - timedelta(days=1)
            data_input = st.date_input("Selecione a data (Dia, Mês, Ano)", min_value=data_minima, max_value=data_maxima)
            confirmar_calculo = st.form_submit_button("Confirmar")
            
            if confirmar_calculo and data_input and valor_input > 0:
                data_selecionada_str = data_input.strftime('%Y/%m/%d')
                if 'dia' in dados_ipca.columns and 'TotalPorcentagem' in dados_ipca.columns:
                    dados_ipca['dia'] = pd.to_datetime(dados_ipca['dia'], format='%d/%m/%Y', errors='coerce')
                    linha_dados = dados_ipca[dados_ipca['dia'] == pd.to_datetime(data_selecionada_str)]
                else:
                    st.write("Colunas 'dia' ou 'TotalPorcentagem' não encontradas.")
                    st.stop()
                
                if not linha_dados.empty:
                    taxa_porcentagem = linha_dados['TotalPorcentagem'].values[0]
                    if pd.isna(taxa_porcentagem):
                        st.write(f"Taxa de juros para a data {data_selecionada_str} está indisponível.")
                    else:
                        if isinstance(taxa_porcentagem, str):
                            taxa_porcentagem = taxa_porcentagem.replace('%', '').replace(',', '.')
                            taxa_porcentagem = float(taxa_porcentagem)
                        st.write(f"Taxa de juros para dia {data_selecionada_str}: {taxa_porcentagem:.2f}%")
                        valor_ajustado = valor_input * (taxa_porcentagem / 100)
                        st.write(f"Valor ajustado: R$ {valor_ajustado:.2f}")
                else:
                    st.write(f"Não há dados disponíveis para a data {data_input}.")
    else:
        st.write("Erro ao carregar os dados do Excel.")

elif pagina_selecionada == "Atualizar Planilha":
    st.title("Atualizar Planilha IPCA")
    with st.form("atualizar_planilha_form"):
        mes = st.number_input("Digite o mês anterior (em número, ex: 09 para Setembro):", min_value=1, max_value=12, format="%02d")
        ano = st.number_input("Digite o ano atual (ex: 2024):", min_value=2024, max_value=datetime.today().year + 10)
        confirmar_atualizacao = st.form_submit_button("Confirmar Data Atualizada")

        if confirmar_atualizacao:
            if mes and ano:
                with st.spinner("Carregando..."):
                    comando = ["python", "atualizar.py", f"{mes:02d}", str(ano)]
                    try:
                        resultado = subprocess.run(comando, capture_output=True, text=True, check=True)
                        dados_ipca = carregar_dados_excel('IPCA-Teste.xlsx')
                        ultimo_mes = obter_ultimo_mes(dados_ipca)
                        st.write(f"Último mês preenchido na planilha: {ultimo_mes}")
                        st.write("Planilha atualizada com sucesso!")
                    except subprocess.CalledProcessError as e:
                        st.write("Erro ao atualizar a planilha.")
                        st.write(e.stderr)
