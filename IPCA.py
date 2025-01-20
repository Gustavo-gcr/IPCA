import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import requests
import calendar
import os
import time

# Função para carregar dados do Excel
def carregar_dados_excel(nome_arquivo):
    try:
        caminho_arquivo = os.path.join(os.path.dirname(__file__), nome_arquivo)
        df = pd.read_excel(caminho_arquivo, engine='openpyxl')
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

# Função para preencher a coluna 'dia' com datas de 01/06/2016 até 10 anos após a data final
def preencher_coluna_dia(mes_fim, ano_fim):
    caminho_arquivo = 'IPCA-Teste.xlsx'
    if not os.path.exists(caminho_arquivo):
        st.write("Erro: Arquivo IPCA-Teste.xlsx não encontrado.")
        return

    data_inicial = datetime(2016, 6, 30)
    data_final = datetime(ano_fim + 10, mes_fim, 1) - timedelta(days=1)
    todas_as_datas = pd.date_range(start=data_inicial, end=data_final, freq='D')
    
    planilha = pd.read_excel(caminho_arquivo, engine='openpyxl')
    num_linhas = len(planilha)
    
    if len(todas_as_datas) > num_linhas:
        todas_as_datas = todas_as_datas[:num_linhas]
    elif len(todas_as_datas) < num_linhas:
        todas_as_datas = todas_as_datas.append(pd.Series([''] * (num_linhas - len(todas_as_datas))))
    
    planilha['dia'] = todas_as_datas.strftime('%d/%m/%Y')
    planilha.to_excel(caminho_arquivo, index=False)
  #  st.write("Coluna 'dia' preenchida com todas as datas de 30/06/2016 até 10 anos após o mês final informado.")

# Função para buscar o IPCA do Banco Central
def buscar_ipca(mes, ano, tentativas=3, timeout=10):
    ultimo_dia = calendar.monthrange(ano, mes)[1]
    url = f"https://api.bcb.gov.br/dados/serie/bcdata.sgs.16121/dados?formato=json&dataInicial=01/{mes:02d}/{ano}&dataFinal={ultimo_dia:02d}/{mes:02d}/{ano}"

    for tentativa in range(tentativas):
        try:
            resposta = requests.get(url, timeout=timeout)
            resposta.raise_for_status()
            dados = resposta.json()
            if len(dados) > 0:
                valor_ipca = dados[0]['valor']
                return float(valor_ipca) / 100
            else:
               # st.write(f"IPCA para {mes}/{ano} não encontrado.")
                return None
        except requests.exceptions.RequestException as err:
            #st.write(f"Tentativa {tentativa + 1} falhou: {err}")
            if tentativa < tentativas - 1:
                time.sleep(5)
            else:
              #st.write(f"Erro ao consultar o Banco Central após {tentativas} tentativas.")
                return None
    return None

# Função para limpar colunas na planilha
def limpar_colunas():
    caminho_arquivo = 'IPCA-Teste.xlsx'
    if not os.path.exists(caminho_arquivo):
        st.write("Erro: Arquivo IPCA-Teste.xlsx não encontrado.")
        return

    planilha = pd.read_excel(caminho_arquivo, engine='openpyxl')
    planilha['TAXA DIA'] = ''
    planilha['taxa 100'] = ''
    planilha['TotalPorcentagem'] = ''
    planilha.to_excel(caminho_arquivo, index=False)
   # st.write("Colunas 'TAXA DIA', 'taxa 100' e 'TotalPorcentagem' foram limpas com sucesso.")

# Função para preencher planilha com IPCA
def preencher_planilha_ipca(ipca_mensal, mes, ano):
    caminho_arquivo = 'IPCA-Teste.xlsx'
    if not os.path.exists(caminho_arquivo):
        st.write("Erro: Arquivo IPCA-Teste.xlsx não encontrado.")
        return

    planilha = pd.read_excel(caminho_arquivo, dtype=str, engine='openpyxl')
    dias_no_mes = calendar.monthrange(ano, mes)[1]
    ipca_diario = ipca_mensal / dias_no_mes
    taxa_100 = 1 + ipca_diario

    ipca_diario_formatado = f"{ipca_diario:.3%}".replace(".", ",")
    taxa_100_formatada = f"{taxa_100:.3%}".replace(".", ",")

    for i, data in planilha['dia'].items():
        dia, mes_data, ano_data = data.split('/')
        
        if int(mes_data) == mes and int(ano_data) == ano:
            planilha.at[i, 'TAXA DIA'] = str(ipca_diario_formatado)
            planilha.at[i, 'taxa 100'] = str(taxa_100_formatada)

    planilha.to_excel(caminho_arquivo, index=False)
   # st.write(f"Planilha preenchida com sucesso para {mes}/{ano}.")

# Função para preencher todas as taxas
def preencher_intervalo_ipca(mes_fim, ano_fim):
    ipca_anterior = None

    for ano in range(2016, ano_fim + 1):
        for mes in range(1, 13):
            if ano == 2016 and mes < 7:
                continue
            if ano == ano_fim and mes > mes_fim:
                break

            ipca_mensal = buscar_ipca(mes, ano)
            if ipca_mensal is None:
                if ipca_anterior is not None:
                   # st.write(f"IPCA para {mes}/{ano} não encontrado. Usando IPCA do mês anterior: {ipca_anterior:.4f}")
                    ipca_mensal = ipca_anterior
                else:
                    st.write(f"Não foi possível encontrar o IPCA para {mes}/{ano} e não há mês anterior para usar.")
                    continue

            ipca_anterior = ipca_mensal
            preencher_planilha_ipca(ipca_mensal, mes, ano)

# Função para calcular o TotalPorcentagem acumulado
def calcular_total_porcentagem():
    caminho_arquivo = 'IPCA-Teste.xlsx'
    if not os.path.exists(caminho_arquivo):
        st.write("Erro: Arquivo IPCA-Teste.xlsx não encontrado.")
        return

    planilha = pd.read_excel(caminho_arquivo, dtype=str)
    valor_acumulado = 1.0

    for i in reversed(range(len(planilha))):
        if pd.notnull(planilha.at[i, 'taxa 100']):
            taxa_100 = float(planilha.at[i, 'taxa 100'].replace(',', '.').replace('%', '')) / 100
            valor_acumulado *= taxa_100
            total_porcentagem = valor_acumulado * 100
            planilha.at[i, 'TotalPorcentagem'] = f"{total_porcentagem:.2f}%".replace('.', ',')

    planilha.to_excel(caminho_arquivo, index=False)
    #st.write("Coluna 'TotalPorcentagem' preenchida com sucesso.")

# Interface do Streamlit
st.sidebar.title("Navegação")
pagina_selecionada = st.sidebar.radio("Selecione a Página", ["Calculadora", "Atualizar Planilha","Selecionar Planilha"])

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
            # Cálculo do valor atualizado baseado na planilha e na data fornecida
            dados_filtro = dados_ipca[dados_ipca['dia'] == pd.to_datetime(data_input)]
            if not dados_filtro.empty:
                # Converte o valor de 'TotalPorcentagem' diretamente para decimal sem somar 1
                taxa = float(dados_filtro['TotalPorcentagem'].values[0].replace('%', '').replace(',', '.')) / 100
                # Multiplica o valor inicial diretamente pela taxa ajustada
                valor_atualizado = valor_input * taxa
                st.write(f"Taxa de juros para dia {data_input}: {taxa * 100:.2f}%")
                st.write(f"Valor Atualizado: R$ {valor_atualizado:.2f}")
            else:
                st.write("Data não encontrada na planilha.")

if pagina_selecionada == "Atualizar Planilha":
    st.title("Atualizar Planilha IPCA")
    with st.form("atualizar_planilha_form"):
        mes = st.number_input("Digite o mês (número, ex: 09 para Setembro):", min_value=1, max_value=12, format="%02d")
        ano = st.number_input("Digite o ano:", min_value=2024, max_value=datetime.today().year + 10)
        confirmar_atualizacao = st.form_submit_button("Confirmar Data Atualizada")

        if confirmar_atualizacao:
            hoje = datetime.today()
            if ano > hoje.year or (ano == hoje.year and mes >= hoje.month):
                st.error("Erro: Só é possível selecionar um mês até o mês passado.")
            else:
                with st.spinner("Atualizando planilha, favor não clicar em nenhum botão nessa janela..."):
                    preencher_coluna_dia(mes, ano)
                    limpar_colunas()
                    preencher_intervalo_ipca(mes, ano)
                    calcular_total_porcentagem()

                    dados_ipca = carregar_dados_excel('IPCA-Teste.xlsx')
                    ultimo_mes = obter_ultimo_mes(dados_ipca)
                    st.write(f"Último mês preenchido na planilha: {ultimo_mes}")
                    st.success("Planilha atualizada com sucesso!")
if pagina_selecionada == "Selecionar Planilha":
    st.title("Selecionar Planilha e Calcular IPCA")
    arquivo = st.file_uploader("Faça upload de sua planilha", type=["xlsx"])

    if arquivo:
        df = pd.read_excel(arquivo, sheet_name=None, engine='openpyxl')
        if "valor" in df and "data" in df:
            planilha_valor = df["valor"]
            planilha_data = df["data"]
            
            if "valor" in planilha_valor.columns and "data" in planilha_data.columns:
                planilha_valor["data"] = pd.to_datetime(planilha_data["data"], errors='coerce')
                planilha_valor["IPCA Calculado"] = planilha_valor.apply(
                    lambda row: row["valor"] * buscar_ipca(row["data"].month, row["data"].year),
                    axis=1
                )
                
                st.write("Planilha processada com sucesso. Visualize abaixo:")
                st.dataframe(planilha_valor)

                # Disponibilizar para download
                caminho_saida = "planilha_atualizada.xlsx"
                planilha_valor.to_excel(caminho_saida, index=False)
                with open(caminho_saida, "rb") as f:
                    st.download_button(
                        label="Baixar Planilha Atualizada",
                        data=f,
                        file_name="planilha_atualizada.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            else:
                st.error("As abas 'valor' ou 'data' não contêm as colunas esperadas.")
        else:
            st.error("As abas 'valor' e 'data' não foram encontradas na planilha.")
