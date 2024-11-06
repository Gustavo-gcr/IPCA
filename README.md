
## Sumário
Este projeto consiste em uma aplicação de calculadora de juros IPCA com interface em Streamlit e atualização de planilha em Excel com taxas diárias e totais de porcentagem. A aplicação realiza diversas operações com a planilha, incluindo preenchimento de datas, busca de valores de IPCA no Banco Central, e cálculo de valores ajustados com base no IPCA.

---

### Estrutura de Pastas
- `atualizar.py`: Script Python para atualizar a planilha Excel com dados IPCA.
- `IPCA-Teste.xlsx`: Planilha Excel com as colunas `dia`, `TAXA DIA`, `taxa 100`, e `TotalPorcentagem`.
- `requirements.txt`: Arquivo para instalar dependências, e para hospedar a automação.

---

### Funções Principais

#### Funções de Interface (Streamlit)
1. **carregar_dados_excel**: Carrega dados do Excel a partir da pasta do script.
2. **calcular_valor_ajustado**: Calcula valor ajustado baseado em uma taxa percentual.
3. **Interface Streamlit**: Permite que o usuário insira valores e selecione uma data, retornando o valor ajustado pela taxa IPCA do dia.

#### Funções de Manipulação da Planilha
1. **preencher_coluna_dia**: Preenche a coluna 'dia' com datas de 01/06/2016 até 10 anos após a data final.
2. **buscar_ipca**: Consulta o IPCA no Banco Central e retorna o valor mensal.
3. **limpar_colunas**: Limpa as colunas `TAXA DIA`, `taxa 100` e `TotalPorcentagem`.
4. **preencher_planilha_ipca**: Preenche a planilha com a taxa diária e a taxa 100.
5. **preencher_intervalo_ipca**: Preenche todas as taxas de 07/2016 até o mês/ano final especificado.
6. **calcular_total_porcentagem**: Calcula o total acumulado de porcentagem para a coluna `TotalPorcentagem`.

#### Script Batch
O arquivo batch (`atualizar.bat`) solicita ao usuário mês e ano, e executa o `atualizar.py` com os parâmetros fornecidos.

---

## Autor:

- [Gustavo Coelho](https://github.com/Gustavo-gcr)
