import streamlit as st
import plotly.express as px
import pandas as pd
from io import BytesIO

# Configurações da página do Streamlit
st.set_page_config(page_title='Filtrar Atividades por Data', layout='wide')

# Carregar os dados
df = pd.read_excel("UT-020-PROJETOS.xlsx")

# Contar as quantidades de OS por setor e status
df_counts = df.groupby(['Setor', 'Status']).size().reset_index(name='Quantidade')

# Gerar as opções do dropdown
opcoes = list(df['Setor'].unique())
opcoes.append("Todos os Setores")

# Título e subtítulo
st.title('MANSERV - PROJETOS - UT 020')
st.subheader('Gráfico Ordem de Serviços')
st.write('Obs: Esse gráfico mostra os Setores.')

# Dropdown para selecionar o setor
selected_setor = st.selectbox('Selecione o Setor:', opcoes, index=opcoes.index("Todos os Setores"))

# Filtrar os dados com base na seleção
if selected_setor == "Todos os Setores":
    df_filtered = df_counts
else:
    df_filtered = df_counts[df_counts['Setor'] == selected_setor]

# Criar a figura
fig = px.bar(df_filtered, x="Setor", y="Quantidade", color="Status", barmode="group")

# Exibir o gráfico
st.plotly_chart(fig)

# Código para filtrar e exibir tabela de outra planilha

# Carregar os dados ignorando a primeira coluna
df_programacao = pd.read_excel("PROGRAMACAO-DIARIA.xlsx", usecols=lambda col: col not in [0])

# Entrada para a data desejada
date_to_filter = st.date_input("Escolha a data para filtrar as atividades")

if date_to_filter:
    # Converter a coluna de datas para datetime se necessário
    if not pd.api.types.is_datetime64_any_dtype(df_programacao['Data Prevista']):
        df_programacao['Data Prevista'] = pd.to_datetime(df_programacao['Data Prevista'], errors='coerce')

    # Filtrar o DataFrame pela data escolhida
    filtered_df_programacao = df_programacao[df_programacao['Data Prevista'].dt.date == date_to_filter]

    # Exibir o DataFrame filtrado
    st.write(f"Atividades para {date_to_filter}:")
    st.dataframe(filtered_df_programacao)

    # Função para converter DataFrame para Excel
    def to_excel(df):
        output = BytesIO()
        writer = pd.ExcelWriter(output, engine='xlsxwriter')
        df.to_excel(writer, index=False, sheet_name='Atividades')
        writer.close()
        processed_data = output.getvalue()
        return processed_data

    # Botão para download do Excel
    st.download_button(
        label="Baixar tabela em Excel",
        data=to_excel(filtered_df_programacao),
        file_name=f"Atividades_{date_to_filter}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# Para rodar o servidor do Streamlit:
# cd /Users/regis/PycharmProjects/grafico.os/
# streamlit run os.servico.py
