import streamlit as st
import plotly.express as px
import pandas as pd

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

# Para rodar o servidor do Streamlit:
# cd /Users/regis/PycharmProjects/grafico.os/
# streamlit run os.servico.py
