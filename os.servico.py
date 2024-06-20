#AQUI ESTA O GRAFICO MAIS ATUALIZADO 19/06/24


import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import io
import locale

# Configurações da página do Streamlit
st.set_page_config(page_title='Filtrar Atividades por Data', layout='wide')

# Definir idioma como português
locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')

# Caminho para a imagem do logo
logo_path = 'logo.png'  # Certifique-se de que o arquivo logo.png está no mesmo diretório do seu script

# Exibir a imagem do logo
st.image(logo_path, width=250)  # Ajuste o tamanho conforme necessário

# Título e subtítulo abaixo da imagem
st.title('PROJETOS - UT 020')
st.subheader('Gráfico Backlog')
st.write('Obs: Esse gráfico mostra os Setores.')

# Carregar os dados
df = pd.read_excel("UT-020-IMPRESAO.xlsx")

# Contar as quantidades de OS por setor e status
df_counts = df.groupby(['Setor', 'Status']).size().reset_index(name='Quantidade')

# Gerar as opções do dropdown
opcoes = list(df['Setor'].unique())
opcoes.append("Todos os Setores")

# Dropdown para selecionar o setor
selected_setor = st.selectbox('Selecione o Setor:', opcoes, index=opcoes.index("Todos os Setores"))

# Filtrar os dados com base na seleção
if selected_setor == "Todos os Setores":
    df_filtered = df_counts
else:
    df_filtered = df_counts[df_counts['Setor'] == selected_setor]

# Criar a figura
fig = px.bar(df_filtered, x="Setor", y="Quantidade", color="Status", barmode="group")

# Exibir o gráfico ocupando toda a largura da tela
st.plotly_chart(fig, use_container_width=True)

# Carregar o arquivo Excel diretamente do diretório
file_path = 'TODAS_OS_20.xlsx'
df_programacao = pd.read_excel(file_path)

# Certifique-se de que a coluna 'Data_Solicitacao' está no formato datetime
df_programacao['Data_Solicitacao'] = pd.to_datetime(df_programacao['Data_Solicitacao'], errors='coerce')

# Entrada para o intervalo de datas desejado
start_date = st.date_input("Escolha a data de início para filtrar as atividades",
                           value=pd.to_datetime("2024-01-01").date())
end_date = st.date_input("Escolha a data de término para filtrar as atividades",
                         value=pd.to_datetime("2024-12-31").date())

if start_date and end_date:
    # Converta start_date e end_date para datetime64[ns]
    start_date = pd.to_datetime(start_date)
    end_date = pd.to_datetime(end_date)

    # Filtrar os dados para o intervalo de datas selecionado
    filtered_df = df_programacao[
        (df_programacao['Data_Solicitacao'] >= start_date) & (df_programacao['Data_Solicitacao'] <= end_date)]

    # Filtro para seleção de setores
    selected_sectors = st.multiselect("Selecione os setores", options=filtered_df['Setor'].unique(),
                                      default=filtered_df['Setor'].unique())

    # Filtro para seleção de status
    selected_status = st.multiselect("Selecione o status das solicitações",
                                     options=filtered_df['Denominação Estado OS'].unique(),
                                     default=filtered_df['Denominação Estado OS'].unique())

    # Aplicar os filtros de seleção de setor e status
    filtered_df = filtered_df[
        filtered_df['Setor'].isin(selected_sectors) & filtered_df['Denominação Estado OS'].isin(selected_status)]

    # Contar a quantidade de Data_Solicitacao para cada setor
    sector_counts = filtered_df['Setor'].value_counts().reset_index()
    sector_counts.columns = ['Setor', 'Quantidade']

    # Calcular a porcentagem de solicitações para cada setor
    total_solicitations = sector_counts['Quantidade'].sum()
    sector_counts['Porcentagem'] = (sector_counts['Quantidade'] / total_solicitations) * 100

    # Filtrar apenas as solicitações encerradas
    closed_df = filtered_df[filtered_df['Denominação Estado OS'] == 'ENCERRADA']

    # Contar a quantidade de solicitações encerradas para cada setor
    closed_counts = closed_df['Setor'].value_counts().reset_index()
    closed_counts.columns = ['Setor', 'Encerradas']

    # Mesclar os dois DataFrames
    merged_counts = pd.merge(sector_counts, closed_counts, on='Setor', how='left')
    merged_counts['Encerradas'] = merged_counts['Encerradas'].fillna(0)  # Substituir NaNs por 0

    # Criar o gráfico de barras agrupadas com Plotly
    fig1 = px.bar(merged_counts, x='Setor', y=['Quantidade', 'Encerradas'],
                  title=f'Quantidade de Solicitações por Setor ({start_date.date()} a {end_date.date()})',
                  labels={'Setor': 'Setor', 'value': 'Quantidade'},
                  barmode='group',
                  text=merged_counts[['Porcentagem', 'Encerradas']].apply(lambda x: f'{x[0]:.2f}%', axis=1))

    # Atualizar o layout do gráfico para mostrar os rótulos de texto e ocupar toda a largura da tela
    fig1.update_traces(texttemplate='%{text}', textposition='outside')
    fig1.update_layout(width=1200)  # Ajuste o valor da largura conforme necessário

    # Mostrar o primeiro gráfico no Streamlit
    st.plotly_chart(fig1, use_container_width=True)

    # Criar o segundo gráfico (histograma)
    total_counts = filtered_df['Data_Solicitacao'].dt.to_period('M').value_counts().sort_index()
    closed_counts = closed_df['Data_Solicitacao'].dt.to_period('M').value_counts().sort_index()

    # Calcular a porcentagem de solicitações encerradas
    closed_percentage = (closed_counts / total_counts) * 100

    fig2 = go.Figure()

    # Adicionar barras totais sem especificar cor (usará as cores padrão do Plotly Express)
    fig2.add_bar(x=total_counts.index.to_timestamp(), y=total_counts.values,
                 name='Total Solicitações')

    # Adicionar barras de solicitações encerradas sem especificar cor
    fig2.add_bar(x=closed_counts.index.to_timestamp(), y=closed_counts.values,
                 name='Total Encerradas')

    # Adicionar porcentagem sobre as barras de solicitações encerradas
    for i, value in enumerate(closed_counts.values):
        x = closed_counts.index[i].to_timestamp()
        y = value
        percentage_text = f'{closed_percentage[i]:.2f}%'

        # Adicionar a anotação sobre a barra
        fig2.add_annotation(x=x, y=y, text=percentage_text,
                            showarrow=False, font=dict(size=14, family="Arial"),
                            xshift=0, yshift=10, align='center', valign='bottom')

    # Atualizar o layout do gráfico
    fig2.update_layout(
        title="Histórico de Solicitações ao Longo do Tempo",
        barmode='group',
        xaxis_title='Data',
        yaxis_title='Quantidade',
        bargap=0.1,
        xaxis=dict(
            ticklabelmode="period",
            dtick="M1",
            tickformat="%b\n%Y"
        )
    )

    # Mostrar o segundo gráfico no Streamlit
    st.plotly_chart(fig2, use_container_width=True)

    # Adicionar uma tabela resumida
    st.subheader("Resumo das Solicitações por Setor e Status")
    summary_df = filtered_df.groupby(['Setor', 'Denominação Estado OS']).size().unstack(fill_value=0)
    st.dataframe(summary_df)

    # Exibir o DataFrame filtrado
    st.write(f"Atividades encerradas de {start_date.date()} até {end_date.date()}:")
    st.dataframe(filtered_df)

    # Função para converter DataFrame em Excel
    def to_excel(df):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='Sheet1')
        processed_data = output.getvalue()
        return processed_data

    # Gerar dados do Excel
    excel_data = to_excel(filtered_df)

    # Botão para download do Excel
    st.download_button(
        label="Baixar tabela em Excel",
        data=excel_data,
        file_name='atividades_filtradas.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

# Código para filtrar e exibir tabela de outra planilha

# Carregar os dados ignorando a primeira coluna
df_programacao_diaria = pd.read_excel("PROGRAMACAO-DIARIA.xlsx", usecols=lambda col: col not in [0])

# Entrada para a data desejada
date_to_filter = st.date_input("Escolha a data para filtrar as atividades", key='programacao')

if date_to_filter:
    # Converter a coluna de datas para datetime se necessário
    if not pd.api.types.is_datetime64_any_dtype(df_programacao_diaria['Data Prevista']):
        df_programacao_diaria['Data Prevista'] = pd.to_datetime(df_programacao_diaria['Data Prevista'], errors='coerce')

    # Filtrar o DataFrame pela data escolhida
    filtered_df_programacao_diaria = df_programacao_diaria[
        df_programacao_diaria['Data Prevista'].dt.date == date_to_filter]


    # Exibir o DataFrame filtrado
    st.write(f"Atividades para {date_to_filter}:")
    st.dataframe(filtered_df_programacao_diaria)

    # Gerar dados do Excel
    excel_data_diaria = to_excel(filtered_df_programacao_diaria)

    # Botão para download do Excel
    st.download_button(
        label="Baixar tabela em Excel - Programação Diária",
        data=excel_data_diaria,
        file_name='programacao_diaria_filtrada.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

# Para rodar o servidor do Streamlit:
# cd /Users/regis/PycharmProjects/grafico.os/
# streamlit run os.servico.py
