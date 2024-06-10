import streamlit as st
import pandas as pd
import plotly.express as px
import io

# Configurações da página do Streamlit
st.set_page_config(page_title='Filtrar Atividades por Data', layout='wide')

# Carregar os dados
df = pd.read_excel("UT-020-IMPRESAO.xlsx")

# Contar as quantidades de OS por setor e status
df_counts = df.groupby(['Setor', 'Status']).size().reset_index(name='Quantidade')

# Gerar as opções do dropdown
opcoes = list(df['Setor'].unique())
opcoes.append("Todos os Setores")

# Título e subtítulo
st.title('MANSERV - PROJETOS - UT 020')
st.subheader('Gráfico Ordem de Serviços')
st.write('Obs: Esse gráfico mostra os Setores.')

# Criar a figura
fig = px.bar(df_counts, x="Setor", y="Quantidade", color="Status", barmode="group")

# Exibir o gráfico ocupando toda a largura da tela
st.plotly_chart(fig, use_container_width=True)

# Carregar o arquivo Excel diretamente do diretório
file_path = 'TODAS_OS_20.xlsx'
df = pd.read_excel(file_path)

# Certifique-se de que a coluna 'Data_Solicitacao' está no formato datetime
df['Data_Solicitacao'] = pd.to_datetime(df['Data_Solicitacao'])

# Entrada para o intervalo de datas desejado
start_date = st.date_input("Escolha a data de início para filtrar as atividades", value=pd.to_datetime("2024-01-01"))
end_date = st.date_input("Escolha a data de término para filtrar as atividades", value=pd.to_datetime("2024-12-31"))

if start_date and end_date:
    # Filtrar os dados para o intervalo de datas selecionado
    filtered_df = df[(df['Data_Solicitacao'] >= pd.to_datetime(start_date)) & (df['Data_Solicitacao'] <= pd.to_datetime(end_date))]

    # Filtro para seleção de setores
    selected_sectors = st.multiselect("Selecione os setores", options=filtered_df['Setor'].unique(), default=filtered_df['Setor'].unique())

    # Filtro para seleção de status
    selected_status = st.multiselect("Selecione o status das solicitações", options=filtered_df['Denominação Estado OS'].unique(), default=filtered_df['Denominação Estado OS'].unique())

    # Aplicar os filtros de seleção de setor e status
    filtered_df = filtered_df[filtered_df['Setor'].isin(selected_sectors) & filtered_df['Denominação Estado OS'].isin(selected_status)]

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
                  title=f'Quantidade de Solicitações por Setor ({start_date} a {end_date})',
                  labels={'Setor': 'Setor', 'value': 'Quantidade'},
                  barmode='group',
                  text=merged_counts[['Porcentagem', 'Encerradas']].apply(lambda x: f'{x[0]:.2f}%', axis=1))

    # Atualizar o layout do gráfico para mostrar os rótulos de texto
    fig1.update_traces(texttemplate='%{text}', textposition='outside')

    # Mostrar o primeiro gráfico no Streamlit
    st.plotly_chart(fig1)

    # Criar o segundo gráfico (histograma)
    fig2 = px.histogram(filtered_df, x="Data_Solicitacao",
                        histfunc="count", title="Historico de Solicitações ao Longo do Tempo")
    fig2.update_traces(xbins_size="M1")
    fig2.update_xaxes(showgrid=True, ticklabelmode="period", dtick="M1", tickformat="%b\n%Y")
    fig2.update_layout(bargap=0.1)

    # Mostrar o segundo gráfico no Streamlit
    st.plotly_chart(fig2)

    # Adicionar uma tabela resumida
    st.subheader("Resumo das Solicitações por Setor e Status")
    summary_df = filtered_df.groupby(['Setor', 'Denominação Estado OS']).size().unstack(fill_value=0)
    st.dataframe(summary_df)

    # Exibir o DataFrame filtrado
    st.write(f"Atividades encerradas de {start_date} até {end_date}:")
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


# Para rodar o servidor do Streamlit:
# cd /Users/regis/PycharmProjects/grafico.os/
# streamlit run os.servico.py
