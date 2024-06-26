#AQUI ESTA O GRAFICO MAIS ATUALIZADO 25/06/24


import streamlit as st
import pandas as pd
import plotly.express as px
import io

# Configurações da página do Streamlit
st.set_page_config(page_title='Filtrar Atividades por Data', layout='wide')

# Caminho para a imagem do logo
logo_path = 'logo.png'
st.image(logo_path, width=250)  # Exibir a imagem do logo

# Título e subtítulo abaixo da imagem
st.title('PROJETOS - UT 020')
st.subheader('Gráfico Backlog')
st.write('Obs: Este gráfico mostra os Setores.')

# Carregar os dados principais
df = pd.read_excel("UT-020-IMPRESAO.xlsm")

# Contar as quantidades de OS por setor e status
df_counts = df.groupby(['Setor', 'Status']).size().reset_index(name='Quantidade')

# Gerar as opções do dropdown para seleção de setor
opcoes = list(df['Setor'].unique())
opcoes.append("Todos os Setores")

# Dropdown para selecionar o setor
selected_setor = st.selectbox('Selecione o Setor:', opcoes, index=opcoes.index("Todos os Setores"))

# Filtrar os dados com base na seleção do setor
if selected_setor == "Todos os Setores":
    df_filtered = df_counts
else:
    df_filtered = df_counts[df_counts['Setor'] == selected_setor]

# Criar o gráfico de barras agrupadas
fig = px.bar(df_filtered, x="Setor", y="Quantidade", color="Status", barmode="group")
st.plotly_chart(fig, use_container_width=True)



# Carregar os dados secundários
file_path = 'TODAS_OS_20.xlsx'
df_programacao = pd.read_excel(file_path)

# Certificar-se de que a coluna 'Data_Solicitacao' está no formato datetime
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

    # Entrada para filtrar por número de OS Cliente
    os_cliente_filtro = st.text_input("Filtrar OS Cliente",
                                      key='os_cliente_filtro')

    # Aplicar o filtro se um número de OS Cliente for especificado
    if os_cliente_filtro:
        filtered_df = filtered_df[filtered_df['OS Cliente'] == os_cliente_filtro]

    # Mostrar a contagem de resultados filtrados
    st.write(f"Atividades encerradas de {start_date.date()} até {end_date.date()} para o OS Cliente '{os_cliente_filtro}':")
    st.write(f"Total de registros encontrados: {len(filtered_df)}")

    # Mostrar o DataFrame filtrado
    st.dataframe(filtered_df)

    # Adicionar uma tabela resumida
    st.subheader("Resumo das Solicitações por Setor e Status")
    summary_df = filtered_df.groupby(['Setor', 'Denominação Estado OS']).size().unstack(fill_value=0)
    st.dataframe(summary_df)

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

# Função para mostrar análise do backlog
def show_backlog_analysis():
    # Carregar os dados principais para o Backlog Geral
    df = pd.read_excel("TODAS_OS_20.xlsx")

    # Ajustar o nome das colunas de datas de acordo com o seu arquivo Excel ('Data/Hora Edição')
    df['Data/Hora Edição'] = pd.to_datetime(df['Data/Hora Edição'], errors='coerce')

    # Drop linhas com datas inválidas (NaT)
    df.dropna(subset=['Data/Hora Edição'], inplace=True)

    # Filtrar as OS que não estão encerradas nem canceladas
    df = df[~df['Denominação Estado OS'].isin(['ENCERRADA', 'CANCELADA'])]

    # Calcular o valor total nominal do backlog com base em todos os dados
    total_nominal = df.shape[0]  # Total de linhas na planilha

    # Exibir o valor total nominal do backlog geral
    st.write(f"Valor total do backlog: {total_nominal}")

    # Filtrar por intervalo de datas para o Backlog Geral
    start_date_backlog_geral = st.date_input("Data de início - Backlog Geral", pd.to_datetime("2024-01-01"), key="start_backlog_geral")
    end_date_backlog_geral = st.date_input("Data de término - Backlog Geral", pd.to_datetime("2024-07-31"), key="end_backlog_geral")

    # Converter para datetime64[ns]
    start_date_backlog_geral = pd.to_datetime(start_date_backlog_geral)
    end_date_backlog_geral = pd.to_datetime(end_date_backlog_geral)

    # Filtrar os dados com base no intervalo de datas selecionado para o Backlog Geral
    df_filtered_backlog_geral = df[(df['Data/Hora Edição'] >= start_date_backlog_geral) &
                                   (df['Data/Hora Edição'] <= end_date_backlog_geral)]

    # Criar um DataFrame para contar as quantidades de OS por mês para o Backlog Geral
    months_backlog_geral = pd.date_range(start=start_date_backlog_geral, end=end_date_backlog_geral, freq='M').to_period('M')
    backlog_geral = pd.DataFrame({'Mês': months_backlog_geral.to_timestamp()})

    # Calcular o backlog acumulado
    initial_backlog = df[(df['Data/Hora Edição'] < start_date_backlog_geral)].shape[0]
    backlog_geral['Entradas'] = backlog_geral['Mês'].apply(lambda x: ((df_filtered_backlog_geral['Data/Hora Edição'] >= x) &
                                                                     (df_filtered_backlog_geral['Data/Hora Edição'] < x + pd.DateOffset(months=1))).sum())
    backlog_geral['Backlog Geral'] = backlog_geral['Entradas'].cumsum() + initial_backlog
    backlog_geral['Mês'] = backlog_geral['Mês'].dt.strftime('%b %Y')

    # Criar o gráfico de linhas para o Backlog Geral
    fig_backlog_geral = px.line(backlog_geral, x="Mês", y="Backlog Geral", markers=True)
    fig_backlog_geral.update_layout(
        title="Evolução do Backlog Geral (2024)",
        xaxis_title='Mês',
        yaxis_title='Quantidade de OS no Backlog Geral'
    )

    # Exibir o gráfico do Backlog Geral
    st.plotly_chart(fig_backlog_geral, use_container_width=True)

    # Filtrar os dados com base no intervalo de datas selecionado para o Backlog por Setor
    df_filtered_backlog_setor = df[(df['Data/Hora Edição'] >= start_date_backlog_geral) &
                                   (df['Data/Hora Edição'] <= end_date_backlog_geral)]

    # Criar um DataFrame para contar as quantidades de OS por mês e por setor para o Backlog por Setor
    backlog_setor = df_filtered_backlog_setor.groupby(['Setor', pd.Grouper(key='Data/Hora Edição', freq='M')]).size().reset_index(name='Entradas')
    backlog_setor['Backlog Setor'] = backlog_setor.groupby('Setor')['Entradas'].cumsum()
    backlog_setor['Mês'] = backlog_setor['Data/Hora Edição'].dt.strftime('%b %Y')

    # Criar o gráfico de linhas por setor para o Backlog por Setor
    fig_backlog_setor = px.line(backlog_setor, x="Mês", y="Backlog Setor", color='Setor', markers=True)
    fig_backlog_setor.update_layout(
        title="Evolução do Backlog por Setor ao Longo do Tempo (2024)",
        xaxis_title='Mês',
        yaxis_title='Quantidade de OS no Backlog por Setor'
    )

    # Exibir o gráfico do Backlog por Setor
    st.plotly_chart(fig_backlog_setor, use_container_width=True)


# Chamar a função para exibir a análise de evolução do backlog
show_backlog_analysis()

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
