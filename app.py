import dash
from dash import dcc, html, Input, Output
import pandas as pd
import plotly.graph_objs as go

# Definir o caminho do arquivo Excel
file_path = 'arquivo.xlsx'

# Criar a aplicação Dash
app = dash.Dash(__name__)

try:
    # Carregar o arquivo Excel
    df = pd.read_excel(file_path)
    print("Arquivo Excel carregado com sucesso.")
    print("Colunas disponíveis no DataFrame:", df.columns.tolist())
except FileNotFoundError:
    print(f"Arquivo não encontrado: {file_path}")
    df = pd.DataFrame()  # DataFrame vazio para evitar erros
except Exception as e:
    print(f"Erro ao carregar o arquivo Excel: {e}")
    df = pd.DataFrame()  # DataFrame vazio para evitar erros

# Verificar se as colunas necessárias existem no DataFrame
if 'Unnamed: 1' in df.columns and 'Unnamed: 4' in df.columns:
    # Filtrar linhas que contêm dados válidos (removendo cabeçalhos repetidos)
    df = df[df['Unnamed: 1'].notna() & df['Unnamed: 4'].notna()]

    # Renomear colunas para facilitar o uso
    df.columns = [
        'Mes', 'Data', 'Ministro', 'Dirigente', 'Qtd_Pessoas', 
        'Qtd_Visitantes', 'Qtd_Criancas', 'Qtd_Conversoes', 
        'Qtd_Batismo_Esp_Santo', 'Qtd_Motos', 'Qtd_Carros'
    ]

    # Extrair o ano e o mês da coluna 'Data'
    df['Data'] = pd.to_datetime(df['Data'], errors='coerce')
    df['Ano'] = df['Data'].dt.year
    df['Mes'] = df['Data'].dt.month

    # Adicionar anos disponíveis entre 2022 e 2024, mesmo que não estejam no DataFrame
    anos_disponiveis = list(range(2022, 2025))

else:
    df = pd.DataFrame()
    anos_disponiveis = [2022, 2023, 2024]
    print("As colunas necessárias não foram encontradas no arquivo Excel.")

# Criar uma lista com todos os meses
meses = list(range(1, 13))
nomes_meses = [
    'Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 
    'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'
]

# Layout do dashboard
app.layout = html.Div([
    html.H1("Dashboard de Frequência dos Cultos", style={'textAlign': 'center'}),
    
    # Dropdown para selecionar o ano
    html.Div([
        html.Label("Selecione o Ano:"),
        dcc.Dropdown(
            id='ano-dropdown',
            options=[{'label': str(ano), 'value': ano} for ano in anos_disponiveis],
            value=anos_disponiveis[0] if anos_disponiveis else None,
            clearable=False,
            style={'width': '50%', 'margin': '0 auto'}
        )
    ], style={'textAlign': 'center', 'padding': '20px'}),
    
    # Gráficos
    dcc.Graph(id='grafico-frequencia-culto'),
    dcc.Graph(id='grafico-media-pessoas'),
    dcc.Graph(id='grafico-media-visitantes'),
    html.Div(id='mensagem-erro', style={'color': 'red', 'textAlign': 'center', 'marginTop': '20px'})
])

# Callback para atualizar os gráficos com base no ano selecionado
@app.callback(
    [Output('grafico-frequencia-culto', 'figure'),
     Output('grafico-media-pessoas', 'figure'),
     Output('grafico-media-visitantes', 'figure'),
     Output('mensagem-erro', 'children')],
    [Input('ano-dropdown', 'value')]
)
def atualizar_graficos(ano_selecionado):
    if ano_selecionado is None or df.empty:
        return {}, {}, {}, "Nenhum dado disponível para exibir os gráficos."

    # Filtrar o DataFrame pelo ano selecionado
    df_filtrado = df[df['Ano'] == ano_selecionado]

    # Preencher os meses ausentes com zero para garantir que todos os meses sejam exibidos
    df_mes = pd.DataFrame({'Mes': meses})
    df_filtrado = df_mes.merge(df_filtrado, on='Mes', how='left')

    # Preencher os campos numéricos com zero quando houver dados ausentes
    df_filtrado['Qtd_Pessoas'].fillna(0, inplace=True)
    df_filtrado['Qtd_Visitantes'].fillna(0, inplace=True)

    # Gráfico de frequência de culto por mês
    trace_frequencia = go.Scatter(
        x=nomes_meses, y=df_filtrado['Qtd_Pessoas'], mode='lines+markers+text',
        name='Qtd. de Pessoas', line=dict(color='orange'), text=df_filtrado['Qtd_Pessoas'],
        textposition='top center'
    )

    layout_frequencia = go.Layout(
        title=f'Frequência dos Cultos - {ano_selecionado} (com números exatos)',
        xaxis=dict(title='Meses'),
        yaxis=dict(title='Quantidade'),
        template='plotly_white'
    )

    fig_frequencia = go.Figure(data=[trace_frequencia], layout=layout_frequencia)

    # Gráfico de média de pessoas por culto
    media_pessoas = df_filtrado['Qtd_Pessoas'].mean()
    fig_media_pessoas = go.Figure(
        data=[go.Bar(x=['Média de Pessoas'], y=[media_pessoas], text=[f"{media_pessoas:.2f}"],
                     textposition='outside', marker=dict(color='#636EFA'))],
        layout=go.Layout(title=f'Média de Pessoas por Culto - {ano_selecionado}', template='plotly_white')
    )

    # Gráfico de média de visitantes por culto
    media_visitantes = df_filtrado['Qtd_Visitantes'].mean()
    fig_media_visitantes = go.Figure(
        data=[go.Bar(x=['Média de Visitantes'], y=[media_visitantes], text=[f"{media_visitantes:.2f}"],
                     textposition='outside', marker=dict(color='#EF553B'))],
        layout=go.Layout(title=f'Média de Visitantes por Culto - {ano_selecionado}', template='plotly_white')
    )

    return fig_frequencia, fig_media_pessoas, fig_media_visitantes, ""

if __name__ == '__main__':
    app.run_server(debug=True)
