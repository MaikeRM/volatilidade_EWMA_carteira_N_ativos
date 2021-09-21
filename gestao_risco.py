######################################################################
# Cálculo de Rentabilidade e Risco da Carteira de Ações - Carlos Lee #
######################################################################

##Importando bibliotecas
import yfinance as yf
import pandas as pd
import numpy as np
import warnings
from scipy.stats import norm
warnings.filterwarnings('ignore')
import matplotlib
import plotly.graph_objects as go
from datetime import date
matplotlib.rcParams['figure.figsize'] = (18,8)

#Fazendo o upload do arquivo trades.xlsx
arquivo = pd.read_excel(r'Excel\Lee S Wen.xlsx', sheet_name='VISTA')
arquivo.columns = arquivo.iloc[0]
arquivo = arquivo.iloc[2:] 
arquivo.index = arquivo["Data"]
arquivo.columns = ["Data", "ticker", "Nome", "C/V", "Quantidade", "Unitário (s corretagem)", 
                    "Total (s corretagem)", "Unitário (c corretagem)", "Total (c corretagem)", "Total Despesas", "NaN"]
arquivo.drop(columns=["Data","Nome", "Unitário (s corretagem)", "Total (s corretagem)", "Total Despesas", "NaN"], inplace=True)

#Tratando dados enviados pela Fabiana

### Transforma sinal de venda para negativo
raws = np.linspace(0, len(arquivo)-1, len(arquivo), dtype=int)
for a in raws:
    if arquivo["C/V"][a] == "V":
        arquivo["Quantidade"][a] = -arquivo["Quantidade"][a]

### Criando tabela com colunas para cada ativo e indexando por data
trade_quant = pd.pivot_table(arquivo, values='Quantidade', index=['Data'], columns=arquivo['ticker'].str.upper(), aggfunc=np.sum, fill_value=0)

### Criando tabela com os preços de compra e venda
trade_price = pd.pivot_table(arquivo, values="Unitário (c corretagem)", index=['Data'], columns=arquivo['ticker'].str.upper(), aggfunc=np.sum,fill_value=0)

### Baixando os cotações das ações
prices = yf.download(tickers=(trade_quant.columns+'.SA').to_list(), start=trade_quant.index[0], rounding=True)['Adj Close']

### Consolida posições
prices.columns  = prices.columns.str.rstrip('.SA')
prices.dropna(how='all', inplace=True)
trades = trade_quant.reindex(index=prices.index)
trades.fillna(value=0, inplace=True)
aportes = (trades * trade_price).sum(axis=1)
posicao = trades.cumsum()

### Consolida saldo
carteira = posicao * prices
carteira['saldo'] = carteira.sum(axis=1)
#carteira = carteira[:-2]

### Exporta para Excel
data_atual = date.today()
carteira.to_excel('carteira_{}-{}-{}.xlsx'.format(data_atual.day, data_atual.month,data_atual.year))

########################
# Cálculo Volatilidade #
########################

#Baixa os dados históricos
hist_price = yf.download(tickers=(trade_quant.columns+'.SA').to_list(), period="3y")['Adj Close']
retr_log = np.log(hist_price).diff()
retr_log.dropna(inplace=True)

#Calcula o peso de cada ação  - dados históricos de movimentações
weights = carteira.div(carteira['saldo'], axis=0)
weights = weights.iloc[:,:-1]
weights = weights.drop(weights.index[243])
weights.columns = retr_log.columns

#Calcula a volatilidade EWMA para todos os ativos em uma janela de 252 dias
r = retr_log.ewm(alpha = 0.05, min_periods=252).std(252)
vol  = (r)**(1/1)
vol['2020':].plot(figsize=(30,10))

#Baixa os dados e calcula volatilidade para o IBOVESPA
prices = yf.download(tickers= "^BVSP", period="3y", rounding=True)['Adj Close']
retr_log_ibov = np.log(prices).diff()
retr_log_ibov.dropna(inplace=True)
ibov_vol = retr_log_ibov.ewm(alpha = 0.05, min_periods=252).std(252)
ibov_vol  = (ibov_vol)**(1/1)*100

#Calculo de correlação entre os ativos
corr_ewma = retr_log.ewm(alpha = 0.05).corr()
new_index = pd.MultiIndex.from_tuples(corr_ewma.index, names=['Date','Ticker'])
corr_ewma = pd.DataFrame(corr_ewma, columns = corr_ewma.columns, index=new_index)

#limpa dados NA
vol.dropna( inplace=True)
vol = vol[weights.index[0]:weights.index[-1]]

#Calcula a volatilidade da carteira do Sr. Carlos utilizando matrizes
a=-1
n = 0
vol_cart = pd.DataFrame(columns=list(weights.columns.values))

#Calcula o corr*pesos par a par ( ex: BBDC4 x BRFS4 .... PETR4 x VALE3)
for n in range(len(weights.index)):

    vol_cart.loc[n] = (vol.iloc[n,]*weights.iloc[n,])*corr_ewma.loc[weights.index[n]].dot(weights.iloc[n,])

    n = n+1

vol_cart = pd.DataFrame(vol_cart.sum(axis=1)*100)
vol_cart.index = weights.index
vol_cart.columns = ['Carteira Lee']

#Adiciona o IBOV para Benchmark
vol_cart['IBOV'] = ibov_vol[weights.index[0]:weights.index[-1]]


####################################################################################################################

#Visualização com o Plotly carteira
# import plotly.express as px
# fig = px.bar(carteira.iloc[-1].loc[~(carteira.iloc[-1]==0)], width=1500, height=800)

# fig.update_xaxes(
#     dtick="M1",
#     tickformat="%b\n%Y",
#     ticklabelmode="period")

# fig.update_layout(
#     margin=dict(l=40, r=40, t=40, b=40),
#     paper_bgcolor="LightSteelBlue",
# )

# fig.show()

####################################################################################################################

# #Visualização com o Plotly da Volatilidade
# fig = px.line(vol_cart, title='Volatilidade', width=1500, height=800)

# fig.update_xaxes(
#     dtick="M1",
#     tickformat="%b\n%Y",
#     ticklabelmode="period")

# fig.update_layout(
#     margin=dict(l=40, r=40, t=40, b=40),
#     paper_bgcolor="LightSteelBlue",
# )

# fig.show()


###############
# Cálculo VaR #
###############

inter_conf = 0.95
#VaR das acoes
VaR = -vol*norm.ppf(inter_conf)

#VaR portifólio
VaR_Cart = -vol_cart*norm.ppf(inter_conf)

#Cálculo do erro 
#n=0
#c = pd.DataFrame(index= VaR.index)
#for acao in VaR:
#    n=0
#    for a in range(len(VaR)):
#       if retr_log[acao]['2020-04-06':'2021-05-20'].iloc[a] < VaR[acao]['2020-04-06':].iloc[a]:
#        n = n+1
#        if a > 275: break
#    c[acao] = n
#            
#erro = (c.sum()/len(VaR))/len(VaR)*100

#Calcula retorno da carteira ponderado pelos pesos
retr_cart = retr_log*weights
retr_cart.dropna(inplace=True)
retr_cart = retr_cart.sum(axis=1)

VaR_Cart['retorno'] = retr_cart*100
VaR_Cart.columns = ['VaR Lee', 'VaR IBOV', 'Retorno Lee']
VaR_Cart['retr_ibov'] = retr_log_ibov*100

#Exportar para excel
VaR_Cart.to_excel('var_cart.xlsx')
vol_cart.to_excel('vol_cart.xlsx')


# #Cria gráfico comparativo var
# fig = px.line(VaR_Cart, width=1500, height=800)

# fig.update_xaxes(
#     dtick="M1",
#     tickformat="%b\n%Y",
#     ticklabelmode="period")

# fig.update_layout(
#     margin=dict(l=40, r=40, t=40, b=40),
#     paper_bgcolor="LightSteelBlue",
# )

# fig.show()

############################################################################

# #Visualização com o Plotly do VaR comparacao com Ibov
# from plotly.subplots import make_subplots

# import pandas as pd
# import re

# df = VaR_Cart
# df['Data'] = VaR_Cart.index
# df = df.reindex(columns=["Data", "VaR Lee", "VaR IBOV", "Retorno Lee", "retr_ibov"])

# fig = make_subplots(
#     rows=3, cols=1,
#     shared_xaxes=True,
#     vertical_spacing=0.03,
#     specs=[[{"type": "table"}],
#            [{"type": "scatter"}],
#            [{"type": "scatter"}]]
# )

# # fig.add_trace(
# #     go.Scatter(
# #         x=df.index, 
# #         y=df['VaR Lee'],
# #         mode='lines',
# #         name='VaR Lee'
# #     ),
# #     row=3, col=1
# # )
# fig.add_trace(
#     go.Scatter(
#         x=df.index, 
#         y=df['Retorno Lee'],
#         mode='lines',
#         name='Retorno Lee'
#     ),
#     row=3, col=1
# )

# fig.add_trace(
#     go.Scatter(
#         x=df.index, 
#         y=df['VaR IBOV'],
#         mode='lines',
#         name='VaR IBOV'
#     ),
#     row=2, col=1
# )

# fig.add_trace(
#     go.Scatter(
#         x=df.index, 
#         y=df['retr_ibov'],
#         mode='lines',
#         name='retr_ibov'
#     ),
#     row=2, col=1
# )

# fig.add_trace(
#     go.Table(
#         header=dict(
#             values=["Data", "VaR Lee", "VaR IBOV",
#                     "Retorno Lee", "retr_ibov"],
#             font=dict(size=10),
#             align="left"
#         ),
#         cells = dict(
#             values=[df[k].tolist() for k in df.columns[0:]],
#             align = "left")
#     ),
#     row=1, col=1
# )

# fig.update_layout(
#     width=1500,
#     height=1500,
#     showlegend=False,
#     title_text="VaR Carlos Lee vs Benchmark (Ibov)",
# )

# fig.show()