import pandas as pd
import numpy as np
from pymcdm.methods import PROMETHEE_II
from pymcdm.helpers import rrankdata
from geographiclib.geodesic import Geodesic


geod = Geodesic.WGS84
# calculo =  geod.Inverse(Lat_critico,Lon_critico,row[coluna_lat_solucoes],row[coluna_lon_solucoes])
# distancia = calculo['s12']


def CalculaDistancia(lat1,long1,lat2,long2):
    '''
    Calcula a distancia em KM
    '''
    calculo =  geod.Inverse(lat1 ,long1, lat2,long2)
    distancia = calculo['s12']
    
    return round((distancia / 1000),2) # retorno em Km

def PrometheeRecomendacao(df,
                          NumMaxRanking,
                          NumRecomendacoes,
                          PesoDistanciaRecomendacao,
                          qRecomendacao,
                          pRecomendacao,
                          PrefixoArqSaida,
                          NomeArquivoOriginal):
    '''
    Executa o promethee e depois o sistema de recomendacao
    '''
    
    weights = np.array(df.iloc[0,1:-4].astype(float)) # SOMA 1
    types = np.array(df.iloc[1,1:-4].astype(float)) # 1 quanto maior melhor.
                        #-1, quanto menor, menor.                       
    q = np.array(df.iloc[2,1:-4].astype(float)) # q ushape PROMETHEE II
    p = np.array(df.iloc[3,1:-4].astype(float)) # p ushape PROMETHEE II
    valores = np.array(df.iloc[4:,1:-4].astype(float)) # tirando a coluna com as OCs         

    pref = PROMETHEE_II('vshape_2')
    pref = pref(valores, weights, types, q=q, p=p)
    ranking = rrankdata(pref)

    novoArquivo = PrefixoArqSaida + '_' + NomeArquivoOriginal.split('/')[-1]
    NomeArquivoOriginal.split('/')[-1]
    path = NomeArquivoOriginal.split('/')
    path= path[:-1]
    path = "/".join(path)
    novoArquivo = path + "/" + novoArquivo

    df_temp = df.iloc[4:,]
    df_temp['ranking'] = ranking
    df_temp.to_excel(novoArquivo, index= False)
    
    ### Parte II Rodar o PROMETHEE II PARA RECOMENDAR SITES
    
    
    
    return np.nan



NumMaxRanking = 150
NumRecomendacoes = 2
PesoDistanciaRecomendacao = 0.50
PesoRankingRecomendacao = 1-PesoDistanciaRecomendacao
qRecomendacao = 10 # km
pRecomendacao = 50 # km
PrefixoArqSaida = 'Prefixo1'
PrefixoRecomendacao = "Recomendacao_" + PrefixoArqSaida
NomeArquivoOriginal = r'C:\Users\f8058552\OneDrive - TIM\__Automacao_de_tarefas\Projetos_em_python\___DECISAO_MULTICRITERIO\priorizar_q2_netflix\priorizar.xlsx'
nomeTemp = r'C:\Users\f8058552\OneDrive - TIM\__Automacao_de_tarefas\Projetos_em_python\___DECISAO_MULTICRITERIO\priorizar_q2_netflix\semRecom_priorizar.xlsx'
#df = pd.read_excel(NomeArquivoOriginal) 

df_passo1 = pd.read_excel(nomeTemp)

# passo 1. tirar colunas q nao interessam
df_passo1 = df_passo1.filter(['Site','ranking','lat','long','vendor'])

df_passo1 = df_passo1.sort_values(by='ranking', ascending= True)

# passo 2. separar o ranking em 2.
# OCs Principais foram as selecionadas
# OCs Fora Ranking são as demais
dfOCsPrincipais = df_passo1.iloc[0:NumMaxRanking,] 
dfOCsForaRanking = df_passo1.iloc[NumMaxRanking:,]

# passo 3. PRINCIPAL
# Iterar o dfOCsPrinciais, filtranto pelo Vendor e depois
# calculando a distancia de cada um deles ao dfOCsForaRanking
# e rodar o Promethee, Salvar os NumRecomendacoes em um df

Colunas = ['Site',
           'rankingGlobal',
           'distancia',
           'rankingLocal',
           'OCPrincipal']

list_saida=[]

dfRecomendacoes = pd.DataFrame( columns=Colunas)
dfOCsForaRanking_copy = dfOCsForaRanking.copy(deep=True)
#listaOCsPrincipais = dfOCsPrincipais.iloc[:,0]

for index, row in dfOCsPrincipais.iterrows():
    NEWROW=[]
    print(index)
    latOCPrin = row['lat']
    longOCPrin = row['long']
    vendorOCPrin = row['vendor']
    OCPrincipal = row['Site']
    df_temp = pd.DataFrame()
    df_temp = dfOCsForaRanking_copy.loc[dfOCsForaRanking_copy['vendor']==vendorOCPrin]
    
    df_temp['distancia'] = df_temp.apply(lambda x: CalculaDistancia(
                                                                    latOCPrin,
                                                                    longOCPrin,
                                                                    x['lat'],
                                                                    x['long']
                                                                    ), axis=1)

    df_temp = df_temp.filter(['Site','ranking','distancia'])
    
    weights = np.array([PesoRankingRecomendacao, PesoDistanciaRecomendacao])
    types = np.array([-1,-1]) # menor ranking e distancia, melhor                     
    q = np.array([0,qRecomendacao]) # q ushape PROMETHEE II. zero pro ranking e 10km pra disy
    p = np.array([0,pRecomendacao]) # p ushape PROMETHEE II. 50km pra dist. Vshape2
    valores = np.array(df_temp.iloc[:,1:].astype(float)) # tirando a coluna com as OCs         

    pref = PROMETHEE_II('vshape_2')
    pref = pref(valores, weights, types, q=q, p=p)
    ranking = rrankdata(pref)
    df_temp['rankingLocal'] = ranking
    df_temp = df_temp.sort_values(by='rankingLocal', ascending= True)
    
    df_temp = df_temp.iloc[0:NumRecomendacoes,]  #corte nas recomendacoes pedidas
    df_temp['OCPrincipal'] = OCPrincipal
    
    # remover as OCs recomendadas do pool
    OCsRecomendadas =  list(df_temp['Site'])
    dfOCsForaRanking_copy= dfOCsForaRanking_copy.query('Site not in @OCsRecomendadas')
    df_temp.rename(columns={'ranking':'rankingGlobal'}, inplace = True)
    
    #  NEWROW = [endid, site, tech[1], indicador[1], -1, -1, -1, -1, -1, resultado]
    #                 list_saida = list_saida + [NEWROW]
    #NEWROW = list(df_temp.values) + [NEWROW]
    dfRecomendacoes = dfRecomendacoes.append(df_temp, ignore_index=True)


#dfRecomendacoes = pd.DataFrame(data=NEWROW, columns=Colunas)
dfRecomendacoes.rename(columns={'Site':'OCRecomendada'}, inplace = True)
dfRecomendacoes.to_excel('OCsRecomendadas.xlsx', index=False)