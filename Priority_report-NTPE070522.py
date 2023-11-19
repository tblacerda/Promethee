import pandas as pd
import numpy as np
from pymcdm.methods import PROMETHEE_II
from pymcdm.helpers import rrankdata
import PySimpleGUI as sg
import warnings
warnings.filterwarnings('ignore')
from geographiclib.geodesic import Geodesic
geod = Geodesic.WGS84


def SplashScreen():
    
    """
    Demo - Splash Screen
    
    Displays a PNG image with transparent areas as see-through on the center
    of the screen for a set amount of time.
    
    Copyright 2020 PySimpleGUI.org
"""

    FILENAME = r'./splash.PNG'          # if you want to use a file instead of data, then use this in Image Element
    DISPLAY_TIME_MILLISECONDS = 4000
    
    sg.Window('Window Title', [[sg.Image(filename=FILENAME)]],transparent_color=sg.theme_background_color(), no_titlebar=True, keep_on_top=True).read(timeout=DISPLAY_TIME_MILLISECONDS, close=True)

    return 1

def CheckPlanilha(df, SisRecomen):
    '''
    Verifica se a planilha tem
    a primeira linha 'peso'
    a segunda linha 'tipo'
    a terceira linha 'q'
    a quarta linha 'p'
    a ultima coluna 'ranking'
    '''
    #print('check')
    TudoOk = False
    print(SisRecomen)

    if SisRecomen == False:
        if (df.iloc[0,0] == 'peso' and
            df.iloc[1,0] == 'tipo' and
            df.iloc[2,0] == 'q' and
            df.iloc[3,0] == 'p' and
            df.columns[-1] == 'ranking'):
            TudoOk = True

        return TudoOk
    elif SisRecomen == True:
        if (df.iloc[0,0] == 'peso' and
            df.iloc[1,0] == 'tipo' and
            df.iloc[2,0] == 'q' and
            df.iloc[3,0] == 'p' and
            df.columns[-3] == 'lat' and
            df.columns[-2] == 'long' and
            df.columns[-1] == 'vendor'):
            TudoOk = True
            #print('Tudo ok', TudoOk)
    
    return TudoOk


####

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
                          NomeArquivoOriginal,
                          UsarRankingSeparado,
                          ArqRankingSeparado):
    '''
    Executa o promethee e depois o sistema de recomendacao
    '''
    ## Teste
    #print('PROMETHERECOMENDACAO')
    # NumMaxRanking = 150
    # NumRecomendacoes = 2
    # PesoDistanciaRecomendacao = 0.50
    # qRecomendacao = 10 # km
    # pRecomendacao = 50 # km
    # PrefixoArqSaida = 'Prefixo1'
    # PrefixoRecomendacao = "Recomendacao_" + PrefixoArqSaida
    # NomeArquivoOriginal = r'C:\Users\tblac\OneDrive\___DECISAO_MULTICRITERIO\priorizar_q2_netflix\priorizar.xlsx'
    # #nomeTemp = r'C:\Users\f8058552\OneDrive - TIM\__Automacao_de_tarefas\Projetos_em_python\___DECISAO_MULTICRITERIO\priorizar_q2_netflix\semRecom_priorizar.xlsx'
    # UsarRankingSeparado = True
    # ArqRankingSeparado = r'C:\Users\tblac\OneDrive\___DECISAO_MULTICRITERIO\priorizar_q2_netflix\ranking_NI.xlsx'
    # df = pd.read_excel(NomeArquivoOriginal)
    
    ## Fim do Teste

    PesoRankingRecomendacao = 1-PesoDistanciaRecomendacao
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
#    NomeArquivoOriginal.split('/')[-1]
    path = NomeArquivoOriginal.split('/')
    path= path[:-1]
    path = "/".join(path)
    novoArquivo = path + "/" + novoArquivo
    ArquivoRecomendacao = path + "/" + PrefixoArqSaida + '_Recomendacao_'+ NomeArquivoOriginal.split('/')[-1]
    
    df_passo1 = df.iloc[4:,]
    df_passo1['ranking'] = ranking
    df_passo1.to_excel(novoArquivo, index= False)
    

    ### Parte II Rodar o PROMETHEE II PARA RECOMENDAR SITES
    
    # passo 1. tirar colunas q nao interessam
    df_passo1 = df_passo1.filter(['Site','ranking','lat','long','vendor'])


    df_passo1 = df_passo1.sort_values(by='ranking', ascending= True)

    # passo 2. separar o ranking em 2.
    # OCs Principais foram as selecionadas
    # OCs Fora Ranking são as demais
    dfOCsPrincipais = df_passo1.iloc[0:NumMaxRanking,] 
    dfOCsForaRanking = df_passo1.iloc[NumMaxRanking:,]

    # Passo 2.5 Verificar se é pra utilizar um ranking separado. Se sim
    # ler o arquivo e atualizar o ranking
    if UsarRankingSeparado == True:
        try:
            dfRankingSeparado = pd.read_excel(ArqRankingSeparado)
            dfOCsForaRanking= dfOCsForaRanking.drop(['ranking'], axis=1)
            dfOCsForaRanking = pd.merge(left=dfOCsForaRanking, right=dfRankingSeparado, on='Site')        
            dfOCsForaRanking = dfOCsForaRanking.sort_values(by='ranking', ascending=True)
        except:
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

    dfRecomendacoes = pd.DataFrame( columns=Colunas)
    dfOCsForaRanking_copy = dfOCsForaRanking.copy(deep=True)
    #listaOCsPrincipais = dfOCsPrincipais.iloc[:,0]
    for index, row in dfOCsPrincipais.iterrows():
        #print(index)
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
        dfRecomendacoes = dfRecomendacoes.append(df_temp, ignore_index=True)
        
    dfRecomendacoes.rename(columns={'Site':'OCRecomendada'}, inplace = True)
    dfRecomendacoes.to_excel(ArquivoRecomendacao, index=False)
    return np.nan
####

def ChamaPromethee(df,values):
    '''
        Função que chama o Promethee e escreve no excel
    '''

    weights = np.array(df.iloc[0,1:-1].astype(float)) # SOMA 1
    types = np.array(df.iloc[1,1:-1].astype(float)) # 1 quanto maior melhor.
                        #-1, quanto menor, menor.                       
    q = np.array(df.iloc[2,1:-1].astype(float)) # q ushape PROMETHEE II
    p = np.array(df.iloc[3,1:-1].astype(float)) # p ushape PROMETHEE II
    valores = np.array(df.iloc[4:,1:-1].astype(float)) # tirando a coluna com as OCs         

    pref = PROMETHEE_II('vshape_2')
    pref = pref(valores, weights, types, q=q, p=p)
    ranking = rrankdata(pref)

    novoArquivo = values[8] + '_' + values[1].split('/')[-1]
    values[1].split('/')[-1]
    path = values[1].split('/')
    path= path[:-1]
    path = "/".join(path)
    novoArquivo = path + "/" + novoArquivo

    df_temp = df.iloc[4:,]
    df_temp['ranking'] = ranking
    df_temp.to_excel(novoArquivo, index= False)

    return np.nan


def main():
    SplashScreen()

    layout = [
            [sg.Image(r'priorityreport.png')],
            [sg.Text('Escolha o arquivo')],[sg.Input(), sg.FileBrowse()],
            [sg.Checkbox(' Recomendação', default=False),
            sg.Text('Total OCs Prioritárias:'),sg.Input(size=4),
            sg.Text('OCs a recomendar:'), sg.Input(size=2),
            sg.Text('Peso da distância (%):'), sg.Input(size=2),
            sg.Text('q (km):'), sg.Input(size=2),
            sg.Text('p (km):'),sg.Input(size=2)],
            [sg.Text('Prefixo de saida'),sg.InputText(size=10)],
            [sg.Checkbox('Ranking separado para Recomendação', default=False),sg.Text('Escolha o arquivo:'),sg.Input(), sg.FileBrowse()],
            [sg.OK("Executar"), sg.Cancel("Fechar"), sg.Button("Ajuda"), sg.Button("Sobre...")]]

    windowPrincipal = sg.Window('Priority Report - PROMETHEE II', layout)


    while True:
        event, values = windowPrincipal.read()
       # windowPrincipal.refresh()
        try:
            NomeArquivoOriginal = values[1] # nome do arquivo 
            SistemaRecomendacao = values[2] # checkbox do sistema de recomendação (True/False)
            NumMaxRanking = int(values[3]) # para o sis de recomendacao precisa ter o corte de quais OCs sao as principais. É esse valor
            NumRecomendacoes = int(values[4]) # numero de recomendações, para cada OC principal
            PesoDistanciaRecomendacao = int(values[5]) # peso da distancia (peso do ranking sera 1-values[3])
            PesoDistanciaRecomendacao = float(PesoDistanciaRecomendacao/100)
            qRecomendacao = int(values[6]) # q, indiferenca, em km
            pRecomendacao = int(values[7]) # p, valor da rampa, em km
            PrefixoArqSaida = values[8] # prefixo
            UsarRankingSeparado = values[9] 
            ArqRankingSeparado = values[10]
        except:
            NomeArquivoOriginal = values[1] # nome do arquivo 
            SistemaRecomendacao = values[2] # checkbox do sistema de recomendação (True/False)
            PrefixoArqSaida = values[8] # prefixo


        if event == "Fechar" or event == sg.WIN_CLOSED:
            break
        if event == "Executar":
            try:
                df = pd.read_excel(NomeArquivoOriginal)
                
                if CheckPlanilha(df, SistemaRecomendacao) == False:
                
                    layout = [[sg.Text('Erro na leitura do arquivo. Cheque o formato')],
                    [sg.OK()]]
                    windowErro = sg.Window('Priority Report - PROMETHEE II', layout, modal =True)
                    eventErro, valuesErro = windowErro.read()
                    windowErro.close()
                    
                elif SistemaRecomendacao == False:
                    ChamaPromethee(df,values)
                    layout = [[sg.Text('executado')],
                    [sg.OK()]]
                    windowExecutado = sg.Window('Priority Report - PROMETHEE II', layout, modal =True)
                    eventExecutado, valuesExecutado = windowExecutado.read()
                    windowExecutado.close()
                
                elif SistemaRecomendacao == True:
                    PrometheeRecomendacao(df,
                                          NumMaxRanking,
                                          NumRecomendacoes,
                                          PesoDistanciaRecomendacao,
                                          qRecomendacao,
                                          pRecomendacao,
                                          PrefixoArqSaida,
                                          NomeArquivoOriginal,
                                          UsarRankingSeparado,
                                          ArqRankingSeparado
                                          )
                    
                    layout = [[sg.Text('executado')],
                    [sg.OK()]]
                    windowExecutado = sg.Window('Priority Report - PROMETHEE II', layout, modal =True)
                    eventExecutado, valuesExecutado = windowExecutado.read()
                    windowExecutado.close()




            except:
                break

    windowPrincipal.close()

if __name__ == "__main__":

# arquivo = r'C:\Users\f8058552\OneDrive - TIM\__Automacao_de_tarefas\Projetos_em_python\___DECISAO_MULTICRITERIO\priorizar_q2_netflix\priorizar.xlsx'

# df=pd.read_excel(arquivo)
# CheckPlanilha(df, )
    main()


