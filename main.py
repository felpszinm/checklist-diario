import os
import pandas as pd
from time import sleep

#* Caminho da Planilha Mãe:
BASE_PATH = os.path.join('planilhas-checklist-alc', 'Planilha-Base.xlsx')


class LeitorDePlanilhas:

    #* Inicializo minha classe com *args das minhas planilhas filhas e o caminho da planilha base.
    def __init__(self, planilhas_auxiliares, caminho_planilha_base=BASE_PATH):
        #* Criação dos camimhos e do dataframe.
        self.caminho_base = caminho_planilha_base
        self.planilhas_auxiliares = planilhas_auxiliares
        self.df_base = pd.DataFrame()
        self.df_colunas_filtradas = {}


    #* Método de leitura onde ele vai verificar se planilha base existe.
    def ler_planilha_base(self, sheet_name=0):
        if os.path.exists(self.caminho_base):
                self.df_base = pd.read_excel(self.caminho_base, sheet_name=sheet_name, engine='openpyxl')
        else:
            print(f'Planilha Base não encontrada em {self.caminho_base}.')
            
    #* Método para unificar os dados de planilhas filhas dentro da planilha base.
    def unificar_dados(self):
        #* Para cada 'planilha' em 'planilha_auxiliares'
        for planilha in self.planilhas_auxiliares:
            caminho_planilha_auxiliar = planilha['caminho']
            destino = planilha['aba_destino']
            colunas = planilha['colunas']

            #* Se o caminho nao tiver as planilhas ele pula e vai pra proxima.
            if not os.path.exists(caminho_planilha_auxiliar):
                print(f'Arquivo {caminho_planilha_auxiliar} NÃO encontrado. Indo para a próxima.')
                continue
            
            #* Tenta fazer a leitura das planilhas filhas com o openpyxl, e as manda as colunas de 'df' para 'df_colunas_filtradas'
            try:
                df = pd.read_excel(caminho_planilha_auxiliar, sheet_name=planilha['indice_aba'], engine='openpyxl')
                self.df_colunas_filtradas[destino] = df[colunas]

            #* Se caso ocorrer algum erro de 'IndentationError' ou 'IndexError' ele ignora o arquivo.
            except Exception as error: 
                print(f'Erro ao ler o arquivo {caminho_planilha_auxiliar}: {error}')
                continue

    #* Salva as informações em planilha base.
    def salvar_planilha_base(self):
        #* Se caso df_colunas_filtradas não ter informações ele não salva nenhum dado.
        if not self.df_colunas_filtradas:
            print('Nenhum dado a ser salvo, o dicionário está vazio.')

        #* Salva os arquivos dentro de df_colunas_filtradas utilizando o 'ExcelWriter'
        with pd.ExcelWriter(self.caminho_base, engine='openpyxl') as escritor:
            for nome_aba, df in self.df_colunas_filtradas.items():

                #* Verifica se df está vazio ou não.
                if not df.empty:
                    df.to_excel(escritor, sheet_name=nome_aba, index=False)
                    print(f'Dados salvos na aba {nome_aba}')

                else:
                    print(f'A aba {nome_aba} está vazia.')



#* Inicialização do Programa:
if __name__ == '__main__':
    
    planilha_base = {
        'caminho': 'planilhas-checklist-alc/Planilha_Base.xlsx',
        'indice_aba': 1
        }

    planilhas_auxiliares = [
        {
            'caminho': 'planilhas-checklist-alc/Planilha-DDS.xlsx',
            'aba': 'Geral',
            'indice_aba': 0,
            'colunas': ['Base', 'Data', 'Placa', 'Motorista'],
            'aba_destino': 'DDS'
        },
        {
            'caminho': 'planilhas-checklist-alc/Planilha-Prolog.xlsx',
            'aba': 'Planilha-Prolog',
            'indice_aba': 0,
            'colunas': ['UNIDADE', 'DATA REALIZAÇÃO', 'COLABORADOR', 'PLACA'],
            'aba_destino': 'PROLOG'
        },
        {
            'caminho': 'planilhas-checklist-alc/Planilha-VecFleet.xlsx',
            'aba': 'Worksheet',
            'indice_aba': 0,
            'colunas': ['Fecha y hora de Creacion', 'Base', 'Movil', 'Sub-Región'],
            'aba_destino': 'VECFLEET'  
        },
        {
            'caminho': 'planilhas-checklists-dispatchers/Planilha-VecFleetDispatcher.xlsx',
            'aba': 'Frota Fixa',
            'indice_aba': 0,
            'colunas': ['SIGLAS', 'PLACA', 'BASE', 'COORDENADOR', 'TEM NO MELI?', 'Coluna1'],
            'aba_destino': 'FROTA FIXA'
        },
        {
            'caminho': 'planilhas-checklists-dispatchers/Planilha-VecFleetDispatcher.xlsx',
            'aba': 'Meli',
            'indice_aba': 1,
            'colunas': ['Base', 'Movil'],
            'aba_destino': 'MELI'
        },
        {
            'caminho': 'planilhas-checklists-dispatchers/Planilha-VecFleetDispatcher.xlsx',
            'aba': 'BASE',
            'indice_aba': 2,
            'colunas': ['PLACA', 'BASE', 'COORDENADOR', 'RESPONSÁVEIS FROTA', 'DISPONIBILIDADE'],
            'aba_destino': 'BASE'
        },
        {
            'caminho': 'planilhas-checklists-dispatchers/Planilha-Quinzenal-VecFleet.xlsx', 
            'aba': 'Worksheet',
            'indice_aba': 0,
            'colunas': ['Base', 'Movil', 'Sub-Región'],
            'aba_destino': 'QUINZENAL VECFLEET'
        }
    ]

    gestor_frotas = LeitorDePlanilhas(planilhas_auxiliares)
    gestor_frotas.ler_planilha_base()
    gestor_frotas.unificar_dados()
    gestor_frotas.salvar_planilha_base()
    print('Dados salvos na planilha BASE!')
    print()
    print('Programa finalizado!')
    sleep(2)

#TODO: Criar uma validação das formulas do excel para certas colunas.
#TODO: Mante-las dentro da planilha base.