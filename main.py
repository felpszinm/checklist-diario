import os
import pandas as pd
from datetime import datetime

#* Caminho da Planilha Mãe:
BASE_PATH = os.path.join('planilhas-checklist-alc', 'Planilha-Base.xlsx')


class LeitorDePlanilhas:

    def __init__(self, planilhas_auxiliares, caminho_planilha_base=BASE_PATH):
        self.caminho_base = caminho_planilha_base
        self.planilhas_auxiliares = planilhas_auxiliares
        self.df_base = pd.DataFrame()
        self.df_colunas_filtradas = {}


    def ler_planilha_base(self, sheet_name=0):
            
        if os.path.exists(self.caminho_base):
                self.df_base = pd.read_excel(self.caminho_base, sheet_name=sheet_name, engine='openpyxl')

        else:
            print(f'Planilha Base não encontrada em {self.caminho_base}.')
            
    
    def unificar_dados(self):

        for planilha in self.planilhas_auxiliares:
            caminho_planilha_auxiliar = planilha['caminho']
            destino = planilha['aba_destino']
            colunas = planilha['colunas']

            if not os.path.exists(caminho_planilha_auxiliar):
                print(f'Arquivo {caminho_planilha_auxiliar} NÃO encontrado. Indo para a próxima.')
                continue
            
            try:
                df = pd.read_excel(caminho_planilha_auxiliar, engine='openpyxl')
                self.df_colunas_filtradas[destino] = df[colunas]
            
            except Exception as error:
                print(f'Erro ao ler o arquivo {caminho_planilha_auxiliar}: {error}')
                continue


    def salvar_planilha_base(self):
        
        if not self.df_colunas_filtradas:
            print('Nenhum dado a ser salvo, o dicionário está vazio.')
            return

        with pd.ExcelWriter(self.caminho_base, engine='openpyxl') as escritor:
            for nome_aba, df in self.df_colunas_filtradas.items():

                if not df.empty:
                    df.to_excel(escritor, sheet_name=nome_aba, index=False)
                    print(f'Dados salvos na aba {nome_aba}')

                else:
                    print(f'A aba {nome_aba} está vazia.')


if __name__ == '__main__':
    
    planilha_base = {
        'caminho': 'planilhas-checklist-alc/Planilha_Base.xlsx'
        }

    planilhas_auxiliares = [
        {
            'caminho': 'planilhas-checklist-alc/Planilha-DDS.xlsx',
            'aba': 'Geral',
            'colunas': ['Base', 'Data', 'Placa', 'Motorista'],
            'aba_destino': 'DDS'
        },
        {
            'caminho': 'planilhas-checklist-alc/Planilha-Prolog.xlsx',
            'aba': 'Planilha-Prolog',
            'colunas': ['UNIDADE', 'DATA REALIZAÇÃO', 'COLABORADOR', 'PLACA'],
            'aba_destino': 'PROLOG'
        },
        {
            'caminho': 'planilhas-checklist-alc/Planilha-VecFleet.xlsx',
            'aba': 'Worksheet',
            'colunas': ['Fecha y hora de Creacion', 'Base', 'Movil', 'Sub-Región'],
            'aba_destino': 'VECFLEET'  
        }
    ]

    gestor_frotas = LeitorDePlanilhas(planilhas_auxiliares)
    gestor_frotas.ler_planilha_base()
    gestor_frotas.unificar_dados()
    #gestor_frotas.salvar_planilha_base()
    print(gestor_frotas.df_colunas_filtradas)



"""
    TODO: Planilha-Base
    -> Criar um método de válidação de linhas de ambas as planilhas dentro de 'Planilha-Base';
    -> Se caso a falta de informações na parte de DDS/PROLOG/VEC FLEET, pegar as informações e colocar diretamente dentro de ambas
    -> Ex: as informações do dia 14/08/2025 não estiver nela mas estiver nas outras, passe as informações do dia 14/08/2025 pra 'Planilha-Base'
#*  -> Dica: Usar o Datetime(Data de Hoje / Now()) para data a data que estiver faltando.

?   Avisos:
    -> Utilização da lib. Pandas, os, datetime e POO.
 """   