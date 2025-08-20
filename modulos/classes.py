import pandas as pd
import os
from openpyxl import load_workbook

#* Caminho da Planilha Mãe:
BASE_PATH = os.path.join('planilhas-checklist-alc', 'Planilha-Base.xlsx')

class LeitorDePlanilhas:

    #* Inicializo minha classe com *args das minhas planilhas filhas e o caminho da planilha base.
    def __init__(self, planilhas_auxiliares, caminho_planilha_base=BASE_PATH):
        #* Criação dos camimhos e do dataframe.
        self._caminho_base = caminho_planilha_base
        self._planilhas_auxiliares = planilhas_auxiliares
        self._df_base = pd.DataFrame()
        self.df_colunas_filtradas = {}


    #* Método de leitura onde ele vai verificar se planilha base existe.
    def ler_planilha_base(self, sheet_name=0):
        if os.path.exists(self._caminho_base):
                self._df_base = pd.read_excel(self._caminho_base, sheet_name=sheet_name, engine='openpyxl')
        else:
            print(f'Planilha Base não encontrada em {self._caminho_base}.')
            
    #* Método para unificar os dados de planilhas filhas dentro da planilha base.
    def unificar_dados(self):
        #* Para cada 'planilha' em 'planilha_auxiliares'
        for planilha in self._planilhas_auxiliares:
            caminho_planilha_auxiliar = planilha['caminho']
            destino = planilha['aba_destino']
            colunas = planilha['colunas']
            cabecario = planilha['cabecario']

            #* Se o caminho nao tiver as planilhas ele pula e vai pra proxima.
            if not os.path.exists(caminho_planilha_auxiliar):
                print(f'\033[1;91mArquivo {caminho_planilha_auxiliar} NÃO encontrado. Indo para a próxima.')
                continue
            
            #* Tenta fazer a leitura das planilhas filhas com o openpyxl, e as manda as colunas de 'df' para 'df_colunas_filtradas'
            try:
                df = pd.read_excel(caminho_planilha_auxiliar, sheet_name=planilha['indice_aba'], engine='openpyxl', header=cabecario)
                self.df_colunas_filtradas[destino] = df[colunas]

            #* Se caso ocorrer algum erro de 'IndentationError' ou 'IndexError' ele ignora o arquivo.
            except Exception as error: 
                print(f'❌ \033[91mErro ao ler o arquivo {caminho_planilha_auxiliar}: {error} ❌')
                continue

    #* Salva as informações em planilha base.
    def salvar_planilha_base(self):
        #* Se caso df_colunas_filtradas não ter informações ele não salva nenhum dado.
        if not self.df_colunas_filtradas:
            print('Nenhum dado a ser salvo, o dicionário está vazio.')

        #* Salva os arquivos dentro de df_colunas_filtradas utilizando o 'ExcelWriter'
        with pd.ExcelWriter(self._caminho_base, engine='openpyxl') as escritor:
            for nome_aba, df in self.df_colunas_filtradas.items():

                #* Verifica se df está vazio ou não.
                if not df.empty:
                    df.to_excel(escritor, sheet_name=nome_aba, index=False)
                    print(f'Dados salvos na aba {nome_aba}')

                else:
                    print(f'A aba {nome_aba} está vazia.')

#* Classe para escrever as formulas:
class EditorDePlanilha:
    def __init__(self, caminho_planilha_base=BASE_PATH):
        self.caminho_base = caminho_planilha_base
        self.planilha_base = load_workbook(BASE_PATH, data_only=False)
        self.frota_fixa = self.planilha_base['FROTA FIXA']
        self.dds = self.planilha_base['DDS']

    #* Abro a planilha com as formulas:
    def incluir_formulas(self):
        index = 2
        
        #* FORMULAS A SEREM INCLUIDAS EM FROTA FIXA:
        formula_DDS = '='

        #* FORMULAS A SEREM INCLUIDAS EM FROTA FIXA:
        placas_frota_fixa = self.frota_fixa[f'B{index}'].value
        meli_frota_fixa = self.frota_fixa[f'E{index}'].value
        disponibilidade_frota_fixa = self.frota_fixa[f'F{index}'].value
        formula_ff_meli = f'=SE(ÉERROS(PROCV({index};Meli!G:G;1;FALSO));"NÃO";"SIM")'
        formula_ff_disponibilidade = f'=PROCV(@B:B;BASE!C:J;8;0)'

# TODO: PEGAR AS COLUNAS E, F e colocar na ordem formula meli e formula disponibilidade.
# TODO: fazer até o ultimo valor da coluna B (placas).