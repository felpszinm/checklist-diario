from openpyxl import load_workbook
from modulos.dicts import planilhas_auxiliares
from modulos.classes import LeitorDePlanilhas, EditorDePlanilha, BASE_PATH
from time import sleep

#* Inicialização do Programa:
if __name__ == '__main__':

    gestor_frotas = LeitorDePlanilhas(planilhas_auxiliares)
    editor_frotas = EditorDePlanilha(caminho_planilha_base=BASE_PATH)
    gestor_frotas.ler_planilha_base()
    gestor_frotas.unificar_dados()
    gestor_frotas.salvar_planilha_base()
    print('\n\033[1;33mTodos os dados do pandas foi salvo na Planilha Base!\n')
    sleep(2)
    print('\033[mIncluindo as fórmulas no seu programa...')
    #TODO: Colocar os métodos do editor de planilha.
    print('\033[1;33mFórmulas salvas!')
    sleep(2)
    print('\n\033[1;32mPrograma Concluído ✅')