import win32com.client
import time

path_practia ='C:\\Users\\USER\\Documents\\Estudos\\python\\financeiro\\input\\controleFinanceiropy.xlsx'
path_jean ='C:\\Users\\jean\\arquivos\\Documentos\\practia\\Practia\\financeiro\\input\\controleFinanceiropy.xlsx'

def atualiza_dados():
    """Obtendo os dados do sharepoint para uma planilha"""
    sucesso = False
    try:
        xlapp = win32com.client.DispatchEx("Excel.Application")
        wb = xlapp.Workbooks.Open(path_jean)
        wb.RefreshAll()

        xlapp.CalculateUntilAsyncQueriesDone()
        time.sleep(5)

        xlapp.DisplayAlerts = False
        wb.Save()
        print("Sucesso!")
        print("Relatório atualizado")
        sucesso = True
    except Exception as e:
        print('Error:')
        print(e)
    finally: 
        wb.Close()
        xlapp.Quit()
        
    return sucesso