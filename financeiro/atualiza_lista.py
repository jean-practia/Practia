import win32com.client
import time

def atualiza_dados():
    """Obtendo os dados do sharepoint para uma planilha"""
    sucesso = False
    try:
        xlapp = win32com.client.DispatchEx("Excel.Application")
        wb = xlapp.Workbooks.Open('C:\\Users\\USER\\Documents\\Estudos\\python\\financeiro\\input\\controleFinanceiropy.xlsx')
        wb.RefreshAll()

        xlapp.CalculateUntilAsyncQueriesDone()
        time.sleep(5)

        xlapp.DisplayAlerts = False
        wb.Save()
        print("Sucesso!")
        sucesso = True
    except Exception as e:
        print('Error:')
        print(e)
    finally: 
        wb.Close()
        xlapp.Quit()
        
    return sucesso