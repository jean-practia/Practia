import pandas as pd
from IPython.display import display
from atualiza_lista import atualiza_dados
from datetime import datetime
from obtendoMesAr import mes_string


# variaveis
data_atual =datetime.today()
lista_colaboradores = []
lista_projeto = []

mes_atual_int = data_atual.month
ano_atual = data_atual.year
tabela = ''
mesString = mes_string()



def get_table():
    """Obtendo tabela atualizada"""
    tabela = pd.read_excel("input\\controleFinanceiropy.xlsx")
    
    tabela = pd.DataFrame(tabela)
    return tabela

def get_colaborador(tabela_colaborador):
    """Obtendo todos os colaboradores uma unica vez"""
    for x in tabela_colaborador:
        if x not in lista_colaboradores:
            lista_colaboradores.append(x)
    
    

def get_projetos(tabela_projetos):
    """Obtendo todos os prjetos uma unica vez"""
    for x in tabela_projetos:
        if x not in lista_projeto:
            lista_projeto.append(x)

# *******************************
# Function to insert row in the dataframe
def Insert_row(row_number, df, row_value):
    """Função para inseria dados no data frame(tabela)"""
    # Starting value of upper half
    start_upper = 0
   
    # End value of upper half
    end_upper = row_number
   
    # Start value of lower half
    start_lower = row_number
   
    # End value of lower half
    end_lower = df.shape[0]
   
    # Create a list of upper_half index
    upper_half = [*range(start_upper, end_upper, 1)]
   
    # Create a list of lower_half index
    lower_half = [*range(start_lower, end_lower, 1)]
   
    # Increment the value of lower half by 1
    lower_half = [x.__add__(1) for x in lower_half]
   
    # Combine the two lists
    index_ = upper_half + lower_half
   
    # Update the index of the dataframe
    df.index = index_
   
    # Insert a row at the end
    df.loc[row_number] = row_value
    
    # Sort the index labels
    df = df.sort_index()
   
    # return the dataframe
    return df
 
# *******************************

def filtra_tabela(tabela):
    """filtrando tabela deixando apenas um de cada colaborador"""
    try:
        soma = 0
        data = { 'Ano':[] ,
                    'Mês': [], 
                    'Nome': [],
                    'Cliente': '',
                    'HorasFeita': []
                    }
        segundo  = pd.DataFrame(data)
        # display(segundo)
        row_number = 0
        for colaborador in lista_colaboradores:
            primeiro = tabela.loc[tabela['Title']==colaborador,['Atividade','Mes','Title','Cliente','HorasFeita']]
            
            for cliente in lista_projeto:
                primeiro = primeiro.loc[primeiro['Cliente']==cliente,['Atividade','Mes','Title','Cliente','HorasFeita']]
                soma = primeiro['HorasFeita'].sum()
                
                if len(primeiro) > 0:
                    primeiro['HorasFeita'] = soma
                    # display(primeiro.drop_duplicates(subset='Title', keep='first'))
                    primeiro = primeiro.drop_duplicates(subset='Title', keep='first')
                    index = primeiro.index[0]
                   
                    # Let's create a row which we want to insert
                    row_value = [primeiro['Atividade'][index], primeiro['Mes'][index], primeiro['Title'][index],primeiro['Cliente'][index] ,primeiro['HorasFeita'][index]] 
                    # print(primeiro['Atividade'])
                    # print(row_value)
                    if row_number > segundo.index.max()+1:
                        print("Invalid row_number")
                    else:
                        # Let's call the function and insert the row
                        # at the second position
                        segundo = Insert_row(row_number, segundo, row_value)
                        
                        # Print the updated dataframe
                        # display(segundo)

                    row_number = row_number + 1 

                    # tabela_final.append(primeiro)

                primeiro = tabela.loc[tabela['Title']==colaborador,['Atividade','Mes','Title','Cliente','HorasFeita']]
    
    except Exception as e:
        raise e    

    return segundo

def filtra_ano_mes(tabela):

    tabela = tabela.loc[tabela['Atividade']==ano_atual,['Atividade','Mes','Title','Cliente','HorasFeita']]
    tabela = tabela.loc[tabela['Mes']==mesString,['Atividade','Mes','Title','Cliente','HorasFeita']]
    return tabela
    

def seleciona_mes():
    mes = ''
    print(mes)
    if mesString == "Janeiro":
        mes = "01/01/", str(ano_atual)
    elif mesString == "Fevereiro":
        mes = "01/02/", str(ano_atual)
    elif mesString == "Março":
        mes = "01/03/", str(ano_atual)
    elif mesString == "Abril":
        mes = "01/04/", str(ano_atual)
    elif mesString == "Maio":
        mes = "01/05/", str(ano_atual)
    elif mesString == "Junho":
        mes = "01/06/", str(ano_atual)
    elif mesString == "Julho":
        mes = "01/07/", str(ano_atual)
    elif mesString == "Agosto":
        mes = "01/08/", str(ano_atual)
    elif mesString == "Setembro":
        mes = "01/09/", str(ano_atual)
    elif mesString == "Outubro":
        mes = "01/10/"+ str(ano_atual)
    elif mesString == "Novembro":
        mes = "01/11/", str(ano_atual)
    elif mesString == "Dezembro":
        mes = "01/12/", str(ano_atual)
    else: 
        mes = "data invalida"

    return mes


def main():
    try:
        base_atualizada = atualiza_dados()

        if base_atualizada:

            tabela = get_table()
            tabela = filtra_ano_mes(tabela)
            # obtendo colaboradores
            get_colaborador(tabela['Title'])
            # obtendo projetos
            get_projetos(tabela['Cliente'])

            table = filtra_tabela(tabela)
            mese = seleciona_mes()

            table['Mês'] = mese
            # table['Ano'] = ano_atual

            table.to_excel("ControleHoras.xlsx")
        else:
            print('erro!')
    except Exception as e:
        print(e)
    finally:
        print('Finalizado')

main()