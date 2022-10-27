from datetime import datetime


def mes_string():
    today = datetime.today()
    days = ['Segunda-feira', 'Terça-feira', 'Quarta-feira', 'Quinta-feira', 'Sexta-feira', 'Sábado', 'Domingo']
    months = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro']

    str_month = months[today.month-1] # obtemos o numero do mês e subtraímos 1 para que haja a correspondência correta com a nossa lista de meses  
    str_weekday = days[today.weekday()] # obtemos o numero do dia da semana, neste caso o numero 0 é segunda feira e coincide com os indexes da nossa lista

    # print(f'{str_weekday}, {today.day} de {str_month} de {today.year}') # Quinta-feira, 11 de Agosto de 2022
    return str_month

