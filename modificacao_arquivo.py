import os
import time
import pandas as pd
import win32com.client as win32

def send_email(local_time):

    # enviar email com os dados do csv

    outlook = win32.Dispatch("outlook.application") # criando integração com o outlook

    email = outlook.CreateItem(0) # criando um email

    email.To = "raul_lopes.camina@daimler.com" # especificando o destino do email

    email.Subject = "Horários modificados" # especificando o assunto no email

    email.HTMLBody = '''<h3>Horário de atualização do arquivo TTL:</h3>
                       {}'''.format(local_time) #com esse comando .format, é possível enviar um dataframe no modo tabela por email

    email.Send()

modification_time = os.path.getmtime(r"R:\FTP\PLV\TLLPontosNPsEIXO.txt")
local_time = time.ctime(modification_time)


while True:
    modification_time_updated = os.path.getmtime(r"R:\FTP\PLV\TLLPontosNPsEIXO.txt")
    local_time_updated = time.ctime(modification_time_updated)

    if local_time != local_time_updated:
        send_email(local_time_updated)
        local_time = local_time_updated

    time.sleep(300)