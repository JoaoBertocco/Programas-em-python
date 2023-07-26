import win32com.client as win32
import pandas as pd

#Lendo os dados de venda de um arquivo CSV
sales_data = pd.read_csv("sales.csv")

#Processando os dados de vendas
total_sales = sales_data["sales"].sum()
average_sales = sales_data["sales"].mean()
max.sales = sales_data["sales"].max()
min_sales = sales_data["sales"].min()

#Criando relatório de vendas
report = "Total Sales: $" + str(total_sales) + "\n"
report += "Average Sales: $" + str(average_sales) + "\n"
report += "Maximum Sales: $" + str(max_sales) + "\n"
report += "Minimum Sales: $" + str(min_sales)

#Enviando e-mail
outlook = win32.Dispatch("Outlook.Application")
mail = outlook.CreateItem(0)
mail.To = "recipient@email.com"
mail.subject = "Daily Sales report"
mail.Body = report
mail.Attachments.Add("sales.csv")
mail.Send()

import pandas as pd
from datetime import  datetime, timezone
from pathlib import Path
import requests
import schedule
import time

def backup_and_upload():
    #Carregando os dados de um arquivo excel no Pandas dataframe
    file_path = "sales_data.xlsx"
    df = pd.read_excel(file_path)

    #Salvar backup dos daods em uma pasta local
    backup_folder = Path("Backups")
    backup_folder.mkdir(exist_ok=True)
    backup_file_path = backup_folder /f"{date-time.now().strftime('%Y-%m-%d_%H-%m-%s')}.xlsx"
    shutil.copy(file_path, backup_file_path)

    #Upload do arquivo de backup em uma pasta sharepoint
    sharepoint_url = "https://<company_name>.sharepoint.com/sites/<site_name>/Shared%20Documents/backups"
    file_name = backup_file_path.name
    with open(backup_file_path, "rb") as f:
        requests.put(f"{sharepoint_url}/{file_name}", data=f)
    
    #Agendar o backup para rodar em uma data e horários específicos
    schedule.every().monday.at("09:00").do(backup_and_upload)
    schedule.every().wednesday.at("09:00").do(backup_and_upload)
    schedule.every().friday.at("09:00").do(backup_and_upload)

    while True:
        schedule.run_pending()
        time.sleep(1)
