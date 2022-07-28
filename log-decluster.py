import pandas as pd
import json
import base64
import ctypes
import tkinter as tk
import os
from tkinter import filedialog

root = tk.Tk()
root.withdraw()

path = filedialog.askopenfilename()
fixed_path = "fixed_msg.json"

with open(path, 'r') as data_file:
  data = data_file.read()

data = data[:-1]                                     ## Remove o último '\n', adiciona [] ao início e ao fim do arquivo, 
data = "[" + data.replace("}\n", "},\n") + "]"       ## e vírgulas em todas as linhas, com exceção da última.

with open(fixed_path, "w") as fixed_file:            ## Salva o JSON em um novo arquivo.
  for line in data:
    fixed_file.write(line)

with open(fixed_path, 'r') as fixed_file:            ## Abre o novo arquivo para leitura.
  json_data = json.loads(fixed_file.read())

df = pd.json_normalize(json_data)                    ## Normaliza e insere o JSON no dataframe.

df = df[["type",                                     ## Seleção das colunas de variáveis relevantes.
         "params.counter_down", 
         "params.counter_up", 
         "meta.time",
         "params.port",
         "params.payload", 
         "params.radio.hardware.power", 
         "params.radio.datarate", 
         "params.lora.header.confirmed",
         "params.lora.header.adr",
         "params.lora.header.adr_ack_req"]]

df = df.loc[df['type'].isin(['downlink', 'uplink'])] ## Apenas downlinks e uplinks.

## Títulos das colunas
dt = "Type"
ts = "Timestamp"
ctup = "Counter up"
ctdwn = "Counter down"
pl = "Payload"
tx = "TXpower"
dr = "DR"
cfd = "Confirmed"
adr = "ADR mode"
req = "adr_ack_req"
port = "Port"

df = df.rename(columns={"type":dt,
                        "meta.time":ts,
                        "params.counter_up":ctup,
                        "params.counter_down":ctdwn,  
                        "params.payload":pl, 
                        "params.radio.hardware.power":tx, 
                        "params.radio.datarate":dr,
                        "params.lora.header.confirmed":cfd,
                        "params.lora.header.adr":adr,
                        "params.lora.header.adr_ack_req":req,
                        "params.port":port})

df[ts] = pd.to_datetime(df[ts], unit='s')
df[pl] = df[pl].apply(lambda b:base64.b64decode(b).hex())
df[port] = df[port].astype('Int64')
df[dr] = df[dr].astype('Int64')
df[ctup] = df[ctup].astype('Int64')
df[ctdwn] = df[ctdwn].astype('Int64')

## Configuração do arquivo .xlsx
writer = pd.ExcelWriter(fixed_path+".xlsx", engine='xlsxwriter')

df.to_excel(writer, sheet_name='Sheet1', startrow=1, header=False, index=False)

workbook = writer.book
worksheet = writer.sheets['Sheet1']

(max_row, max_col) = df.shape
column_settings = []
for header in df.columns:
    column_settings.append({'header': header})

worksheet.add_table(0, 0, max_row, max_col - 1, {'columns': column_settings})
worksheet.set_column(0, max_col - 1, 12)
writer.save()

MessageBox = ctypes.windll.user32.MessageBoxW
MessageBox(None, 'Processo finalizado.\nArquivos .json e .xlsx salvos em:\n\n'+os.getcwd(), 'Sucesso', 0)