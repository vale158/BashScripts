#!/usr/bin/env python
# coding: utf-8

# In[8]:


import win32com.client
from datetime import datetime

# Abre el libro de Excel
excel = win32com.client.Dispatch("Excel.Application")
workbook1 = excel.Workbooks.Open(r"e:/users/vvaldesg/Documents/EnvioAutomatico/CobranzaPorDia_2023_06.xlsx")
workbook2 = excel.Workbooks.Open(r"e:/users/vvaldesg/Documents/EnvioAutomatico/RankingCDTSem23_06.xlsx")
workbook3 = excel.Workbooks.Open(r"e:/users/vvaldesg/Documents/EnvioAutomatico/RankingRel2023_06.xlsx")
workbook4 = excel.Workbooks.Open(r"e:/users/vvaldesg/Documents/EnvioAutomatico/RankingVenc2023_06.xlsx")
#workbook5 = excel.Workbooks.Open(r"e:/users/vvaldesg/Documents/EnvioAutomatico/REKTContencion23Sem04.xlsx")


archivos=[workbook1,workbook2,workbook3,workbook4]
for x in archivos:
# Actualiza los datos de la conexión
    x.RefreshAll()
# Guarda los cambios en el libro
    x.Save()
# Cierra el libro y Excel
    x.Close()

excel.Quit()


fecha = datetime.now()
formato = fecha.strftime("%Y Sem %W, %d %b Corte %I %p")
outlook = win32com.client.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'geografiametrosur@onuriscp.com'
#mail.To = 'valentin.valdes@elektra.com.mx'

mail.Subject = 'Seguimiento diario ' + formato
mail.Body = 'Buen día. Comparto archivos de seguimiento al corte. \n Saludos.\n Valentín Valdés \n 5585537967\n'

mail.Attachments.Add('e:/users/vvaldesg/Documents/EnvioAutomatico/CobranzaPorDia_2023_06.xlsx')
mail.Attachments.Add('e:/users/vvaldesg/Documents/EnvioAutomatico/RankingCDTSem23_06.xlsx')
mail.Attachments.Add('e:/users/vvaldesg/Documents/EnvioAutomatico/RankingRel2023_06.xlsx')
mail.Attachments.Add('e:/users/vvaldesg/Documents/EnvioAutomatico/RankingVenc2023_06.xlsx')
#mail.Attachments.Add('e:/users/vvaldesg/Documents/EnvioAutomatico/REKTContencion23Sem04.xlsx')
mail.Send()



