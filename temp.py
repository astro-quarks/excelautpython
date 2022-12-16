# -*- coding: utf-8 -*-

import openpyxl as xl
import bot

#obtenemos todos los datos 3.1 de la sheet cultivos de los 34
print(bot.content)


casillas = ['B']
hoja = 6
f =1
fila =  '15'
archivo = xl.load_workbook('../../au2.xlsx')
sheet = archivo['Hoja1']
for file in bot.content:
    df = xl.load_workbook(bot.pth + file)
    st = df[bot.stn[hoja]]

    for casilla in casillas:
        valor = st[casilla + fila].value

        sheet[casilla + str(f)] = valor
        print("Obteniendo informacion " + str(f) + "/34")

        archivo.save(filename='../../au2.xlsx')
    f += 1


    df.close()









