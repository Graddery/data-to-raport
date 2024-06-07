from tkinter import *  
from tkinter.ttk import Radiobutton  
from tkinter import ttk
from tkinter import filedialog
import os
import pandas as pd
import openpyxl
from tkinter import messagebox 

global file_name
global raport_name

def find_csv():
    global file_name
    file_name = filedialog.askopenfilename() 

def find_raport():
    global raport_name
    raport_name = filedialog.askopenfilename() 

def create_raport():
    global file_name
    global raport_name

    df_new = pd.read_csv(file_name)
    base = os.path.splitext(file_name)[0]
    new_filepath = base + '.xlsx'
     
    GFG = pd.ExcelWriter(new_filepath)
    df_new.to_excel(GFG, index=False)
     
    GFG._save() 

    wb1 = openpyxl.load_workbook(new_filepath)
    sheet1 = wb1.active

    wb2 = openpyxl.load_workbook(raport_name)
    sheet2 = wb2.active

    #Вставка значений - Дата проверки
    b = 3
    dR = 15
    for i in range(19): #19
        if (dR >= 15 and dR <= 19) or dR == 36:
            value = 'ТА: ' + str(sheet1['B' + str(b)].value)
            sheet2['D' + str(dR)] = value
            
        elif dR == 21:
            value = str(sheet1['B' + str(b)].value)
            array = value.split()
            first_data = array[0]
            second_data = array[1]
            sheet2['D' + str(dR)] = 'осн.: ' + first_data + '\n рез.: ' + second_data
            dR = dR-1

        elif dR == 22 or dR == 24 or dR == 28 or dR == 38:
            value = str(sheet1['B' + str(b)].value)
            array = value.split()
            first_data = array[0]
            second_data = array[1]
            sheet2['D' + str(dR)] = 'р/ст: ' + first_data + '\n пульт: ' + second_data

        elif dR == 26:
            value = 'Пульт: ' + str(sheet1['B' + str(b)].value)
            sheet2['D' + str(dR)] = value

        elif dR == 30 or dR == 34 or dR == 40 or dR == 44 or dR == 45:
            value = 'X'
            sheet2['D' + str(dR)] = value
            if dR == 40:
                dR = dR-1
            else:
                dR=dR

        elif dR == 32:
            value = 'ДПС: ' + str(sheet1['B' + str(b)].value)
            sheet2['D' + str(dR)] = value

        elif dR == 41 or dR == 43:
            value = 'РИ: ' + str(sheet1['B' + str(b)].value)
            sheet2['D' + str(dR)] = value
            if dR == 43:
                dR = dR-1
            else:
                dR=dR
        else:
            value = sheet1['B' + str(b)].value
            sheet2['D' + str(dR)] = value

        if dR >= 44:
            dR = dR+1
        else:
            dR = dR+2
        b = b+1

    #Вставка значений - Количество проверяемых устройств
    c = 2
    eR = 13
    for i in range(20): #20
        if eR == 22 or eR == 24 or eR == 40 or eR == 43:
            value = 'X'
            sheet2['F' + str(eR)] = value
            if eR == 40 or eR == 43:
                eR = eR-1
            else:
                eR=eR
        else:
            if eR == 21:
                value = sheet1['C' + str(c)].value
                sheet2['E' + str(eR)] = value
                eR = eR-1
                
            elif eR == 41:
                value = 'Количество проверяемых каналов \n' + sheet1['C' + str(c)].value
                sheet2['E' + str(eR)] = value
                
            elif eR == 44 or eR == 46:
                value = 'Количество ознакомленных \n' + sheet1['C' + str(c)].value
                sheet2['E' + str(eR)] = value
                eR = eR-1
                
            elif eR == 45:
                value = str(sheet1['C' + str(c)].value)
                array = value.split()
                first_data = array[0]
                second_data = array[1]
                sheet2['E' + str(eR)] = 'Количество пломб: ' + first_data + '\n \n Количество стикеров: ' + second_data
                eR = eR-1

            else:
                value = sheet1['C' + str(c)].value
                sheet2['E' + str(eR)] = value
            
        eR = eR+2
        c = c+1

    #Вставка значений - С кем проверена свзяь
    d = 2
    fR = 13 
    for i in range(20): #20
        if fR == 21 or fR == 34 or fR == 40 or fR > 42:
            value = 'X'
            sheet2['F' + str(fR)] = value

            if fR == 34:
                fR=fR
            else: 
                fR = fR-1
        else:
            value = sheet1['D' + str(d)].value
            sheet2['F' + str(fR)] = value
            
        fR = fR+2
        d = d+1

    #Вставка значений - Время проверки
    e = 2
    gR = 13
    for i in range(20): #20
        value = sheet1['E' + str(e)].value
        sheet2['G' + str(gR)] = value
        
        if gR == 21 or gR == 40 or gR > 42:
            gR = gR+1 
        else:
            gR = gR+2    
        e = e+1

    #Вставка значений - Замечания
    f = 2
    hR = 13
    for i in range(20): #20
        value = sheet1['F' + str(f)].value
        sheet2['H' + str(hR)] = value
        
        if hR == 21 or hR == 40 or hR > 42:
            hR = hR+1 
        else:
            hR = hR+2    
        f = f+1

    #Вставка значений - Меры
    g = 2
    iR = 13
    for k in range(20): #20
        value = sheet1['G' + str(g)].value
        sheet2['I' + str(iR)] = value

        if iR == 21 or iR == 40 or iR > 42:
            iR = iR+1 
        else:
            iR = iR+2    
        g = g+1

    #Вставка общих сведений
    j = 2
    h = 2
    value = sheet1['J' + str(j)].value
    sheet2['H' + str(h)] = value

    o = 2
    h = 6
    value = sheet1['J' + str(j)].value
    sheet2['H' + str(h)] = value
        
    k = 2
    h = 7
    value = 'от ' + sheet1['K' + str(k)].value
    sheet2['H' + str(h)] = value
        
    i = 2
    a = 10
    value = 'проведения комиссионного месячного осмотра по станции ' + sheet1['I' + str(i)].value
    sheet2['A' + str(a)] = value

    l = 2
    c = 50
    value = sheet1['L' + str(l)].value
    sheet2['C' + str(c)] = value

    m = 2
    e = 50
    value = sheet1['M' + str(m)].value
    sheet2['E' + str(e)] = value

    n = 2
    i = 50
    value = sheet1['N' + str(n)].value
    sheet2['I' + str(i)] = value


    wb2.save(raport_name)
    messagebox.showinfo('Рапорт КМО', 'Готово.\n Рапорт заполнен!')

window = Tk()  
window.title("Рапорт КМО")  
window.geometry('530x120')  
selected = IntVar()  
 
btn = Button(window, text="Выберите CSV файл", command = find_csv)
btn2 = Button(window, text="Выберите файл рапорта", command = find_raport)
btn3 = Button(window, text="Начать заполнение рапорта", command = create_raport)
btn.grid(column=1, row=1, pady=40, padx=15)
btn2.grid(column=2, row=1, pady=40,padx=15)
btn3.grid(column=3, row=1, pady=40,padx=15)
path_icon = os.path.dirname(__file__) + '\icon.ico'
window.iconbitmap(path_icon)

window.mainloop()




