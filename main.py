from openpyxl import load_workbook
wb = load_workbook(filename = 'prueba.xlsx')
sheet_ranges = wb['Hoja1']
columna= ("A")
numColumna= 1
evaColumna= ("A")
evanumColumna= 2
evanumColumna1= 2
for i in sheet_ranges:
    value=(sheet_ranges[columna+str(numColumna)].value)
    for j in sheet_ranges:
        value1=(sheet_ranges[evaColumna+str(evanumColumna)].value)
        evanumColumna += 1
        if value1==value:
            print ("Repetido"+" "+str(value))
        if value1==None:
            evanumColumna= numColumna+2
            break
    numColumna += 1

        
print("Finished")



#D:\OneDrive\Escritorio