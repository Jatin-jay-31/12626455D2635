import openpyxl
import math

def TPI(filename,sheet_name,column):
    wb=openpyxl.load_workbook(filename,data_only=True)
    ws=wb[sheet_name]
    values=[]
    for row in range(5,ws.max_row+1):
        value=ws.cell(row,column).value
        if value is not None:
            values.append(value)
        print(values)
    return math.sum(values)/len(values)
col_tpi= TPI("12626455.xlxs","Daily Log",5)

def AAI(filename,sheet_name,column1,column2):
    wb=openpyxl.load_workbook(filename,data_only=True)
    ws=wb[sheet_name]
    values1=[]
    values2=[]
    for row in range(5,ws.max_row+1):
        value1=ws.cell(row,column1).value
        value2=ws.cell(row,column2).value
        if value1 is not None:
            values1.append(value1)
        if value2 is not None:
            values2.append(value2)
        print(values1)
    return math.sum(values1)+math.sum(values2)/len(values1)
col_aai= AAI("12626455.xlxs","Daily Log",4,6)

def PhAI(filename,sheet_name,column):
    wb=openpyxl.load_workbook(filename,data_only=True)
    ws=wb[sheet_name]
    values=[]
    for row in range(5,ws.max_row+1):
        value=ws.cell(row,column).value
        if value is not None:
            values.append(value)
        print(values)
    return math.sum(values)/len(values)
col_phai= PhAI("12626455.xlxs","Daily Log",3)

def SRI(filename,sheet_name,column):
    wb=openpyxl.load_workbook(filename,data_only=True)
    ws=wb[sheet_name]
    values=[]
    for row in range(5,ws.max_row+1):
        value=ws.cell(row,column).value
        if value is not None:
            values.append(value)
        print(values)
    return math.sum(values)/len(values)
col_sri= SRI("12626455.xlxs","Daily Log",2)

def ABI(filename,sheet_name,column):
    wb=openpyxl.load_workbook(filename,data_only=True)
    ws=wb[sheet_name]
    values=[]
    for row in range(5,ws.max_row+1):
        value=ws.cell(row,column).value
        if value is not None:
            values.append(value)
        print(values)
    return math.sum(values)/len(values)
col_abi= ABI("12626455.xlxs","Daily Log",10)

def TUI(filename,sheet_name,column):
    wb=openpyxl.load_workbook(filename,data_only=True)
    ws=wb[sheet_name]
    values=[]
    for row in range(5,ws.max_row+1):
        value=ws.cell(row,column).value
        if value is not None:
            values.append(value)
        print(values)
    return math.sum(values)/len(values)
col_tui= TUI("12626455.xlxs","Daily Log",9)

    

def EI(filename,sheet_name,col1,col2,col3):
    wb=openpyxl.load_workbook(filename,data_only=True)
    ws=wb[sheet_name]
    day_feeling={1:'Stressed',2:'Low',3:'Neutral',4:'Good',5:'Excellent'}
    energy_level={1:'Low',2:'Medium',3:"High"}
    sats_level={1:'Very Unsatisfied',2:'Unsatisfied',3:'Neutral',4:'Satisfied',5:'Very Satisfied'}
    lst1=[]
    lst2=[]
    lst3=[]
    for row in range(5,ws.max_row+1):
        value1=ws.cell(row,lst1).value
        value2=ws.cell(row,lst2).value
        value3=ws.cell(row,lst3).value
        if value1 in day_feeling:
            lst1.append(day_feeling[value1])
        if value2 in day_feeling:
            lst2.append(day_feeling[value2])
        if value3 in day_feeling:
            lst3.append(day_feeling[value3])
    return (math.sum(lst1)+ math.sum(lst1)+math.sum(lst1))/(3 *len(lst1))
col_ei= EI("12626455.xlxs","Daily Log",11,12,13)

def DCI(filename,sheet_name,column):
    wb=openpyxl.load_workbook(filename,data_only=True)
    ws=wb[sheet_name]
    values=[]
    for row in range(5,ws.max_row+1):
        value=ws.cell(row,column).value
        if value is not None:
            values.append(value)
        print(values)
    return math.sum(values)/len(values)
col_dci= DCI("12626455.xlxs","Daily Log",2)

def PAI(a,b,c,d,e,f,g):
    pai_value=(0.15*a + 0.20*b + 0.15*c + 0.20*d + 0.15*e+ 0.10*f + 0.05*g)
    return pai_value
result= PAI(col_tpi,col_aai,col_phai,col_sri,col_abi,col_ei,col_dci)
    