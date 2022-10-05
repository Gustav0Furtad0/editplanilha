import openpyxl as opx;
#import pyexcel as opexcel;

arquivoFinal = opx.load_workbook(filename='itensBranetJF.xlsx')
planilhaFinal = arquivoFinal['MEDICAMENTOS']

arquivoAuxiliar = input('Digite o nome da planilha a ser inserida com seu formato')
arquivoAuxiliar = opx.load_workbook(filename=arquivoAuxiliar)
planilhaAuxiliar = arquivoAuxiliar["CMM UBS's"]

i = 3
while planilhaAuxiliar.cell(row=i, column=2).value != None:
    val = [i, planilhaAuxiliar.cell(row=i, column=1).value, planilhaAuxiliar.cell(row=i, column=2).value]
    print(val)
    i += 1

# arquivofinal.save(filename='relatorios.xlsx')