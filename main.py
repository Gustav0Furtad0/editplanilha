import openpyxl as opx;
import pyexcel as opexcel;

arquivoFinal = opx.load_workbook(filename='intensBranetJF.xlsx');
planilhaFinal = arquivoFinal['MEDICAMENTOS'];

planilhaAuxiliar = input('Digite o nome da planilha a ser inserida')
planilhaAuxiliar = opx.load_workbook(filename='planilha'

for i in range(3, 464):
    val = planilhaFinal.cell(row=i, column=2).value.split()
    print(val)

# arquivofinal.save(filename='relatorios.xlsx')