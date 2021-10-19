from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from copy import copy

planilha_preco_name = 'F4792-001-00 - PC - ANEXO 01 - Planilha de Preço.xlsx'
mc_name = 'F4792-001-00 - MC - PIAUÍ NÍQUEL - Bay de Saída 230 kV.xlsm'

wb_planilha_preco = load_workbook(planilha_preco_name)
wb_mc = load_workbook(mc_name, data_only=True)

planilha_preco = wb_planilha_preco['Teste']
memo = wb_mc['Memo Geral']
db = wb_mc['DashBoard']

  
def get_se_names():
  names_se = []
  for row in range(1, memo.max_row):
    value = memo[f'Z{row}'].internal_value
    if value != None and value[0] == 'S' and value[1] == 'E' and value not in names_se:
      names_se.append(value)
  print(names_se)
  return names_se

def get_total_row():
  for row in range(1, planilha_preco.max_row + 1):
    cell = f'A{row}'
    if planilha_preco[cell].value == 'TOTAL':
      return row

def make_titles(names_se):
  for se in names_se:
    if names_se.index(se) == 0: #preencher caso for a primeira SE
      planilha_preco.cell(row=4, column=1,value=str(se)) # Nome do título
      planilha_preco.cell(row=8, column=2, value=str(se)) # Nome da primeira SE
    
    else: #preenche o restante das SEs
      total_row = get_total_row()
      planilha_preco.insert_rows(total_row, 6)
      
      for row in range(total_row, total_row + 6):
        for column in range(1, planilha_preco.max_column + 1):
        
          copy_cell = planilha_preco[f'{get_column_letter(column)}{row-6}']
          new_cell = planilha_preco.cell(row=row, column=column, value="")
          new_cell._style = copy(copy_cell._style)

          if row == total_row and column == 2: # preenche o nome da SE
            planilha_preco.cell(row=row, column=column, value=str(se))

          elif column == 2:
            new_cell = planilha_preco.cell(row=row, column=column, value=copy_cell.value) #preenche a descrição

          if row == total_row and column == 1:
            planilha_preco.cell(row=row, column=column, value=int(names_se.index(se) + 1)) #preenche o primeiro item
          
          if row >= total_row + 1 and column == 1:
            planilha_preco.cell(row=row, column=column, value=str(f'{names_se.index(se) + 1}.{row-total_row}')) #preenche os restantes dos itens
          
       
          


names_se = get_se_names()
make_titles(names_se)
wb_planilha_preco.save('Nova.xlsx')