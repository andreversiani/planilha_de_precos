from openpyxl import Workbook, load_workbook

wb_planilha_preco = load_workbook('F4792-001-00 - PC - ANEXO 01 - Planilha de Preço - Copia.xlsx')
wb_mc = load_workbook('F4792-001-00 - MC - PIAUÍ NÍQUEL - Bay de Saída 230 kV.xlsm', data_only=True)

planilha_preco = wb_planilha_preco['Teste']
memo = wb_mc['Memo Geral']
db = wb_mc['DashBoard']

def get_se_names():
  names_se = []
  for row in range(1, memo.max_row):
    value = memo[f'Z{row}'].internal_value
    if value != None and value[0] == 'S' and value[1] == 'E' and value not in names_se:
      names_se.append(value)

  return names_se

names_se = get_se_names()

print(names_se)