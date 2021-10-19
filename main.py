from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from copy import copy

planilha_preco_name = 'F4792-001-00 - PC - ANEXO 01 - Planilha de Preço.xlsx'
mc_name = 'F4792-001-00 - MC - PIAUÍ NÍQUEL - Bay de Saída 230 kV.xlsm'

wb_planilha_preco = load_workbook(planilha_preco_name)
wb_mc = load_workbook(mc_name, data_only=True)

planilha_preco = wb_planilha_preco['Teste']
styles = wb_planilha_preco['Styles']
memo = wb_mc['Memo Geral']
db = wb_mc['DashBoard']

conferencia_column = 'X'
descricao_column = 'B'
subestacao_column = 'Z'
qte_column = 'D'
preco_impostos_column = 'T'

pis_confins_eq = '$M$14'
pis_confins_sv = '$M$13'
icms = '$M$22'
iss_bh = '$M$15'
iss_cliente = '$M$16'

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
          

def make_engenharia():
  
  for planilha_preco_row in range(1, 1000):
    if planilha_preco[f'B{planilha_preco_row}'].value == 'ENGENHARIA':
      se = str(planilha_preco[f'B{planilha_preco_row - 1}'].value)
      i = 1
      for row in range(1, memo.max_row + 1):
        if memo[f'{conferencia_column}{row}'].value == "Projetos" and memo[f'{subestacao_column}{row}'].value == se and int(memo[f'{qte_column}{row}'].value) >= 1:
          
          planilha_preco.insert_rows(planilha_preco_row + i, 1)
          
          #Somas
          planilha_preco.cell(row=planilha_preco_row, column=16, value=f"=SUM(P{planilha_preco_row+1}:P{planilha_preco_row+i})")
          planilha_preco.cell(row=planilha_preco_row, column=15, value=f"=SUM(O{planilha_preco_row+1}:O{planilha_preco_row+i})")
          planilha_preco.cell(row=planilha_preco_row, column=14, value=f"=SUM(N{planilha_preco_row+1}:N{planilha_preco_row+i})")
          planilha_preco.cell(row=planilha_preco_row, column=12, value=f"=SUM(L{planilha_preco_row+1}:L{planilha_preco_row+i})")
          planilha_preco.cell(row=planilha_preco_row, column=11, value=f"=SUM(K{planilha_preco_row+1}:K{planilha_preco_row+i})")
          planilha_preco.cell(row=planilha_preco_row, column=9, value=f"=SUM(I{planilha_preco_row+1}:I{planilha_preco_row+i})")
          planilha_preco.cell(row=planilha_preco_row, column=7, value=f"=SUM(G{planilha_preco_row+1}:G{planilha_preco_row+i})")
          planilha_preco.cell(row=planilha_preco_row, column=5, value=f"=SUM(E{planilha_preco_row+1}:E{planilha_preco_row+i})")
          
          #Campos brancos
          planilha_preco.cell(row=planilha_preco_row + i, column=2, value=f"='[{mc_name}]Memo Geral'!${descricao_column}${row}")
          planilha_preco.cell(row=planilha_preco_row + i, column=3, value=f"='[{mc_name}]Memo Geral'!${qte_column}${row}")
          planilha_preco.cell(row=planilha_preco_row + i, column=4, value="R$")
          planilha_preco.cell(row=planilha_preco_row + i, column=16, value=f"='[{mc_name}]Memo Geral'!{preco_impostos_column}${row}")
          
          #formulas fixadas nos campos brancos
          planilha_preco.cell(row=planilha_preco_row + i, column=6, value=f"=$F${planilha_preco_row+1}")
          planilha_preco.cell(row=planilha_preco_row + i, column=8, value=f"=$H${planilha_preco_row+1}")
          planilha_preco.cell(row=planilha_preco_row + i, column=10, value=f"=$J${planilha_preco_row+1}")
          planilha_preco.cell(row=planilha_preco_row + i, column=13, value=f"=$M${planilha_preco_row+1}")
          
          #estilos
          for column in range(1, planilha_preco.max_column + 1):
            cell = planilha_preco[f'{get_column_letter(column)}{planilha_preco_row + i}']
            cell._style = copy(styles[f'{get_column_letter(column)}3']._style)
          i += 1
  
  #impostos
  for planilha_preco_row in range(1, 1000):
    if planilha_preco[f'B{planilha_preco_row}'].value == 'ENGENHARIA':
      planilha_preco.cell(row=planilha_preco_row + 1, column=6, value=f"='[{mc_name}]DashBoard'!{pis_confins_eq}")
      planilha_preco.cell(row=planilha_preco_row + 1, column=8, value=0)
      planilha_preco.cell(row=planilha_preco_row + 1, column=10, value=f"='[{mc_name}]DashBoard'!{iss_bh}")
      planilha_preco.cell(row=planilha_preco_row + 1, column=13, value=0)



names_se = get_se_names()
make_titles(names_se)
make_engenharia()
wb_planilha_preco.save('Nova.xlsx')
