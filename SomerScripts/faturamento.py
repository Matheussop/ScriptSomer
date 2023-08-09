from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.cell import Cell
from openpyxl.worksheet.worksheet import Worksheet

ROOT_FOLDER = Path(__file__).parent
WORKBOOK_PATH_COMPANY = ROOT_FOLDER / 'empresas.xlsx'
WORKBOOK_PATH_FUNC = ROOT_FOLDER / 'funcionarios.xlsx'

# Carregando um arquivo do excel
workbook_company: Workbook = load_workbook(WORKBOOK_PATH_COMPANY)
workbook_employees: Workbook = load_workbook(WORKBOOK_PATH_FUNC)

companyList = workbook_company.sheetnames
sheet_name_employees = 'Test Sheet'
for company in companyList:  # Company is a sheet name
    """
    A partir daqui eu ja peguei a empresa que eu qro trabalhar.
    """
    # Selecionou a planilha da empresa'
    worksheet_company: Worksheet = workbook_company[company]
    worksheet_employees: Worksheet = workbook_employees[sheet_name_employees]

    # exam_name = worksheet_employees.iter_rows(min_row=0)
    list_of_company = []

    for exam in worksheet_company.iter_rows(min_row=1, values_only=True):
        list_of_company.append(exam)
    print("=================")
    print('Empresas -> ', list_of_company)
    print("=================")

    name_exam = []
    index = 0
    employees_cost = []

    # row: tuple[Cell]
    for row in worksheet_employees.iter_rows(min_row=1, values_only=True):
        if (row[0] != 'Nome'):
            name_exam.append((row[0], row[1]))
    print("NAME_EXAM -> ", name_exam)
    for name in name_exam:
        exams_exect = name[1].split('/')
        employees_total_cost = (0, '')
        hasExam = False
        for exam in exams_exect:
            hasExam = False
            for company in list_of_company:
                if (company[0] != 'Exame' and company[0] == exam):
                    employees_total_cost = (
                        employees_total_cost[0] + company[1], '')
                    hasExam = True
            if (not hasExam):
                if (employees_total_cost[1] == ''):
                    employees_total_cost = (
                        employees_total_cost[0], "ah empresa não tem o(s) exame(s)")
                employees_total_cost = (
                    employees_total_cost[0],  employees_total_cost[1] + " " + exam)

        employees_cost.append(
            (name[0], employees_total_cost[0], employees_total_cost[1]))
        # for i, cell in enumerate(row):
        # index = index+1
        # if (row[i].value != 'Exames' and type(row[i].value) == 'str'):
        # print(cell.value, end='\t')
        # print(list_of_company[i])
        # print()

    print("==================")
    print("EMPLOYEES_COST -> ", employees_cost)
workbook_company.save(WORKBOOK_PATH_COMPANY)
workbook_employees.save(WORKBOOK_PATH_FUNC)
