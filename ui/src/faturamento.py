import pandas as pd
import math
import json
import locale
import sys
import asyncio
import os

from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.worksheet import Worksheet
from typing import List
from datetime import datetime
from openpyxl.styles import Font, PatternFill, NamedStyle, Alignment, Border
from openpyxl.styles import Side
from unidecode import unidecode


# dictionary_exams = [
#     ("CLINICO", "Clínico"),   ("AUDIO", "Audiometria"),
#     ("HEMO", "Hemograma c/ Plaquetas"), ("AC METIL", "Ácido Metil Hipúrico"),
#     ("AC HIPURICO", "Ácido Hipúrico"), ("AC MANDELICO", "Ácido Mandélico"),
#     ("AC METIL HIPURICO", "Ácido Metil Hipúrico"),
#     ("RX DE TORAX", "Raio-x de Tórax"),
#     ("AV PSICOSSOCIAL", "Avaliação Psicossocial"),
#     ("AV PSICOLÓGICA", "Avaliação Psicológica"),
#     ("CLINICO EXTERNO", "Exame Clínico (in loco)"),
#     ("BIOMETRIA IMC", "Biometria"),
#     ("COLESTEROL LDL", "LDL Colesterol"),
#     ("LDL", "LDL Colesterol"),
#     ("COLESTEROL HDL", "HDL Colesterol"),
#     ("HDL", "HDL Colesterol"),
#     ("COLESTEROL VLDL", "VLDL Colesterol"),
#     ("VLDL", "VLDL Colesterol"),
#     ("AC Úrico", "Acido Úrico"),
#     ("RX DE COLUNA DORSAL", "Raio-x de Coluna Dorsal"),
# ]

dictionary_company_names = [
    ("Minas Brasil", "BRASIL COMERCIAL"),
    ("ACTur", "AC TRANSPORTES"),
]


class Employee:
    cost: float
    name: str
    exams: str
    function: str
    date: str
    examsCost: List
    typeExam: str

    def __init__(self, examsCost: List = [], name: str = '', cost: float = 0.0,
                 function: str = '', exams: str = '', date: str = '',
                 typeExam: str = '',):
        self.name = name
        self.cost = cost
        self.function = function
        self.exams = exams
        self.date = date
        self.typeExam = typeExam
        self.examsCost = examsCost


class Company:
    name: str
    exams: str
    missingExams: str
    employees: List[Employee]

    def __init__(self, name: str = '',
                 exams: str = '', employees: List[Employee] = [],
                 missingExams: str = ''):
        self.name = name
        self.missingExams = missingExams
        self.exams = exams
        self.employees = employees

    def __repr__(self):
        return f'Nome empresa: {self.name}, \n Lista de Exames: {self.exams},'\
            f'\n Exames Faltando: {self.missingExams})'

    def employeesList(self):
        employeeString = []
        for employee_person in self.employees:
            formatedCoast = math.ceil(employee_person.cost * 100)/100
            employeeString.append(f'Nome: {employee_person.name}, '
                                  f'Custo: {formatedCoast}')
        return employeeString


class EmployeesBilling:
    name: str
    function: str
    data: datetime
    exams: str
    typeExam: str

    def __init__(self, name: str = '', function: str = '',
                 exams: str = '',
                 data: datetime = datetime.now(), typeExam: str = ''):
        self.name = name
        self.function = function
        self.exams = exams
        self.data = data
        self.typeExam = typeExam


class CompanyBilling:
    name: str
    employeesBilling: List[EmployeesBilling]

    def __init__(self, name: str = '', employeesBilling=[]):
        self.name = name
        self.employeesBilling = employeesBilling


class Billing:
    dataExam: str
    nameCompany: str
    nameEmployees: str
    functionEmployees: str
    exams: str
    typeExam: str

    def __init__(self, dataExam: str = '', nameCompany: str = '',
                 nameEmployees: str = '', functionEmployees: str = '',
                 exams: str = '', typeExam: str = ''):
        self.nameCompany = nameCompany
        self.nameEmployees = nameEmployees
        self.functionEmployees = functionEmployees
        self.exams = exams
        self.dataExam = dataExam
        self.typeExam = typeExam


# ROOT_FOLDER = Path(__file__).parent
# WORKBOOK_PATH_COMPANY = ROOT_FOLDER / 'ValoresEmpresas.xlsx'
# WORKBOOK_PATH_FUNC = ROOT_FOLDER / 'examesRealizados.xlsx'
WORKBOOK_PATH_COMPANY = './ValoresEmpresas.xlsx'
WORKBOOK_PATH_FUNC = './examesRealizados.xlsx'
hasDetailed = False


def getExams(company_name, list_of_company, newCompany):
    df = pd.read_excel(WORKBOOK_PATH_COMPANY, sheet_name=company_name)
    # Get row 3 to infinity and Column A and B
    tableOfCompany_Value = df.iloc[3:, [0, 1]]
    for exam in tableOfCompany_Value.values.tolist():
        if (str(exam[0]) == 'nan'):
            break
        if (exam[0] is not None):
            list_of_company.append((exam[0], exam[1]))
            if (newCompany.exams != ''):
                newCompany.exams = newCompany.exams + ", " + str(exam[0])
            else:
                newCompany.exams = str(exam[0])


def setDateXlsxToString(dataInXlsx) -> str:
    # Bloco para fixar a data como string
    data_str_fmt_for_date = '%Y-%m-%d %H:%M:%S'
    data_str_fmt = '%d/%m/%Y '
    date_Typed = datetime.strptime(dataInXlsx, data_str_fmt_for_date)
    return date_Typed.strftime(data_str_fmt)


def styleXlsx(wb):
    # Estilo da tabela
    bd = Side(style='thin', color="000000")
    style_header = NamedStyle(name="style_header")
    style_header.font = Font(size=12, color='FFFFFF', name='Arial', )
    style_header.font.bold = True
    style_header.alignment = Alignment(horizontal="center", vertical="center")
    style_header.fill = PatternFill("solid", fgColor="76933c")
    style_header.border = Border(left=bd, top=bd, right=bd, bottom=bd)

    style_content = NamedStyle(name="style_content")
    style_content.font = Font(size=10, name='Arial')
    style_content.fill = PatternFill("solid", fgColor="d8e4bc")
    style_content.border = Border(left=bd, top=bd, right=bd, bottom=bd)

    style_Cost = NamedStyle(name="style_Cost")
    style_Cost.font = Font(size=12, name='Arial', )
    style_Cost.font.bold = True
    style_Cost.alignment = Alignment(horizontal="center", vertical="center")
    style_Cost.fill = PatternFill("solid", fgColor="d8e4bc")
    style_Cost.border = Border(left=bd, top=bd, right=bd, bottom=bd)

    wb.add_named_style(style_header)
    wb.add_named_style(style_content)
    wb.add_named_style(style_Cost)


def setContentDataFrame(arrayUnformattedData: List[Employee]) -> List:
    data = []
    for employeesItem in arrayUnformattedData:
        aux = [employeesItem.date,
               employeesItem.name,
               employeesItem.function,
               employeesItem.exams,
               employeesItem.typeExam,
               employeesItem.cost]
        data.append(aux)
    return data


def setContentDetailDataFrame(arrayUnformattedData: List[Employee]) -> List:
    data = []
    for employeesItem in arrayUnformattedData:
        listOfExams = []
        listOfCost = []
        for examCost in employeesItem.examsCost:
            listOfExams.append(examCost[0])  # name exam
            listOfCost.append(examCost[1])  # value exam

        aux = [employeesItem.date,
               employeesItem.name,
               employeesItem.function,
               listOfExams,
               employeesItem.typeExam,
               listOfCost]
        data.append(aux)
    return data


def mountStyleHeaderTable(monthNumber: int, ws: Worksheet, companyName: str):
    columnsTable = ['A', 'B', 'C', 'D', 'E', 'F']
    # Cabeçalho da planilha
    ws.append([f'Mês {monthNumber:02}/23',
               companyName, '', '', '', 'Valor'])

    # Aplicando Header Styles
    for i in columnsTable:
        ws[f'{i}1'].style = 'style_header'

    # Unindo a células onde ficara o nome da empresa
    ws.merge_cells('B1:E1')


def mountStyleContentTable(ws):
    columnsTable = ['A', 'B', 'C', 'D', 'E', 'F']
    count_row = ws.max_row

    # Definindo o estilo e formatação do conteúdo da tabela
    for i in range(0, count_row-1):
        for y in columnsTable:
            ws[f'{y}{i+2}'].style = 'style_content'

            ws[f'{y}{i+2}'].alignment = Alignment(
                horizontal="center", vertical="center")
            if (y == 'F'):
                # Formula para forçar a ser um campo de dinheiro "BR"
                number_forma = '"$"#,##0.00_);[Red]("$"#,##0.00)'
                ws[f'{y}{i+2}'].number_format = number_forma
                ws[f'{y}{i+2}'].alignment = Alignment(
                    horizontal="right", vertical="center")
            if (y == 'D'):
                ws[f'{y}{i+2}'].alignment = Alignment(
                    horizontal="left", vertical="center")


def resizeTable(ws):
    # Redimensionar colunas da tabela
    dims = {}
    for row in ws.rows:
        for cell in row:
            if cell.value:
                dims[cell.column_letter] = max(
                    (dims.get(cell.column_letter, 0), len(str(cell.value))))
    for col, value in dims.items():
        ws.column_dimensions[col].width = value + 4


def calculateCost(company) -> int:
    cost = 0
    for employee in company.employees:
        cost += employee.cost
    return cost


def setTotalCostStyle(ws):
    columnsTable = ['A', 'B', 'C', 'D', 'E', 'F']
    row_count = ws.max_row
    # Formula para forçar a ser um campo de dinheiro "BR"
    number_forma = '"$"#,##0.00_);[Red]("$"#,##0.00)'
    for letter in columnsTable:
        ws[f'{letter}{row_count}'].style = 'style_Cost'

    ws[f'F{row_count}'].number_format = number_forma


def setTotalCostRow(ws):
    row_count = ws.max_row
    # Espaços em branco pois será onde a célula sofrera um merge
    ws.append(['TOTAL', '', '', '', '', f"=SUM(F2: F{row_count})"])
    ws.merge_cells(f'A{row_count+1}:E{row_count+1}')
    setTotalCostStyle(ws)


def addDetailDataFrame(ws, item):
    row_count = ws.max_row
    # Tamanho da lista de exames feitos pelo funcionário
    lenOfExams = len(item[3])
    exams = item[3]
    examsCost = item[5]
    # Caso não tenha nenhum exame limpa o campo.
    if (exams == []):
        exams.append("-")
        examsCost.append("-")

    firstRow = [item[0], item[1], item[2], exams[0], item[4], examsCost[0]]
    ws.append(firstRow)
    for i in range(1, lenOfExams):
        # Espaços em branco pois será onde a célula sofrera um merge
        ws.append(['', '', '',  exams[i], '', examsCost[i]])

    columnsMerge = ['A', 'B', 'C', 'E']
    # merge_cells A, B, C, E
    if (len(exams) >= 2):
        for letter in columnsMerge:
            ws.merge_cells(
                f'{letter}{row_count+1}:{letter}{row_count+lenOfExams}')


def getNameCompanyInSheet() -> List[tuple]:
    workbook_company: Workbook = load_workbook(WORKBOOK_PATH_COMPANY)

    # Percorrer o sheet e pegar o nome da empresa dentro da tabela
    names_of_workbook = workbook_company.sheetnames
    namesOfCompanys: List[tuple] = []
    for names in names_of_workbook:
        ws: Worksheet = workbook_company[f'{names}']
        namesOfCompanys.append((ws['B1'].value, names))

    for nameCompany in dictionary_company_names:
        namesOfCompanys.append((nameCompany[1], nameCompany[0]))

    return namesOfCompanys


def showMissingExams(ws, missingExams):
    text = 'Esta empresa não tem o(s) seguintes exame(s) cadastrados: '
    ws.append([f'{text}', '', f'{missingExams}',
              '', '', ''])
    ws.merge_cells(f'A{ws.max_row}:B{ws.max_row}')
    ws.merge_cells(f'C{ws.max_row}:F{ws.max_row}')


def showCompanyInTerminal(companyList: List[Company]):
    for company in companyList:
        print("========================================================")
        print("COMPANY LIST: \n",  company)
        print("COMPANY LIST OF EMPLOYEES: \n", company.employeesList())
        print("========================================================")


def showInTerminalCompanyNameNumber(companyName, number):
    print('Numero: ', number,
          ' Nome empresa', companyName)


class Faturamento:
    yearText = ''
    monthText = ''
    folderText = ''
    hasDetailed = False
    dictionary_exams = []
    maxEmployees = 0
    progress = 0
    worksheet_employees: Worksheet
    workbook_employees: Workbook
    namesOfCompanys = List[tuple]
    companys_not_found = []
    companyList = []
    companyList_Billing = []

    def __init__(self, yearText='', monthText='', folderText='',
                 hasDetailed=False, dictionary_exams=[],
                 namesOfCompanys: List[tuple] = []) -> None:
        self.yearText = yearText
        self.monthText = monthText
        self.folderText = folderText
        self.hasDetailed = hasDetailed
        self.dictionary_exams = dictionary_exams
        self.namesOfCompanys = namesOfCompanys

    def setParamsBilling(self, yearText, monthText, folderText,
                         hasDetailed) -> None:
        self.yearText = yearText
        self.monthText = monthText
        self.folderText = folderText
        self.hasDetailed = hasDetailed

    def setEmployeesFile(self):
        # Carregando o arquivo de empresas do excel
        self.workbook_employees: Workbook = load_workbook(WORKBOOK_PATH_FUNC)

        name_month = self.monthText.upper()
        year = str(self.yearText)
        sheet_name_employees = f'{name_month} {year}'
        # Carregando o arquivo de exames realizados do excel
        self.worksheet_employees: Worksheet = \
            self.workbook_employees[sheet_name_employees]

    def getCompanysNotFound(self): return self.companys_not_found

    async def generatedBaseDataCompany(self):
        billingList = []
        for billing_row in \
            self.worksheet_employees.iter_rows(min_row=2,
                                               values_only=True):
            if (billing_row[0] is not None):
                billingList.append(Billing(
                    # (Data exame, Name company)
                    str(billing_row[0]), str(billing_row[1]),
                    # (Name employee, Function Employee)
                    str(billing_row[2]), str(billing_row[3]),
                    # (Exams, Type Exam)
                    str(billing_row[4]), str(billing_row[5])))

        for billing_row in billingList:
            self.checkCompanys(billing_row)

        self.maxEmployees = len(billingList)
        sys.stderr.write(f'maxEmployees {len(billingList)}')
        # Bloco para testar numero x de empresas
        # companyList_BillingTeste = []
        # for i in range(20):
        #     companyList_BillingTeste.append(companyList_Billing[i])

        # self.companyList_Billing = [
        #     self.companyList_Billing[0], self.companyList_Billing[1]]
        # ----------------------------------------------------

        self.namesOfCompanys = getNameCompanyInSheet()
        self.companys_not_found.append('\nEmpresas não encontradas: ')

    # ----------------------------------------------------------------

    def saveDictionaryExams(self):
        with open('dictionary_exams.json', 'w', encoding='utf8') as arquivo:
            json.dump(
                self.dictionary_exams,
                arquivo,
                ensure_ascii=False,
                indent=4,
            )

    def getDictionaryExams(self) -> List[tuple]:
        newDictionary = []
        with open('dictionary_exams.json', 'r', encoding='utf8') as arquivo:
            newDictionary = json.load(arquivo)
        self.dictionary_exams = newDictionary
        return newDictionary

    async def createSheet(self, company, monthText, monthNumber):
        wb: Workbook = Workbook()
        styleXlsx(wb)
        wb.remove(wb['Sheet'])
        dataCompany = []

        if (self.hasDetailed):
            dataCompany = setContentDetailDataFrame(company.employees)
        else:
            dataCompany = setContentDataFrame(company.employees)

        ws: Worksheet = wb.create_sheet(company.name[:15])

        mountStyleHeaderTable(monthNumber, ws, company.name)

        # cost = 0
        # cost = calculateCost(company)
        # print(cost)
        for item in dataCompany:
            self.progress += 1
            sys.stderr.write(f"Total complete: {self.progress}%\n")
            if (self.hasDetailed):
                addDetailDataFrame(ws, item)
            else:
                # Caso eu não tenha exames encontrados
                # para esses empregado
                if (item[5] == 0):
                    item[3] = '-'
                    item[5] = '-'
                ws.append(item)
        if (company.missingExams != ''):
            showMissingExams(ws, company.missingExams)
        mountStyleContentTable(ws)
        setTotalCostRow(ws)
        resizeTable(ws)

        nameFile = f'{company.name} - {monthNumber} {monthText}'
        folderAndName = f'{self.folderText}/Faturamento - {nameFile}'
        wb.save(f'{folderAndName} 2023.xlsx')

    def checkCompanys(self, billing_row):
        # if (i > 50):
        #     continue
        noHasCompany = True
        # Para cada empresa criar uma instancia dela com o emprego caso ja
        # exista adicionar apenas o empregado a ela.
        for company in self.companyList_Billing:
            newCompanyBilling = CompanyBilling()
            if (billing_row.nameCompany == company.name):
                employeesBillingAux = EmployeesBilling(
                    billing_row.nameEmployees,
                    billing_row.functionEmployees,
                    billing_row.exams,
                    data=billing_row.dataExam,
                    typeExam=billing_row.typeExam)
                company.employeesBilling.append(employeesBillingAux)
                noHasCompany = False
        if (noHasCompany):
            employeesBillingAux = EmployeesBilling(
                name=billing_row.nameEmployees,
                function=billing_row.functionEmployees,
                exams=billing_row.exams,
                data=billing_row.dataExam,
                typeExam=billing_row.typeExam)
            newCompanyBilling = CompanyBilling(
                name=billing_row.nameCompany,
                employeesBilling=[employeesBillingAux])
            self.companyList_Billing.append(newCompanyBilling)
        if (len(self.companyList_Billing) == 0):
            newCompanyBilling = CompanyBilling()
            employeesBillingAux = EmployeesBilling(
                name=billing_row.nameEmployees,
                function=billing_row.functionEmployees,
                exams=billing_row.exams,
                data=billing_row.dataExam,
                typeExam=billing_row.typeExam)
            newCompanyBilling = CompanyBilling(
                name=billing_row.nameCompany,
                employeesBilling=[employeesBillingAux])
            self.companyList_Billing.append(newCompanyBilling)

    def getCompanyList(self, company_Billing) -> List[Company]:
        companyListAux = []
        newCompany = Company(company_Billing.name)
        newCompany.employees = []

        company_name_billing = unidecode(company_Billing.name.lower())
        hasName = False
        company_name = ''
        for names in self.namesOfCompanys:  # type:ignore
            # Name of company
            realNameOfCompany = unidecode(names[0].lower())
            # Name of page sheet company
            nameOfPageSheetCompany = unidecode(names[1].lower())
            company_name_billing = company_name_billing.replace('ltda', '')
            company_name_billing = company_name_billing.replace(
                'eireli', '')
            realNameOfCompany = realNameOfCompany.replace('ltda', '')
            realNameOfCompany = realNameOfCompany.replace('eireli', '')
    # print('Empresa que tenho o sheet', realNameOfCompany,
    #       '  | Empresa que ta na tabela de exames:', \
        # company_name_billing)

            if ((realNameOfCompany in company_name_billing)
                    or (nameOfPageSheetCompany in company_name_billing)):
                company_name = names[1]
                hasName = True
                break
        if (not hasName):
            self.companys_not_found.append('\n' + company_Billing.name)
        else:
            list_of_company = []
            getExams(company_name, list_of_company, newCompany)
            namesAndExams = []

            for employees in company_Billing.employeesBilling:
                namesAndExams.append(
                    (employees.name, employees.exams, employees.function,
                        employees.exams, employees.data, employees.typeExam))

            for nameAndExam in namesAndExams:
                dateFormatted = setDateXlsxToString(nameAndExam[4])
                employee_company = Employee()
                employee_company.name = nameAndExam[0]
                employee_company.function = nameAndExam[2]
                employee_company.exams = nameAndExam[3]
                employee_company.typeExam = nameAndExam[5]
                employee_company.date = dateFormatted
                employee_company.examsCost = []
                exams_exact = nameAndExam[1].split('/')

                hasExam = False
                for exam in exams_exact:
                    exam = exam.strip()
                    exam_compar = unidecode(exam.lower())
                    for exam_significant in self.dictionary_exams:
                        examDictionary = unidecode(
                            exam_significant[0].lower())
                        if (exam_compar == examDictionary):
                            exam = exam_significant[1]

                    hasExam = False
                    for company in list_of_company:
                        if (unidecode(exam.lower())
                                in unidecode(company[0].lower())
                                and not hasExam):
                            if (('externo' not in exam.lower() and
                                    'externo' not in company[0].lower())):
                                employee_company.cost += company[1]
                                if (self.hasDetailed):
                                    employee_company.examsCost.append(
                                        (exam, company[1]))
                                hasExam = True
                    if (not hasExam):
                        if (newCompany.missingExams == ''):
                            auxStr = ""
                            newCompany.missingExams = auxStr + exam

                        elif (not (exam in newCompany.missingExams)):
                            newCompany.missingExams = \
                                newCompany.missingExams \
                                + ", " + exam
                newCompany.employees.append(employee_company)
            companyListAux.append(newCompany)
        return companyListAux

    # ----------------------------------------------------------------

    async def getAllExams(self, companyList):
        locale.setlocale(locale.LC_ALL, 'pt_BR')
        monthText: str = self.monthText
        monthNumber = datetime.strptime(monthText, '%B').month
        for i, company in enumerate(companyList):
            if len(company) > 0:
                await self.createSheet(company[0], monthText.upper(),
                                       monthNumber)
                self.progress += 1
                sys.stderr.write(f" {self.progress}%\n")

    async def generateFiles(self):
        companyListAux = []
        for i, companyBilling in enumerate(self.companyList_Billing):
            companyListAux.append(self.getCompanyList(companyBilling))
            self.progress += 1
            sys.stderr.write(f" {self.progress}%\n")
        return companyListAux

    def callGeneratedFiles(self):
        self.setEmployeesFile()
        # self.progressBar.setRange(0, self.billing.maxEmployees)
        asyncio.run(self.generatedBaseDataCompany())

        companys = asyncio.run(self.generateFiles())
        # missingCompanys = [f'Teste {item}' for item in range(100)]
        if len(companys) > 0:
            try:
                os.mkdir(f'{self.folderText}')
            except FileExistsError:
                ...
            except Exception:
                sys.stdout.write(f'Error ao criar pasta {self.folderText}\n')

        asyncio.run(self.getAllExams(companys))

        for company in self.companys_not_found:
            sys.stdout.write(company)


if __name__ == '__main__':
    billing = Faturamento()
    isDetail = False
    if (sys.argv[4] == 'True'):
        isDetail = True
    billing.setParamsBilling(sys.argv[1], sys.argv[2],
                             sys.argv[3], isDetail)
    billing.callGeneratedFiles()