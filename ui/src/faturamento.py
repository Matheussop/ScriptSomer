import pandas as pd
import math
import json
import locale
import sys
import asyncio
import os

from PySide6.QtCore import QObject, Signal
from argparse import ArgumentParser
from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.worksheet import Worksheet
from typing import List
from datetime import datetime
from openpyxl.styles import Font, PatternFill, NamedStyle, Alignment, Border
from openpyxl.styles import Side
from unidecode import unidecode

dictionary_company_names = [
    ("Minas Brasil", "BRASIL COMERCIAL"),
    ("ACTur", "AC TRANSPORTES"),
]


class ErrorBilling(Exception):
    ...


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
            formattedCoast = math.ceil(employee_person.cost * 100)/100
            employeeString.append(f'Nome: {employee_person.name}, '
                                  f'Custo: {formattedCoast}')
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
WORKBOOK_PATH_COMPANY = './files/ValoresEmpresas.xlsx'
WORKBOOK_PATH_FUNC = './files/examesRealizados.xlsx'
hasDetailed = False


def getExams(company_name, list_of_company, newCompany):
    ROW_START = 3
    COLUMN_INDEX = [0, 1]
    try:

        df = pd.read_excel(WORKBOOK_PATH_COMPANY, sheet_name=company_name)
        # Get row 3 to infinity and Column A and B
        tableOfCompany_Value = df.iloc[ROW_START:, COLUMN_INDEX]
        exams = []
        for exam in tableOfCompany_Value.values.tolist():
            if pd.isna(exam[0]):
                break
            if (exam[0] is not None):
                exams.append((exam[0], exam[1]))
                if (newCompany.exams != ''):
                    newCompany.exams = newCompany.exams + ", " + str(exam[0])
                else:
                    newCompany.exams = str(exam[0])
        list_of_company.extend(exams)
    except FileNotFoundError:
        exception_ = ErrorBilling(
            f'Error ao tentar acessar o arquivo: {WORKBOOK_PATH_COMPANY}')
        raise exception_


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


def mountStyleHeaderTable(yearText: str, monthNumber: int, ws: Worksheet,
                          companyName: str):
    columnsTable = ['A', 'B', 'C', 'D', 'E', 'F']

    # Cabeçalho da planilha
    ws.append([f'Mês {monthNumber:02}/{int(yearText[2:])}',
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
    try:
        workbook_company: Workbook = load_workbook(WORKBOOK_PATH_COMPANY)
    except Exception:
        exception_ = ErrorBilling(
            f'Error ao tentar acessar o arquivo: {WORKBOOK_PATH_COMPANY}')
        raise exception_
    # Percorrer o sheet e pegar o nome da empresa dentro da tabela
    names_of_workbook = workbook_company.sheetnames
    namesOfCompanies: List[tuple] = []
    for names in names_of_workbook:
        ws: Worksheet = workbook_company[f'{names}']
        namesOfCompanies.append((ws['B1'].value, names))

    for nameCompany in dictionary_company_names:
        namesOfCompanies.append((nameCompany[1], nameCompany[0]))

    return namesOfCompanies


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


class Faturamento(QObject):
    yearText = ''
    monthText = ''
    folderText = ''
    hasDetailed = False
    dictionary_exams = []
    maxEmployees = 0
    worksheet_employees: Worksheet
    workbook_employees: Workbook
    namesOfCompanies = List[tuple]
    companies_not_found = []
    companyList = []
    companyList_Billing = []
    started = Signal(str)
    progressed = Signal(int)
    finished = Signal(list)
    rangeProgress = Signal(int)
    companiesNotFound = Signal(list)
    progress = 0

    def __init__(self, yearText='', monthText='',
                 folderText='', hasDetailed=False,
                 dictionary_exams=[],
                 namesOfCompanies: List[tuple] = [],
                 parent=None) -> None:
        super().__init__(parent)
        self.yearText = yearText
        self.monthText = monthText
        self.folderText = folderText
        self.hasDetailed = hasDetailed
        self.dictionary_exams = dictionary_exams
        self.namesOfCompanies = namesOfCompanies

    def setParamsBilling(self, yearText, monthText, folderText,
                         hasDetailed) -> None:
        self.yearText = yearText
        self.monthText = monthText
        self.folderText = folderText
        self.hasDetailed = hasDetailed

    def setEmployeesFile(self):
        try:
            # Carregando o arquivo de empresas do excel
            # print(open("WORKBOOK_PATH_COMPANY", "r"))
            self.workbook_employees: Workbook = load_workbook(
                WORKBOOK_PATH_FUNC)
        except Exception:
            exception_ = ErrorBilling(
                f'Error ao tentar acessar o arquivo: {WORKBOOK_PATH_FUNC} ')
            raise exception_

        if (self.workbook_employees):
            name_month = self.monthText.upper()
            year = str(self.yearText)
            sheet_name_employees = f'{name_month} {year}'
            # Carregando o arquivo de exames realizados do excel
            try:
                self.worksheet_employees: Worksheet = \
                    self.workbook_employees[sheet_name_employees]
            except Exception:
                error = f'Error ao tentar encontrar '\
                    'uma página da planilha chamada ' \
                    f'{sheet_name_employees}'
                exception_ = ErrorBilling(error)

                raise exception_

    def getCompaniesNotFound(self): return self.companies_not_found

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

        self.companyList_Billing = []
        for billing_row in billingList:
            self.checkCompanies(billing_row)

        # Bloco para testar numero x de empresas
        # companyList_BillingTeste = []

        # for i in range(2):
        #     companyList_BillingTeste.append(self.companyList_Billing[i])

        # self.companyList_Billing = companyList_BillingTeste

        # self.companyList_Billing = [
        #     self.companyList_Billing[2], self.companyList_Billing[3]]

        # Progress Bar
        self.maxEmployees = len(self.companyList_Billing)
        self.rangeProgress.emit(self.maxEmployees)
        # ----------------------------------------------------

        self.namesOfCompanies = getNameCompanyInSheet()
        self.companies_not_found.append('\nEmpresas não encontradas: ')

    # ----------------------------------------------------------------

    def saveDictionaryExams(self):
        pathFile = './files/dictionary_exams.json'

        with open(pathFile, 'w', encoding='utf8') as arquivo:
            json.dump(
                self.dictionary_exams,
                arquivo,
                ensure_ascii=False,
                indent=4,
            )

    def getDictionaryExams(self) -> List[tuple]:
        newDictionary = []
        pathFile = './files/dictionary_exams.json'
        with open(pathFile, 'r', encoding='utf8') as arquivo:
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

        mountStyleHeaderTable(self.yearText, monthNumber, ws, company.name)

        # cost = 0
        # cost = calculateCost(company)
        # print(cost)
        for item in dataCompany:
            self.progress += 1
            self.progressed.emit(self.progress)
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
        wb.save(f'{folderAndName} {self.yearText}.xlsx')

    def checkCompanies(self, billing_row):
        # if (i > 50):
        #     continue
        noHasCompany = True

        # Para cada empresa criar uma instancia dela com o emprego caso ja
        # exista adicionar apenas o empregado a ela.
        for company in self.companyList_Billing:
            newCompanyBilling = CompanyBilling()
            newCompanyBilling.employeesBilling = []
            employeesBillingAux = EmployeesBilling()
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
        for names in self.namesOfCompanies:  # type:ignore
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
            self.companies_not_found.append('\n' + company_Billing.name)
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
                    examWithoutFormat = exam
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
                                if (isinstance(company[1], (float, int))):
                                    employee_company.cost += company[1]
                                if (self.hasDetailed):
                                    employee_company.examsCost.append(
                                        (examWithoutFormat, company[1]))
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
                # Progress Bar
                self.maxEmployees += len(company[0].employees)
                self.rangeProgress.emit(self.maxEmployees)

                await self.createSheet(company[0], monthText.upper(),
                                       monthNumber)
                self.progress += 1
                self.progressed.emit(self.progress)

    async def generateFiles(self):
        companyListAux = []
        for i, companyBilling in enumerate(self.companyList_Billing):
            companyListAux.append(self.getCompanyList(companyBilling))
            self.progress += 1
            self.progressed.emit(self.progress)
        return companyListAux

    def callGeneratedFiles(self):
        try:
            self.started.emit('Processo iniciado...')
            self.progress = 0
            # Carregando dados do arquivo de exames realizados
            self.setEmployeesFile()
            asyncio.run(self.generatedBaseDataCompany())

            # Definindo limite máximo da barra de progresso
            maxEmployeesAux = self.maxEmployees
            self.maxEmployees += len(self.companyList_Billing) - \
                len(self.companies_not_found)

            # Progress Bar
            self.rangeProgress.emit(self.maxEmployees)
            companies = asyncio.run(self.generateFiles())

            self.progress += maxEmployeesAux
            self.progressed.emit(self.progress)

            # Criando pasta que terá os arquivos de faturamentos
            if len(companies) > 0:
                try:
                    os.mkdir(f'{self.folderText}')
                except FileExistsError:
                    ...
                except Exception:
                    exception_ = ErrorBilling(
                        f'Error ao criar pasta {self.folderText}\n')
                    raise exception_

            # Progress Bar
            self.maxEmployees += len(companies) - len(self.companies_not_found)
            self.rangeProgress.emit(self.maxEmployees)

            # Criando os arquivos de faturamentos e populando.
            asyncio.run(self.getAllExams(companies))
            companiesList = []
            for company in companies:
                if company != []:
                    companiesList.append({
                        'name': company[0].name,
                        'missingExams': company[0].missingExams,
                    })

            # print(companiesList)
            self.finished.emit(companiesList)
            self.companiesNotFound.emit(self.companies_not_found)

            self.progress += maxEmployeesAux
            self.progressed.emit(self.progress)
        except ErrorBilling as errorBilling:
            exception_ = 'Ocorreu um error durante a geração'
            self.finished.emit(['Error: ', errorBilling])


# Parâmetros necessários para executar o script independentemente.
parser = ArgumentParser()

parser.add_argument("-y", "--year",
                    type=str,
                    help='Repassar o ano para a classe Billing')
parser.add_argument("-m", "--month",
                    type=str,
                    help='Repassar o mês para a classe Billing')
parser.add_argument("-f", "--folder",
                    type=str,
                    help='Repassa a pasta gerada para a classe Billing')
parser.add_argument("-d", "--detail",
                    action='store_true',
                    help='Repassa se as planilhas devem ser detalhadas')

args = parser.parse_args()

if __name__ == '__main__':
    if (len(sys.argv) > 1):
        billing = Faturamento()
        isDetail = False
        if (args.detail == 'True'):
            isDetail = True
        billing.setParamsBilling(args.year, args.month,
                                 args.folder, isDetail)

        billing.callGeneratedFiles()
