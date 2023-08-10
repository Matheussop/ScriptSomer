import pandas as pd
from pandas import DataFrame
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.worksheet import Worksheet
from typing import List


class Employee:
    cost: float
    name: str
    exams: str

    def __init__(self, name: str = '', cost: float = 0.0):
        self.name = name
        self.cost = cost


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
            employeeString.append(f'Nome: {employee_person.name}, '
                                  f'Custo: {employee_person.cost:.2f}')
        return employeeString


class EmployeesBilling:
    name: str
    function: str
    exams: str

    def __init__(self, name: str = '', function: str = '', exams: str = ''):
        self.name = name
        self.function = function
        self.exams = exams


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

    def __init__(self, dataExam: str = '', nameCompany: str = '',
                 nameEmployees: str = '', functionEmployees: str = '',
                 exams: str = ''):
        self.nameCompany = nameCompany
        self.nameEmployees = nameEmployees
        self.functionEmployees = functionEmployees
        self.exams = exams
        self.dataExam = dataExam


def getExams(company_name, list_of_company, newCompany):
    df = pd.read_excel(WORKBOOK_PATH_COMPANY, sheet_name=company_name)

    table = df.iloc[3:19, 0:8]
    # tableOfCompany_ValueTeste = table.iloc[:, [0, 1]]
    # print(tableOfCompany_ValueTeste)
    tableOfCompany_Value = table.iloc[:, [0, 1]]
    # for teste in testeFull:
    for exam in tableOfCompany_Value.values.tolist():
        if (exam[0] is not None):
            list_of_company.append((exam[0], exam[1]))
            if (newCompany.exams != ''):
                newCompany.exams = newCompany.exams + ", " + str(exam[0])
            else:
                newCompany.exams = str(exam[0])


ROOT_FOLDER = Path(__file__).parent
WORKBOOK_PATH_COMPANY = ROOT_FOLDER / 'ValoresEmpresas.xlsx'
WORKBOOK_PATH_FUNC = ROOT_FOLDER / 'examesRealizados.xlsx'

# Carregando um arquivo do excel
workbook_company: Workbook = load_workbook(WORKBOOK_PATH_COMPANY)
workbook_employees: Workbook = load_workbook(WORKBOOK_PATH_FUNC)


sheet_name_employees = 'JUNHO 2023'

worksheet_employees: Worksheet = workbook_employees[sheet_name_employees]

billingList = []
for billing_row in worksheet_employees.iter_rows(min_row=2, values_only=True):
    if (billing_row[0] is not None):
        billingList.append(Billing(
            str(billing_row[0]), str(billing_row[1]),
            str(billing_row[2]), str(billing_row[3]),
            str(billing_row[4])))

companyList_Billing = []
for billing_row in billingList:
    # if (i > 50):
    #     continue
    noHasCompany = True
    for company in companyList_Billing:
        newCompanyBilling = CompanyBilling()
        if (billing_row.nameCompany == company.name):
            employeesBillingAux = EmployeesBilling(
                billing_row.nameEmployees,
                billing_row.functionEmployees,
                billing_row.exams)
            company.employeesBilling.append(employeesBillingAux)
            noHasCompany = False
    if (noHasCompany):
        employeesBillingAux = EmployeesBilling(
            name=billing_row.nameEmployees,
            function=billing_row.functionEmployees,
            exams=billing_row.exams)
        newCompanyBilling = CompanyBilling(
            name=billing_row.nameCompany,
            employeesBilling=[employeesBillingAux])
        companyList_Billing.append(newCompanyBilling)
    if (len(companyList_Billing) == 0):
        newCompanyBilling = CompanyBilling()

        employeesBillingAux = EmployeesBilling(
            name=billing_row.nameEmployees,
            function=billing_row.functionEmployees,
            exams=billing_row.exams)
        newCompanyBilling = CompanyBilling(
            name=billing_row.nameCompany,
            employeesBilling=[employeesBillingAux])
        companyList_Billing.append(newCompanyBilling)

companyList_Billing = [companyList_Billing[0], companyList_Billing[1]]

companys_not_found = 'Empresas não encontradas: '
companyList = []
for company_Billing in companyList_Billing:

    newCompany = Company(company_Billing.name)
    newCompany.employees = []

    company_name_billing = company_Billing.name.lower()
    names_of_workbook = workbook_company.sheetnames
    hasName = False
    company_name = ''
    for names in names_of_workbook:
        if (names.lower() in company_name_billing):
            # Selecionou a planilha da empresa'
            company_name = names
            hasName = True
    if (not hasName):
        companys_not_found = companys_not_found + ', \n' + company_Billing.name
    else:
        list_of_company = []
        # print("Company name:", company_name)
        getExams(company_name, list_of_company, newCompany)
        print(list_of_company, '\n \n')
        name_exam = []
        index = 0
        employees_cost = []

        for employees in company_Billing.employeesBilling:
            name_exam.append((employees.name, employees.exams))

        for name in name_exam:
            employee_company = Employee()
            employee_company.name = name[0]
            exams_exect = name[1].split('/')

            hasExam = False
            for exam in exams_exect:
                exam = exam.strip()
                if (exam == 'CLINICO'):
                    exam = 'Exame Clínico'
                if (exam == 'AUDIO'):
                    exam = 'Audiometria'
                elif (exam == "HEMO"):
                    exam = 'Hemograma c/ Plaquetas'
                elif (exam == 'AC HIPURICO'):
                    exam = 'Ácido Hipúrico'
                elif (exam == 'AC METIL'):
                    exam = 'Ácido Metil Hipúrico'
                elif (exam == 'AC METIL HIPURICO'):
                    exam = 'Ácido Metil Hipúrico'
                elif (exam == 'AC MANDELICO'):
                    exam = 'Ácido Mandélico'
                elif (exam == 'RX DE TORAX'):
                    exam = 'Raio-x de Tórax'
                elif (exam == 'RX DE TORAX'):
                    exam = 'Raio-x de Tórax'
                elif (exam == 'AV PSICOSSOCIAL'):
                    exam = 'Avaliação Psocossocial'
                hasExam = False
                for company in list_of_company:
                    if (exam.lower() in company[0].lower() and not hasExam):
                        employee_company.cost += company[1]
                        hasExam = True
                if (not hasExam):
                    if (newCompany.missingExams == ''):
                        auxStr = "ah empresa não tem o(s) exame(s): "
                        newCompany.missingExams = auxStr + exam

                    elif (not (exam in newCompany.missingExams)):
                        newCompany.missingExams = newCompany.missingExams + ""\
                            ", " + exam
            newCompany.employees.append(employee_company)
        companyList.append(newCompany)

for company in companyList:
    print("========================================================")
    print("COMPANY LIST: \n",  company)
    print("COMPANY LIST OF EMPLOYEES: \n", company.employeesList())
    print("========================================================")
print(companys_not_found)
# workbook_company.save(WORKBOOK_PATH_COMPANY)
# workbook_employees.save(WORKBOOK_PATH_FUNC)
