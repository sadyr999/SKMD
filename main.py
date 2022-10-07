import openpyxl
import shutil
import datetime
import glob
import os
import pandas

def backupBase():
    shutil.copy(pathBase, f"BKPbase{currStamp}.xlsx")

def dailyImport(insReport:str):
    bookBase = openpyxl.load_workbook(pathBase)
    sheetBase = bookBase.active
    lastRawBase = sheetBase.max_row+1

    pathIns = insReport
    bookIns = openpyxl.load_workbook(pathIns)
    sheetIns = bookIns.active
    sheetIns.delete_cols(12)
    cellA = sheetIns['L3']
    cellA.value = "INN"
    cellB = sheetIns['M3']
    cellB.value = "PhoneNumber"

    maxRow = sheetIns.max_row
    maxColumn = sheetIns.max_column
    counterA = 4
    while counterA < maxRow:
        phoneCell = sheetIns.cell(row=counterA, column=13)
        phoneOrig = str(phoneCell.value)
        phoneRaw = ""
        for i in phoneOrig:
            if i.isnumeric():
                phoneRaw += i
        phoneRaw = "000000000" + phoneRaw
        phoneCell.value = "996" + phoneRaw[-9:]

        for i in range(1, maxColumn+1):
            sheetBase.cell(row=lastRawBase, column=i).value = sheetIns.cell(row=counterA, column=i).value

        lastRawBase += 1
        counterA += 1

    bookBase.save(pathBase)
    bookIns.save(f"XX 00-00 {pathIns}")
    print(f"MD Import File - {pathIns} - Processed")

def createContacts():
    bookBase = openpyxl.load_workbook(pathBase)
    sheetBase = bookBase.active

    firstEmpty = 2
    lastRowBase = sheetBase.max_row
    while sheetBase.cell(row=firstEmpty, column=14).value != None:
        firstEmpty += 1
    currIndex = str(int(sheetBase.cell(row=firstEmpty-1, column=15).value)+1)
    if firstEmpty <= lastRowBase:
        bookContact = openpyxl.load_workbook("TemplateContacts.xlsx")
        sheetContact = bookContact.active
        bookPrint = openpyxl.load_workbook("TemplatePrintOut.xlsx")
        sheetPrint = bookPrint.active
        counterContact = 2
        while firstEmpty <= lastRowBase:
            sheetBase.cell(row=firstEmpty, column=14).value = "+"+str(sheetBase.cell(row=firstEmpty, column=13).value)
            sheetBase.cell(row=firstEmpty, column=15).value = currIndex
            sheetBase.cell(row=firstEmpty, column=16).value = "НС"+currIndex+" "+sheetBase.cell(row=firstEmpty, column=3).value
            sheetContact.cell(row=counterContact, column=1).value = sheetBase.cell(row=firstEmpty, column=16).value
            sheetContact.cell(row=counterContact, column=21).value = sheetBase.cell(row=firstEmpty, column=14).value
            sheetPrint.cell(row=counterContact, column=1).value = sheetBase.cell(row=firstEmpty, column=16).value
            sheetPrint.cell(row=counterContact, column=2).value = sheetBase.cell(row=firstEmpty, column=14).value
            sheetPrint.cell(row=counterContact, column=3).value = sheetBase.cell(row=firstEmpty, column=2).value
            firstEmpty += 1
            counterContact += 1
        bookBase.save(pathBase)
        nameContact = "НС"+currIndex+" на "+currStamp
        bookContact.save(nameContact+".xlsx")
        bookPrint.save("ZZ Print "+nameContact+".xlsx")
        pandasRead = pandas.read_excel(nameContact+".xlsx")
        pandasRead.to_csv(nameContact+".csv", index=None, header=True, encoding='utf-8')
        os.remove(nameContact+".xlsx")
        print(f"Contacts/Print File - {nameContact}.csv - Created")

def updateCallsBase(file:str) -> str:
    shutil.copy(pathEmer, f"BKPemer{currStamp}.xlsx")
    shutil.copy(pathPlan, f"BKPplan{currStamp}.xlsx")

    bookEmer = openpyxl.load_workbook(pathEmer)
    sheetEmer = bookEmer.active
    bookPlan = openpyxl.load_workbook(pathPlan)
    sheetPlan = bookPlan.active
    bookImport = openpyxl.load_workbook(file)
    sheetImport = bookImport.active

    counterRow = 2
    display = 0
    lastRowEmer = sheetEmer.max_row + 1
    lastRowPlan = sheetPlan.max_row + 1
    lastRowImport = sheetImport.max_row
    print("Processing Receptions File ", end="")
    bookBase = openpyxl.load_workbook(pathBase)
    sheetBase = bookBase.active
    while counterRow <= lastRowImport:
        maxRow = sheetBase.max_row
        policyNum = "НЕТ ПОЛИСА"
        nameIns = sheetImport.cell(row=counterRow, column=4).value
        while maxRow > 1:
            if sheetBase.cell(row=maxRow, column=3).value == nameIns:
                policyNum = sheetBase.cell(row=maxRow, column=4).value
                maxRow = 0
            maxRow -= 1
        if len(sheetImport.cell(row=counterRow, column=2).value) < 10:
            sheetPlan.cell(row=lastRowPlan, column=1).value = str(sheetImport.cell(row=counterRow, column=1).value)[:10]
            for i in range(2, 5):
                sheetPlan.cell(row=lastRowPlan, column=i).value = sheetImport.cell(row=counterRow, column=i).value
            doctor = str(sheetImport.cell(row=counterRow, column=3).value)
            if doctor in open(pathSov).read():
                sheetPlan.cell(row=lastRowPlan, column=5).value = "Медицинский Советник"
            elif doctor in open(pathOrg).read():
                sheetPlan.cell(row=lastRowPlan, column=5).value = "Организация лечения"
            else:
                sheetPlan.cell(row=lastRowPlan, column=5).value = "Специалист"
            sheetPlan.cell(row=lastRowPlan, column=6).value = policyNum
            sheetPlan.cell(row=lastRowPlan, column=7).value = currStamp
            lastRowPlan += 1
        else:
            sheetEmer.cell(row=lastRowEmer, column=1).value = str(sheetImport.cell(row=counterRow, column=1).value)[:10]
            for i in range (2,5):
                sheetEmer.cell(row=lastRowEmer, column=i).value = sheetImport.cell(row=counterRow, column=i).value
            topic = str(sheetImport.cell(row=counterRow, column=5).value)
            if "не" in topic:
                sheetEmer.cell(row=lastRowEmer, column=5).value = "Дежурный врач"
            else:
                sheetEmer.cell(row=lastRowEmer, column=5).value = "Дежурный врач с назначением лечения"
            sheetEmer.cell(row=lastRowEmer, column=6).value = policyNum
            sheetEmer.cell(row=lastRowEmer, column=7).value = currStamp
            lastRowEmer += 1
        if counterRow/lastRowImport*100 > display:
            print(">"+str(display)+"%", end="")
            display += 10
        counterRow += 1
    bookEmer.save(pathEmer)
    bookPlan.save(pathPlan)
    print("")
    rangeReturn = str(sheetImport.cell(row=2, column=1).value)[:10] + "-" + str(sheetImport.cell(row=sheetImport.max_row, column=1).value)[:10]
    return rangeReturn

def smsReportImport(smsReport:str):
    bookBase = openpyxl.load_workbook(pathBase)
    sheetBase = bookBase.active
    bookSMS = openpyxl.load_workbook(smsReport)
    sheetSMS = bookSMS.active

    sheetSMS.delete_cols(9)
    sheetSMS.delete_cols(8)
    sheetSMS.delete_cols(7)
    sheetSMS.delete_cols(6)
    sheetSMS.delete_cols(5)
    sheetSMS.delete_cols(4)
    sheetSMS.delete_cols(2)

    sheetSMS['C1'].value = "Застрахованный"
    sheetSMS['D1'].value = "Полис"
    sheetSMS['E1'].value = "Дата выдачи"
    sheetSMS['F1'].value = "Офис"
    sheetSMS['G1'].value = "Агент"

    maxRowSMS = sheetSMS.max_row
    maxRowBase = sheetBase.max_row

    print("Processing SMS Report ", end="")
    display = 0
    counterA = maxRowSMS
    while counterA > 1:
        phoneCell = sheetSMS.cell(row=counterA, column=2).value
        counterB = counterA - 1
        while counterB > 0:
            if phoneCell == sheetSMS.cell(row=counterB, column=2).value:
                sheetSMS.delete_rows(counterB)
                counterA -= 1
            counterB -= 1
        counterC = 1
        while counterC <= maxRowBase:
            if phoneCell == sheetBase.cell(row=counterC, column=13).value:
                sheetSMS.cell(row=counterA, column=3).value = sheetBase.cell(row=counterC, column=3).value
                sheetSMS.cell(row=counterA, column=4).value = sheetBase.cell(row=counterC, column=4).value
                sheetSMS.cell(row=counterA, column=5).value = sheetBase.cell(row=counterC, column=5).value
                sheetSMS.cell(row=counterA, column=6).value = sheetBase.cell(row=counterC, column=2).value
                sheetSMS.cell(row=counterA, column=7).value = sheetBase.cell(row=counterC, column=1).value
                break
            counterC += 1
            if counterC > maxRowBase:
                sheetSMS.delete_rows(counterA)
        counterA -= 1
        if (maxRowSMS-counterA) / maxRowSMS * 100 > display:
            print(">" + str(display) + "%", end="")
            display += 10

    bookSMS.save(smsReport)
    print("")
    print(f"SMS Report - {smsReport} - Processed")

    bookReport = openpyxl.load_workbook("TemplateReport.xlsx")
    sheetReport = bookReport.active

    maxRowSMS = sheetSMS.max_row
    maxRowBase = sheetBase.max_row
    maxRowReport = sheetReport.max_row
    currMonth = datetime.datetime.now().month
    dateStart = sheetSMS.cell(row=maxRowSMS, column=1).value
    dateEnd = sheetSMS.cell(row=2, column=1).value
    counterD = 4
    sheetReport.cell(row=1, column=1).value = "Отчет по регистрациям в приложении за период с " + str(dateStart) + " по " + str (dateEnd)

    while counterD < maxRowReport:
        quantityTotal = 0
        quantityRegis = 0
        quantityTotalThisM = 0
        quantityRegisThisM = 0
        quantityTotalLastM = 0
        quantityRegisLastM = 0
        quantityTotalPrevM = 0
        quantityRegisPrevM = 0
        counterE = 2
        counterF = 2
        officeName = sheetReport.cell(row=counterD, column=1).value
        while counterE <= maxRowBase:
            if officeName == sheetBase.cell(row=counterE, column=2).value:
                quantityTotal += 1
                if str(currMonth) in str(sheetBase.cell(row=counterE, column=5).value)[5:7]:
                    quantityTotalThisM += 1
                elif str(currMonth-1) in str(sheetBase.cell(row=counterE, column=5).value)[5:7]:
                    quantityTotalLastM += 1
                elif str(currMonth-2) in str(sheetBase.cell(row=counterE, column=5).value)[5:7]:
                    quantityTotalPrevM += 1
            counterE += 1
        while counterF <= maxRowSMS:
            if officeName == sheetSMS.cell(row=counterF, column=6).value:
                quantityRegis += 1
                if str(currMonth) in str(sheetSMS.cell(row=counterF, column=1).value)[3:5]:
                    quantityRegisThisM += 1
                elif str(currMonth-1) in str(sheetSMS.cell(row=counterF, column=1).value)[3:5]:
                    quantityRegisLastM += 1
                elif str(currMonth-2) in str(sheetSMS.cell(row=counterF, column=1).value)[3:5]:
                    quantityRegisPrevM += 1
            counterF += 1
        sheetReport.cell(row=counterD, column=2).value = quantityTotal
        sheetReport.cell(row=counterD, column=3).value = quantityRegis
        sheetReport.cell(row=counterD, column=5).value = quantityTotalThisM
        sheetReport.cell(row=counterD, column=6).value = quantityRegisThisM
        sheetReport.cell(row=counterD, column=8).value = quantityTotalLastM
        sheetReport.cell(row=counterD, column=9).value = quantityRegisLastM
        sheetReport.cell(row=counterD, column=11).value = quantityTotalPrevM
        sheetReport.cell(row=counterD, column=12).value = quantityRegisPrevM
        if quantityTotal != 0:
            sheetReport.cell(row=counterD, column=4).value = str(quantityRegis/quantityTotal*100)[:3] + "%"
        else:
            sheetReport.cell(row=counterD, column=4).value = "NA"
        if quantityTotalThisM != 0:
            sheetReport.cell(row=counterD, column=7).value = str(quantityRegisThisM/quantityTotalThisM*100)[:3] + "%"
        else:
            sheetReport.cell(row=counterD, column=7).value = "NA"
        if quantityTotalLastM != 0:
            sheetReport.cell(row=counterD, column=10).value = str(quantityRegisLastM/quantityTotalLastM*100)[:3] + "%"
        else:
            sheetReport.cell(row=counterD, column=10).value = "NA"
        if quantityTotalPrevM != 0:
            sheetReport.cell(row=counterD, column=13).value = str(quantityRegisPrevM/quantityTotalPrevM*100)[:3] + "%"
        else:
            sheetReport.cell(row=counterD, column=13).value = "NA"
        counterD += 1

    bookReport.save("REP" + smsReport +".xlsx")




#----------------M---A---I---N----------------

pathBase = "base.xlsx"
pathLog = "LOG.txt"
pathEmer = "baseEmer.xlsx"
pathPlan = "basePlan.xlsx"
pathOrg = "listOrg.txt"
pathSov = "listSov.txt"
currStamp = str(datetime.datetime.now())[:10]

with open(pathLog, "a+") as f:
    f.write(f"{str(datetime.datetime.now())}\n")
    f.close();

while True:
    x = input("11 for ОБРАБОТКА СПИСКОВ ЗАСТРАХОВАННЫХ \n22 for ИМПОРТ СМС ОТЧЕТОВ \n33 for ИМПОРТ СПИСКА ПРИЕМОВ \n00 for ВЫХОД \nIN: ")

    if x == "11":
        backupBase()
        for file in glob.glob("Страх*.xlsx"):
            if file in open(pathLog).read():
                pass
            else:
                with open(pathLog, "a+") as f:
                    f.write(f"{file}\n")
                    dailyImport(file)
                    f.close()
        createContacts()

    elif x == "33":
        for file in glob.glob("rece*.xlsx"):
            dateRange = updateCallsBase(file)
            source_path = file
            destination_path = "arhvRecep/recep"+ dateRange + ".xlsx"
            new_location = shutil.move(source_path, destination_path)
            print("File {0} Processed and Moved to \n  > > > > >  {1}".format(source_path, new_location))

    elif x == "22":
        for file in glob.glob("Отчет за период*.xlsx"):
            if file in open(pathLog).read():
                pass
            else:
                with open(pathLog, "a+") as f:
                    f.write(f"{file}\n")
                    smsReportImport(file)
                    f.close()

    elif x == "00":
        break

    else:
        print("--- NOT VALID ---")

print("--- COMPLETED ---")