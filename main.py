import openpyxl
import shutil
import datetime
import glob
import os
import pandas

def backupBase():
    shutil.copy(pathBase, f"arhvBKP/BKPbase{currStamp}.xlsx")

def dailyImport(insReport:str):
    bookBase = openpyxl.load_workbook(pathBase)
    sheetBase = bookBase.active
    lastRawBase = sheetBase.max_row+1

    pathIns = insReport
    bookIns = openpyxl.load_workbook(pathIns)
    sheetIns = bookIns.active
    while True:
        while True:
            x = input(f"{insReport} --- 11-Dinara 22-Venera 33-Web: ")
            if x == "11":
                if input(f"{sheetIns.cell(row=4,column=4).value} / {sheetIns.cell(row=4,column=15).value} press 00 if OK: ") == "00":
                    break
            if x == "22":
                if input(f"{sheetIns.cell(row=4,column=4).value} / {sheetIns.cell(row=4,column=14).value} press 00 if OK: ") == "00":
                    break
            if x == "33":
                if input(f"{sheetIns.cell(row=4,column=3).value} / {sheetIns.cell(row=4,column=14).value} press 00 if OK: ") == "00":
                    break
        if x == "11":
            sheetIns.delete_cols(1)
            sheetIns.delete_cols(12)
            break
        elif x == "22":
            #sheetIns.cell(row=sheetIns.max_row, column=11).value = sheetIns.cell(row=sheetIns.max_row, column=15).value
            sheetIns.delete_cols(15)
            sheetIns.delete_cols(1)
            break
        elif x == 33:
            sheetIns.delete_cols(12)
            break
        else:
            print("INCORRECT INPUT")

    cellA = sheetIns['L3']
    cellA.value = "INN"
    cellB = sheetIns['M3']
    cellB.value = "PhoneNumber"

    maxRow = sheetIns.max_row
    maxColumn = sheetIns.max_column
    counterA = 4
    #while counterA < maxRow:
    while sheetIns.cell(row=counterA, column=2).value != None:
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
    bookIns.save(f"1XX 00-00 {pathIns}")
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
        bookPrint.save("arhvContacts/ZZ Print "+nameContact+".xlsx")
        pandasRead = pandas.read_excel(nameContact+".xlsx")
        pandasRead.to_csv("arhvContacts/" + nameContact +".csv", index=None, header=True, encoding='utf-8')
        os.remove(nameContact+".xlsx")
        print(f"Contacts/Print File - {nameContact}.csv - Created")

def updateCallsBase(file:str) -> str:
    shutil.copy(pathEmer, f"arhvBKP/BKPemer{currStamp}.xlsx")
    shutil.copy(pathPlan, f"arhvBKP/BKPplan{currStamp}.xlsx")
    shutil.copy(pathAll, f"arhvBKP/BKPall{currStamp}.xlsx")

    bookEmer = openpyxl.load_workbook(pathEmer)
    sheetEmer = bookEmer.active
    bookPlan = openpyxl.load_workbook(pathPlan)
    sheetPlan = bookPlan.active
    bookImport = openpyxl.load_workbook(file)
    sheetImport = bookImport.active
    bookAll = openpyxl.load_workbook(pathAll)
    sheetAll = bookAll.active

    display = 0
    lastRowEmer = sheetEmer.max_row + 1
    lastRowPlan = sheetPlan.max_row + 1
    lastRowAll = sheetAll.max_row + 1
    lastRowImport = sheetImport.max_row
    print("Processing Receptions File ", end="")
    bookBase = openpyxl.load_workbook(pathBase)
    sheetBase = bookBase.active
    counterRow = lastRowImport
    print("+base ", end="")
    while counterRow > 1:
        maxRow = sheetBase.max_row
        policyNum = "НЕТ ПОЛИСА"
        officeName = "НЕТ ПОЛИСА"
        nameIns = sheetImport.cell(row=counterRow, column=4).value
        while maxRow > 1:
            if sheetBase.cell(row=maxRow, column=3).value == nameIns:
                policyNum = sheetBase.cell(row=maxRow, column=4).value
                officeName = sheetBase.cell(row=maxRow, column=2).value
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
            sheetPlan.cell(row=lastRowPlan, column=8).value = officeName
            for i in range(1, 9):
                sheetAll.cell(row=lastRowAll, column=i).value = sheetPlan.cell(row=lastRowPlan, column=i).value
            lastRowPlan += 1
            lastRowAll += 1
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
            sheetEmer.cell(row=lastRowEmer, column=8).value = officeName
            for i in range(1, 9):
                sheetAll.cell(row=lastRowAll, column=i).value = sheetEmer.cell(row=lastRowEmer, column=i).value
            lastRowAll += 1
            lastRowEmer += 1
        if counterRow/lastRowImport*100 > display:
            print(">"+str(display)+"%", end="")
            display += 10
        counterRow -= 1
    bookEmer.save(pathEmer)
    bookPlan.save(pathPlan)
    bookAll.save(pathAll)
    print("")
    rangeReturn = str(sheetImport.cell(row=2, column=1).value)[:10] + "-" + str(sheetImport.cell(row=sheetImport.max_row, column=1).value)[:10]
    return rangeReturn

def VerifyCallList():
    bookAll = openpyxl.load_workbook(pathAll)
    sheetAll = bookAll.active

    counterRow = 2
    display = 0

    lastRowAll = sheetAll.max_row
    print("Verifying POLICY CARDS", end="")
    while counterRow <= lastRowAll:
        policyNum = sheetAll.cell(row=counterRow, column=6).value
        if policyNum == "НЕТ ПОЛИСА":
            counterRow += 1
            continue
        counterPolicy = counterRow-1
        policyNew = True
        while counterPolicy > 2:
            if policyNum == sheetAll.cell(row=counterPolicy, column=6).value:
                sheetAll.cell(row=counterRow, column=9).value = "повторное"
                sheetAll.cell(row=counterRow, column=10).value = sheetAll.cell(row=counterPolicy, column=10).value
                if "Дежурный" in sheetAll.cell(row=counterRow, column=5).value:
                    sheetAll.cell(row=counterRow, column=11).value = "дежурный"
                    sheetAll.cell(row=counterRow, column=12).value = sheetAll.cell(row=counterPolicy, column=12).value
                else:
                    sheetAll.cell(row=counterRow, column=11).value = "плановый"
                    cellInqNum = int(sheetAll.cell(row=counterPolicy, column=12).value) + 1
                    sheetAll.cell(row=counterRow, column=12).value = cellInqNum
                policyNew = False
                break
            counterPolicy -= 1
        if policyNew:
            if "Дежурный" in sheetAll.cell(row=counterRow, column=5).value:
                sheetAll.cell(row=counterRow, column=11).value = "дежурный"
                sheetAll.cell(row=counterRow, column=12).value = 0
            else:
                sheetAll.cell(row=counterRow, column=11).value = "плановый"
                sheetAll.cell(row=counterRow, column=12).value = 1
            sheetAll.cell(row=counterRow, column=9).value = "открытие файла"
            sheetAll.cell(row=counterRow, column=10).value = policyNum + " от " + sheetAll.cell(row=counterRow, column=1).value

        if counterRow/lastRowAll*100 > display:
            print(">"+str(display)+"%", end="")
            display += 10
        counterRow += 1
    print("")

    bookAll.save(pathAll)

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
    sheetSMS['H1'].value = "ИНН"

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
                sheetSMS.cell(row=counterA, column=8).value = sheetBase.cell(row=counterC, column=12).value
                break
            counterC += 1
            if counterC > maxRowBase:
                sheetSMS.delete_rows(counterA)
        counterA -= 1
        if (maxRowSMS-counterA) / maxRowSMS * 100 > display:
            print(">" + str(display) + "%", end="")
            display += 10

    bookSMS.save(f"PRCD-{smsReport}")
    print("")
    print(f"SMS Report - {smsReport} - Processed")

def ReportPRCD(smsReport: str):
    bookBase = openpyxl.load_workbook(pathBase)
    sheetBase = bookBase.active
    bookSMS = openpyxl.load_workbook(smsReport)
    sheetSMS = bookSMS.active
    bookCalls = openpyxl.load_workbook(pathAll)
    sheetCalls = bookCalls.active

    d1 = str(input("Day 1 of the week / mm-dd : "))
    d2 = str(input("Day 2 of the week / mm-dd : "))
    d3 = str(input("Day 3 of the week / mm-dd : "))
    d4 = str(input("Day 4 of the week / mm-dd : "))
    d5 = str(input("Day 5 of the week / mm-dd : "))
    d6 = str(input("Day 6 of the week / mm-dd : "))
    d7 = str(input("Day 7 of the week / mm-dd : "))
    s1 = d1[3] + d1[4] + "." + d1[0] + d1[1]
    s2 = d2[3] + d2[4] + "." + d2[0] + d2[1]
    s3 = d3[3] + d3[4] + "." + d3[0] + d3[1]
    s4 = d4[3] + d4[4] + "." + d4[0] + d4[1]
    s5 = d5[3] + d5[4] + "." + d5[0] + d5[1]
    s6 = d6[3] + d6[4] + "." + d6[0] + d6[1]
    s7 = d7[3] + d7[4] + "." + d7[0] + d7[1]

    bookReport = openpyxl.load_workbook("TemplateReport.xlsx")
    sheetReport = bookReport.active

    maxRowSMS = sheetSMS.max_row
    maxRowBase = sheetBase.max_row
    maxRowReport = sheetReport.max_row - 2
    maxRowCalls = sheetCalls.max_row

    # move all data
    counterFcol = 47
    counterFrow = 2
    for i in range(1, 12):
        while counterFrow <= sheetReport.max_row:
            for j in range(0, 4):
                sheetReport.cell(row=counterFrow, column=counterFcol+4+j).value = \
                    sheetReport.cell(row=counterFrow, column=counterFcol+j).value
            counterFrow += 1
        counterFcol = counterFcol - 4
        counterFrow = 2

    dateEnd = sheetSMS.cell(row=2, column=1).value
    counterD = 4
    sheetReport.cell(row=1, column=6).value = "по " + str(dateEnd)[:9]
    sheetReport.cell(row=2, column=7).value = "неделя " + d1 + " по " + d7
    femaleWeekReg = 0
    maleWeekReg = 0
    femaleWeekIns = 0
    maleWeekIns = 0
    femaleTotalReg = 0
    maleTotalReg = 0
    femaleTotalIns = 0
    maleTotalIns = 0

    print("Processing Report... ", end="")
    display = 0
    while counterD < maxRowReport:
        if counterD / maxRowReport * 100 > display:
            print(">" + str(display) + "%", end="")
            display += 3
        quantityTotal = 0
        quantityRegis = 0
        quantityTotalLastW = 0
        quantityRegisLastW = 0
        quantityCalls = 0
        quantityCallsLastW = 0
        counterE = 2
        counterF = 2
        officeName = sheetReport.cell(row=counterD, column=2).value

        #print("Counting new members..")
        while counterE <= maxRowBase:
            if officeName == sheetBase.cell(row=counterE, column=2).value:
                if sheetBase.cell(row=counterE, column=20).value == "uniq":
                    quantityTotal += 1
                    if str(sheetBase.cell(row=counterE, column=12).value)[0] == "1":
                        femaleTotalIns += 1
                    else:
                        maleTotalIns += 1
                    dayAndMonth = str(sheetBase.cell(row=counterE, column=5).value)[-5:]
                    if d1 in dayAndMonth or d2 in dayAndMonth or d3 in dayAndMonth or \
                            d4 in dayAndMonth or d5 in dayAndMonth or d6 in dayAndMonth or d7 in dayAndMonth:
                        quantityTotalLastW += 1
                        if str(sheetBase.cell(row=counterE, column=12).value)[0] == "1":
                            femaleWeekIns += 1
                        else:
                            maleWeekIns += 1
            counterE += 1
        #print("Counting new registrations..")
        while counterF <= maxRowSMS:
            if officeName == sheetSMS.cell(row=counterF, column=6).value:
                quantityRegis += 1
                if str(sheetSMS.cell(row=counterF, column=8).value)[0] == "1":
                    femaleTotalReg += 1
                else:
                    maleTotalReg += 1
                dayAndMonth = str(sheetSMS.cell(row=counterF, column=1).value)[:5]
                if s1 in dayAndMonth or s2 in dayAndMonth or s3 in dayAndMonth or s4 in dayAndMonth or s5 in dayAndMonth or s6 in dayAndMonth or s7 in dayAndMonth:
                    quantityRegisLastW += 1
                    if str(sheetSMS.cell(row=counterF, column=8).value)[0] == "1":
                        femaleWeekReg += 1
                    else:
                        maleWeekReg += 1
            counterF += 1
        #print("Counting receptions..")
        counterF = 2
        while counterF <= maxRowCalls:
            if officeName == sheetCalls.cell(row=counterF, column=8).value:
                quantityCalls += 1
                dayAndMonth = str(sheetCalls.cell(row=counterF, column=1).value)[-5:]
                if d1 in dayAndMonth or d2 in dayAndMonth or d3 in dayAndMonth or \
                        d4 in dayAndMonth or d5 in dayAndMonth or d6 in dayAndMonth or d7 in dayAndMonth:
                    quantityCallsLastW += 1
            counterF += 1
        sheetReport.cell(row=counterD, column=3).value = quantityTotal
        sheetReport.cell(row=counterD, column=4).value = quantityRegis
        sheetReport.cell(row=counterD, column=7).value = quantityTotalLastW
        sheetReport.cell(row=counterD, column=8).value = quantityRegisLastW
        sheetReport.cell(row=counterD, column=6).value = quantityCalls
        sheetReport.cell(row=counterD, column=10).value = quantityCallsLastW
        counterD += 1

    sheetReport.cell(row=maxRowReport + 1, column=3).value = maleTotalIns
    sheetReport.cell(row=maxRowReport + 2, column=3).value = femaleTotalIns
    sheetReport.cell(row=maxRowReport + 1, column=4).value = maleTotalReg
    sheetReport.cell(row=maxRowReport + 2, column=4).value = femaleTotalReg
    sheetReport.cell(row=maxRowReport + 1, column=7).value = maleWeekIns
    sheetReport.cell(row=maxRowReport + 2, column=7).value = femaleWeekIns
    sheetReport.cell(row=maxRowReport + 1, column=8).value = maleWeekReg
    sheetReport.cell(row=maxRowReport + 2, column=8).value = femaleWeekReg

    print(">>100!>>>")
    #bookReport.save("Отчет по регистрациям в Мой Доктор на " + currStamp + ".xlsx")
    bookReport.save("TemplateReport.xlsx")
def checkUniqueClient():
    bookBase = openpyxl.load_workbook(pathBase)
    sheetBase = bookBase.active
    maxRowBase = sheetBase.max_row
    counterE = 3
    print("Processing Base... ", end="")
    display = 0
    while counterE <= maxRowBase:
        if counterE / maxRowBase * 100 > display:
            print(">" + str(display) + "%", end="")
            display += 3
        if sheetBase.cell(row=counterE, column=20).value == None:
            sheetBase.cell(row=counterE, column=20).value = "rept"
            counterDuplicate = counterE - 1
            while sheetBase.cell(row=counterE, column=3).value != sheetBase.cell(row=counterDuplicate, column=3).value:
                counterDuplicate -= 1
                if counterDuplicate == 1:
                    sheetBase.cell(row=counterE, column=20).value = "uniq"
                    break
        counterE += 1
    print(">>100!>>>")
    bookBase.save(pathBase)
def transformDocList(docList:str):
    bookList = openpyxl.load_workbook(docList)
    sheetList = bookList.active
    bookNew = openpyxl.load_workbook("TemplateDocs.xlsx")
    sheetNew = bookNew.active

    maxRowList = sheetList.max_row
    counterA = 6
    counterNew = 1
    while counterA <= maxRowList:
        #print("start")
        counterB = counterA - 1
        while counterB > 0:
            #print(str(counterA) + " / " + str(counterB) + str(sheetList.cell(row=counterA, column=3).value) + str(sheetNew.cell(row=counterB, column=2).value))
            if (sheetList.cell(row=counterA, column=1).value == sheetNew.cell(row=counterB, column=1).value and
                    sheetList.cell(row=counterA, column=3).value == sheetNew.cell(row=counterB, column=2).value and
                    sheetList.cell(row=counterA, column=7).value == sheetNew.cell(row=counterB, column=4).value):
                counterA += 1
                counterB = counterA - 1
                if counterA > maxRowList:
                    break
            else:
                counterB -= 1
        if counterA > maxRowList:
            break
        counterColumn = 1
        for i in [1, 3, 6, 7]:
            sheetNew.cell(row=counterNew, column=counterColumn).value = sheetList.cell(row=counterA, column=i).value
            counterColumn += 1
        insCase = str(sheetNew.cell(row=counterNew, column=4).value)
        sheetNew.cell(row=counterNew, column=5).value = "заявление"
        sheetNew.cell(row=counterNew, column=6).value = "ожидаем"

        if "смерть" in insCase:
            counterNew += 1
            for i in [1, 2, 3, 4, 6]:
                sheetNew.cell(row=counterNew, column=i).value = sheetNew.cell(row=counterNew-1, column=i).value
            sheetNew.cell(row=counterNew, column=5).value = "свидетельство о смерти нот.зав.копия"
        elif "инвалидность" in insCase:
            counterNew += 1
            for i in [1, 2, 3, 4, 6]:
                sheetNew.cell(row=counterNew, column=i).value = sheetNew.cell(row=counterNew-1, column=i).value
            sheetNew.cell(row=counterNew, column=5).value = "нот.зав.копия МСЭК"
        elif "критическое" in insCase:
            for i in [1, 2, 3, 4, 6]:
                sheetNew.cell(row=counterNew, column=i).value = sheetNew.cell(row=counterNew-1, column=i).value
            sheetNew.cell(row=counterNew, column=5).value = "эпикриз оригинал"
        counterNew += 1
        counterA += 1


    bookNew.save(f"БФоригиналы-{str(sheetNew.cell(row=1, column=1).value)[:5]}-{str(sheetNew.cell(row=counterNew-1, column=1).value)[:5]}.xlsx")
    #print(sheetNew.max_row)
    print(f"DOCS Report - {docList} - Processed")

#----------------M---A---I---N----------------

pathBase = "base.xlsx"
pathLog = "LOG.txt"
pathAll = "baseAllCalls.xlsx"
pathEmer = "baseEmer.xlsx"
pathPlan = "basePlan.xlsx"
pathOrg = "listOrg.txt"
pathSov = "listSov.txt"
currStamp = str(datetime.datetime.now())[:10]

with open(pathLog, "a+") as f:
    f.write(f"{str(datetime.datetime.now())}\n")
    f.close();

while True:
    x = input("11 for ОБРАБОТКА СПИСКОВ ЗАСТРАХОВАННЫХ \n22 for ИМПОРТ СМС ОТЧЕТОВ \n"
              "33 for ИМПОРТ СПИСКА ПРИЕМОВ \n44 for СПИСОК ДОКОВ СК \n00 for ВЫХОД \nIN: ")

    if x == "11":
        backupBase()
        for file in glob.glob("Страх*.xlsx"):
            if file in open(pathLog).read():
                pass
            else:
                with open(pathLog, "a+") as f:
                    f.write(f"{file}\n")
                    f.close()
                    dailyImport(file)
                    destination_path = "arhvIns/recep" + file
                    shutil.move(file, destination_path)
        for file in glob.glob("E-Strah*.xlsx"):
            if file in open(pathLog).read():
                pass
            else:
                with open(pathLog, "a+") as f:
                    f.write(f"{file}\n")
                    f.close()
                    dailyImport(file)
                    destination_path = "arhvIns/" + file
                    shutil.move(file, destination_path)
        createContacts()

    elif x == "33":
        for file in glob.glob("rece*.xlsx"):
            dateRange = updateCallsBase(file)
            source_path = file
            destination_path = "arhvRecep/recep" + dateRange + ".xlsx"
            new_location = shutil.move(source_path, destination_path)
            print("File {0} Processed and Moved to \n  > > > > >  {1}".format(source_path, new_location))
            VerifyCallList()
    elif x == "22":
        for file in glob.glob("Отчет за период*.xlsx"):
            if file in open(pathLog).read():
                pass
            else:
                with open(pathLog, "a+") as f:
                    #f.write(f"{file}\n")
                    f.close()
                    smsReportImport(file)
        for file in glob.glob("PRCD*.xlsx"):
            ReportPRCD(file)
    elif x == "44":
        for file in glob.glob("доки*.xlsx"):
            if file in open(pathLog).read():
                pass
            else:
                with open(pathLog, "a+") as f:
                    f.write(f"{file}\n")
                    f.close()
                    transformDocList(file)
    elif x == "66":
        for file in glob.glob("PRCD*.xlsx"):
            ReportPRCD(file)
    elif x == "88":
        checkUniqueClient()
    elif x == "00":
        break

    else:
        print("--- NOT VALID ---")

print("--- COMPLETED ---")