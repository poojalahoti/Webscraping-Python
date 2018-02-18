import requests
import bs4
import re
import json
import pandas as pd
from pandas import  ExcelWriter


mainLink = "https://www.opensecrets.org/lobby/"
firmLookUp = "https://www.opensecrets.org/lobby/lookup.php"
idRegex = r'id\=([^(]*)\&'
summaryTotalIncomeColHeader = ['Firm', 'Firm Id', 'Year', 'Total Lobbying Income']
lobbyistTotalIncomeHeader = ['Firm', 'Firm Id', 'Year', 'Total Lobbying Income', "Total number of Revolvers", "Total Number of Former Members"]
summaryTableHeader = ['Firm', 'Firm Id', 'Year', 'Client', 'Income', 'Subsidairy', 'Industry']
lobbyistTableHeader = ['Firm', 'Firm Id', 'Year', 'Lobbyist', "Client"]
firmIssuesTableHeader = ['Firm', 'Firm Id', 'Year', 'Issue', "No of Reports", "No of Lobbyist"]
billsTableHeader = ['Firm','Firm Id', 'Year', 'Bill Number',"Congress","Client","Bill Title","No.of Reports"]
agenciesTableHeader = ['Firm','Firm Id', 'Year', 'Agencies',"No.of Reports Listing Agency"]



def saveErrorLog(firm, ex):
    errorLogFile.write(
        "-------Start of error log for firm : " + firm + "-------------------------------------------------------\n")
    errorLogFile.write("Failed to parse: " + firm + "\n")
    errorLogFile.write(str(ex) + "\n")
    errorLogFile.write(
        "-------End of error log for firm : " + firm + "-------------------------------------------------------\n\n")


def parseLobbyingIncome(firm,id,year):
    yearData = requests.get(mainLink + "firmsum.php?" + "id=" + id + "&year=" + year)
    yearBs = bs4.BeautifulSoup(yearData.text, "lxml")
    try:
        income = yearBs.find("p", text=re.compile("Total Lobbying Income")).getText()
        income = income.split(':')
    except:
        saveErrorLog(firm + ' - ' + year, 'Good Exception Data not available')
        income = "N/A"

    summaryTotalIncomeList.append([firm, id, year, str(income[1]).strip()])
    parseLobbyingSummaryTable(firm, id, year, yearBs)


def parseLobbyingSummaryTable(firm,id,year, soup):
    try:
        table = soup.find("table", {"id": "firm_summary"})
        if table is not None:
            tableData = table.find("tbody").find_all('tr')
            for tr in tableData:
                rowList = []
                rowList.append(firm)
                rowList.append(id)
                rowList.append(year)

                tdData = tr.find_all("td")
                for td in tdData:
                    rowList.append(td.getText())
                summaryTableList.append(rowList)
    except:
        saveErrorLog(firm + ' - ' + year, 'Summary Table Parsing Failed')



def parseLobbyistdata(firm,id,year):
    yearData = requests.get(mainLink + "firmlbs.php?" + "id=" + id + "&year=" + year)
    yearBs = bs4.BeautifulSoup(yearData.text, "lxml")

    try:
        income = yearBs.find(text=re.compile("Total Lobbying Income"))
        income = str(income.split(':')[1]).strip()
    except Exception as ex:
        # Log error details for analysis
        saveErrorLog(firm + ' - ' + year, 'Good Exception Data not available')
        income = "N/A"

    try:
        revolvers = yearBs.find(text = re.compile("Total number of revolvers"))
        revolvers = str(revolvers.split(":")[1]).strip()
    except Exception as ex:
        # Log error details for analysis
        saveErrorLog(firm + ' - ' + year, 'Good Exception Data not available')
        revolvers = "N/A"

    try:
        former_mem = yearBs.find(text = re.compile("Total number of former members"))
        former_mem = str(former_mem.split(":")[1]).strip()
    except Exception as ex:
        # Log error details for analysis
        saveErrorLog(firm + ' - ' + year, 'Good Exception Data not available')
        former_mem = "N/A"

    lobbyistTotalIncomeList.append([firm, id, year, income, revolvers, former_mem])

    parseLobbyistTable(firm, id, year, yearBs)


def parseLobbyistTable(firm,id,year, soup):
    try:
        table = soup.find("table", {"id": "firm_lobbyists"})
        if table is not None:
            tableData = table.find("tbody").find_all('tr')
            for tr in tableData:
                rowList = []
                rowList.append(firm)
                rowList.append(id)
                rowList.append(year)

                tdData = tr.find_all("td")
                for td in tdData:

                    anchorTags = td.find_all("a")
                    tdText = ''
                    if(len(anchorTags) > 0):
                        for i in range(len(anchorTags)):
                            if(i == len(anchorTags) - 1):
                              tdText += anchorTags[i].getText()
                            else:
                                tdText += anchorTags[i].getText() + ', '
                    else:
                        tdText = td.getText()

                    rowList.append(tdText)
                lobbyistTableList.append(rowList)
    except:
        saveErrorLog(firm + ' - ' + year, 'Lobbyist Table Parsing Failed')


def parseFirmIssuesTable(firm,id,year):
    yearData = requests.get(mainLink + "firmissues.php?" + "id=" + id + "&year=" + year)
    yearBs = bs4.BeautifulSoup(yearData.text, "lxml")
    try:
        table = yearBs.find("table", {"id": "firm_issues"})
        if table is not None:
            tableData = table.find("tbody").find_all('tr')
            for tr in tableData:
                rowList = []
                rowList.append(firm)
                rowList.append(id)
                rowList.append(year)

                tdData = tr.find_all("td")
                for td in tdData:

                    anchorTags = td.find_all("a")
                    tdText = ''
                    if(len(anchorTags) > 0):
                        for i in range(len(anchorTags)):
                            if(i == len(anchorTags) - 1):
                              tdText += anchorTags[i].getText()
                            else:
                                tdText += anchorTags[i].getText() + ', '
                    else:
                        tdText = td.getText()

                    rowList.append(tdText)
                firmIssuesTableList.append(rowList)
    except:
        saveErrorLog(firm + ' - ' + year, 'Firm Issues Table Parsing Failed')


def parseLobbyingBillsTable(firm,id,year):
    yearData = requests.get(mainLink + "firmbills.php?" + "id=" + id + "&year=" + year)
    yearBs = bs4.BeautifulSoup(yearData.text, "lxml")
    try:
        table = yearBs.find("table", {"id": "client_bills"})
        if table is not None:
            tableData = table.find("tbody").find_all('tr')
            for tr in tableData:
                rowList = []
                rowList.append(firm)
                rowList.append(id)
                rowList.append(year)

                tdData = tr.find_all("td")
                for td in tdData:

                    anchorTags = td.find_all("a")
                    tdText = ''
                    if (len(anchorTags) > 0):
                        for i in range(len(anchorTags)):
                            if (i == len(anchorTags) - 1):
                                tdText += anchorTags[i].getText()
                            else:
                                tdText += anchorTags[i].getText() + ', '
                    else:
                        tdText = td.getText()

                    rowList.append(tdText)
                billsTableList.append(rowList)
    except:
        saveErrorLog(firm + ' - ' + year, 'Bills Table Parsing Failed')

def parseLobbyingAgenciesTable(firm,id,year):
    yearData = requests.get(mainLink + "firmagns.php?" + "id=" + id + "&year=" + year)
    yearBs = bs4.BeautifulSoup(yearData.text, "lxml")
    try:
        table = yearBs.find("table", {"id": "client_issues"})
        if table is not None:
            tableData = table.find("tbody").find_all('tr')
            for tr in tableData:
                rowList = []
                rowList.append(firm)
                rowList.append(id)
                rowList.append(year)

                tdData = tr.find_all("td")
                for td in tdData:

                    anchorTags = td.find_all("a")
                    tdText = ''
                    if (len(anchorTags) > 0):
                        for i in range(len(anchorTags)):
                            if (i == len(anchorTags) - 1):
                                tdText += anchorTags[i].getText()
                            else:
                                tdText += anchorTags[i].getText() + ', '
                    else:
                        tdText = td.getText()

                    rowList.append(tdText)
                agenciesTableList.append(rowList)
    except:
        saveErrorLog(firm + ' - ' + year, 'Agencies Table Parsing Failed')


def parseLobbyingFirm(firm):

    try:
        postData= {'type':'f', 'q':firm}

        r = requests.post(firmLookUp, data= postData )

        if(r.status_code == 200):

            bs = bs4.BeautifulSoup(r.text, "lxml")
            lobbyingLink = bs.select_one("a[href*=firmsum.php]")

            if lobbyingLink is not None:

                # save firm_id
                firmId = re.findall(idRegex, lobbyingLink.get('href'), re.DOTALL)[0]

                lobbyDataRequest = requests.get(mainLink + lobbyingLink.get('href'))

                bs = bs4.BeautifulSoup(lobbyDataRequest.text, "lxml")
                lobbyingYears = bs.select("option[value*=firmsum.php]")

                if(len(lobbyingYears) > 0):

                    for year in lobbyingYears:
                        year1 = year.getText()
                        parseLobbyingIncome(firm, firmId , year1)
                        parseLobbyistdata(firm, firmId , year1)
                        parseFirmIssuesTable(firm, firmId, year1)
                        parseLobbyingBillsTable(firm , firmId, year1)
                        parseLobbyingAgenciesTable(firm, firmId, year1)

            else:
                summaryTotalIncomeList.append([firm, "N/A", "N/A", "N/A"])

        else:
            saveErrorLog(firm, 'No exception. But request failed!')

    except Exception as ex:
        # Log error details for analysis
        saveErrorLog(firm, ex)


# Start of Main Logic

try:
    # Open files for logging
    errorLogFile = open("error_log.txt", 'wb')

    summaryTotalIncomeList = []
    summaryTableList = []
    lobbyistTotalIncomeList = []
    lobbyistTableList = []
    firmIssuesTableList = []
    billsTableList = []
    agenciesTableList = []
    firmSet = set()

    # iterate over firms
    excelFile = pd.ExcelFile('Lobbying Firm Data.xlsx', header=None, index_col=False, index=False).parse(0)
    # excelFile = pd.ExcelFile('exceltest.xlsx', header=None, index_col=False, index=False).parse(0)

    column = excelFile.iloc[:, 0]

    for firm in column:

        firm  = firm.strip()
        if(firm not in firmSet):
            print firm

            firmSet.add(firm)
            parseLobbyingFirm(firm.strip())

    filename1 = 'MainData_Final.xlsx'
    writer = ExcelWriter(filename1)

    # summary total income
    df = pd.DataFrame(summaryTotalIncomeList, columns=summaryTotalIncomeColHeader)
    df.to_excel(writer, 'Summary_Total_Income', index=False)

    # lobbyist total income
    df = pd.DataFrame(lobbyistTotalIncomeList, columns=lobbyistTotalIncomeHeader)
    df.to_excel(writer, 'Lobbyist_Total_Income', index=False)

    # summary table
    df = pd.DataFrame(summaryTableList, columns=summaryTableHeader)
    df.to_excel(writer, 'Summary_Table', index=False)

    # lobbyist table table
    df = pd.DataFrame(lobbyistTableList, columns=lobbyistTableHeader)
    df.to_excel(writer, 'Lobbyist_Table', index=False)

    # firm issues table
    df = pd.DataFrame(firmIssuesTableList, columns=firmIssuesTableHeader)
    df.to_excel(writer, 'Firm_Issues_Table', index=False)

    # agencies table
    df = pd.DataFrame(agenciesTableList, columns=agenciesTableHeader)
    df.to_excel(writer, 'Agencies_Table', index=False)

    # bills table
    df = pd.DataFrame(billsTableList, columns=billsTableHeader)
    df.to_excel(writer, 'Bills_Table', index=False)

    writer.save()

finally:
    # Close opened files
    errorLogFile.close()
