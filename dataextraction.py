import openpyxl

# load workbook, files must be called 'Sales.xlsx' and 'campaign.xlsx' on the c drive
sa = openpyxl.load_workbook("C:\\Sales.xlsx")
md = openpyxl.load_workbook("C:\\Campaign.xlsx")

store = ["Manchester", "Liverpool", "Leeds", "Birmingham", "Glasgow", "london", "Stoke", "Newcastle", "York", "Cardiff"]

# creating object of the sheet i want to work on
shs = sa[0]
shm = md[0]

#find rows & coiumns
srows = shs.max_row
scolumns = shs.max_column
mrows = shm.max_row
mcolumns = shm.max_column

#make a class to find specfic data
class Store:
    def __init__(self, store, sales, week):
        self.store = store
        self.sales = sales
        self.week = week

campaignform = ["Store", "Start", "Finish"]
salesform = ["Store", "Week", "Sales"]

#function to create a list within a list to iterate though the week values easier
def singlestore(sheets, rnum, calpha):
    temp = ""
    v1 = []
    v2 = []
    for i in range(2, rnum + 1):
        for j in range(2, calpha + 1):
            v1 += sheet.cell(i, 1)
            if sheets.cell(i, 1) == sheets.cell(i - 1, 1):
                if sheets.cell(i, j) == NULL:
                    temp = 0
                elif sheets.cell(i, j).isnum():
                    temp = int(sheets.cell(i, j))
                else:
                    temp = sheets.cell(i, j)
                    print(f"{temp} is not a valid entry for media start or end date")
                    break
                v1.append(temp)
            else:
                store = sheet.cell(i, j)
                v2.append(v1)
                v1 = []
    return v2
                

#function to check if file titles are correct
def correct(cellname, version):
    name = [] 
    for i in cellname:
        name += i.title()    
    if name == version:
        return True
    else:
        return False
#function to automate what are in the title cells for each spreadsheet    
def titlemaker(sheet, col):
    cellnames = []
    for i in range(1, col + 1):
        temp = sheet.cell(1, i)
        if not temp.isalpha():
            print("invalid entry on file")
            break
        else:
            cellnames += temp
    return cellnames

#creating a variable to check if true or false
mc = correct(titlemaker(mhm, mcolumns), campaignform)
sc = correct(titlemaker(shm, scolumns), salesform)

while mc & sc == True:
    
