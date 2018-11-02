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

#funtion to check lists are compatible by first cheching the length on the list are the ame, 
def check(sx, my):
    a = len(sx)
    b = len(my)
    nonm = list()
    med = list()
    while a == b:
        for i in range(a):
            sn = sx[i][0].lower()
            mn = my[i][0].lower()
            while sn == mn:
                c = len(sx[i])
                for j in range(1, c):
                    for k in range(2):
                        week = my[i][j][k]
                        if not week.isnum():
                            s = str(J)
                            if s[-1] == 3:
                                col = s + "rd"
                            elif s[-1] == 2:
                                col = s + "nn"
                            elif s[-1] == 1:
                                col = s + "st"
                            else:
                                col = s + "th"
                                print(f"file {my} has an error on row {i}, column {col}") 
                        elif week == 0:
                            nonm += 
            
    

#function to create a list within a list to iterate though the week values easier
def liststore(sheets, rnum, calpha):
    v1 = list()
    v2 = list()
    v3 = list()
    for i in range(2, rnum + 1):
        for j in range(2, calpha + 1):
            storename = sheets.ceel(i, 1)
            v2 == [stroename, [value, sheets.cell(i, j + 1)]]
            if sheets.cell(i, 1) == sheets.cell(i - 1, 1):
                temp = ""
                value = sheet.cell(i, j)
                if store == NULL:
                    temp = 0
                elif store.isnum():
                    temp = int(value)
                else:
                    temp = value
                    print(f"{temp} is not a valid entry for media start or end date")
                    break
                v1.append(temp)
            elif len(v1) == 2:
                v2.append(v1)
                v1 = list()
            else:
                v3.append(v2)
                v2 = list()
    return v3
                

#function to check if file titles are correct
def correct(cellname, version):
    name = list() 
    for i in cellname:
        name += i.title()    
    if name == version:
        return True
    else:
        return False
#function to automate what are in the title cells for each spreadsheet    
def titlemaker(sheet, col):
    cellnames = list()
    for i in range(1, col + 1):
        temp = sheet.cell(1, i)
        if not temp.isalpha():
            print("invalid entry on file")
            break
        else:
            cellnames += temp
    return cellnames

#function to 


#creating a variable to check if true or false
mc = correct(titlemaker(mhm, mcolumns), campaignform)
sc = correct(titlemaker(shm, scolumns), salesform)

if mc not True:
    print(f"ERROR: file {md} doesn't appear to be formatted correctly")
    break
elif sc not True:
    print(f"ERROR: file {md} doesn't appear to be formatted correctly")
    break
else:
    uglystore = liststore(shs, srows, scolumns)
    uglymedia = liststore(shm, mrows, mcolumns)



            
