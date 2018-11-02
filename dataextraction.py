import openpyxl

#make a class to find specfic data
class Store:
    def __init__(self, store, presales, gosales):
        self.store = store
        self.presales = presales
        self.gosales = gosales

#funtion to check lists are compatible by first cheching the length on the list are the ame, 
def check(sx, my, key, k1):
    a = len(sx), b = len(my)
    nonm = list(), med = list() 
    while a == b: #because the way i have iterated it they should be the same length
        for i in range(a): #could of been b didn't matter
            sn = sx[i][0].lower(), mn = my[i][0].lower() 
            while sn == mn:
                c = len(sx[i])
                for j in range(1, c):
                    for k in range(2):
                        week = my[i][j][k]
                        if not week.isnum():
                            s = str(J)
                            if s[-1] == "3":
                                col = s + "rd"
                            elif s[-1] == "2":
                                col = s + "nn"
                            elif s[-1] == "1":
                                col = s + "st"
                            else:
                                col = s + "th"
                                print(f"file media has an error on row {i}, on the {col} column") 
                        elif week == 0:
                            temp = sx[i][key - 12: key + 3]
                            nonm.append(temp)
                        else:
                            if key1 < key:
                                print("check correct weeks have been input on row {i} for media start and finish") 
                            temp = sx[i][key - 12: key + (key1 - key)]
                            med.append(temp)
                        med.append(nonm)
                return med #a list that sperates my data so the 0 index is media sales, then in that list it contains string for store, 
            #and a list of atleast 15 lists containing 2 csv, 1 for week, 1 for sales 

#function to create a list within a list to iterate though the week values easier
def liststore(sheets, rnum, calpha):
    v1 = list(), v2 = list(), v3 = list()
    for vi in range(2, rnum + 1):
        for vj in range(2, calpha + 1):
            storename = sheets.ceel(i, 1), v2 == [stroename, [value, sheets.cell(i, j + 1)]]
            if sheets.cell(vi, 1) == sheets.cell(vi - 1, 1):
                temp = "", value = sheet.cell(vi, vj)
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
    for icn in cellname:
        name += icn.title()    
    if name == version:
        return True
    else:
        return False

#function to automate what are in the title cells for each spreadsheet    
def titlemaker(sheet, col):
    cellnames = list()
    for inc in range(1, col + 1):
        temp = sheet.cell(1, inc)
        if not temp.isalpha():
            print("invalid entry on file")
            break
        else:
            cellnames += temp
    return cellnames

#xyz & xyz1 = readable[mediaside] or readable[nonside]
def readSale(xyz):
    abc = len(xyz)
    for irsale in range(abc):
        if xyz == readable[0]:
            irsale += abc + 2
        else:
            xrs = len(xyz[irsale])
            for jrsale in range(1, xrs):
                sale = xyz[i][j][1::2] # [1::2] means only get the values on the right of the right
                if sale.isnum():
                    sale = int(sale)
                return sale
            
def readYAXIS(xyz2):
    abc2 = len(abc2)
    for irweek in range(abc2):
        xrw = len(xyz[irweek])
        for jrweek in range(1, xrw):
            yA = xyz2[irweek][jrweek][::2] # means values on left
            return yA
        
 def readXAxis(xyz1):
    abc1 = len(xyz1)
    for iraxis in range(abc):
        xA = xyz1[iraxis][0]
        return xA  

# this function checks the axis entry if it is numeric we know that it belongs to the x-axis/columns or else it belongs in the y-axis
# /rows. which means i would literally go down the rows if it was the y-axis, and accross the columns if the x-axis 
def writeAxis(axisxy, sheetx):
    lowerxy = str(axisxy).lower()
    if lowerxy.isnum():
        for axisi in range(1, len(axisxy) + 1):
            Axis = sheetx.cell(axisi + 1, 1 )
    else:
        for axisi in range(1, len(axisxy) + 1):
            Axis = sheetx.cell(1, axisi + 1)
    return Axis
    
# so im going to call the value of either media(readable[0]) or non media(readable[1]) in the function readSale with the argument of 
# the sheet im working on them im going to find the length of the lists in the data set(since it's excel we need 2 add 1 to each value
# in the range, the first one as to not include the store name. then after the string name. each index should contain a list of size 2
# therefore i must be consider the sytle of the data, which i intended to plot the prelaunch week
def writeSale(yz, sheetyz):
    zz = len(yz)
    for w1i in range(1, zz + 1):
        az = len(zz[w1i])
        for w1j in range(1, az + 1):
            return sheetyz.cell(w1i, w1j).value = 
        
def whichsheet(argx, as1, as2):
    if argx == readable[mediaside]:
        Wsheet = asheet
    elif argx == readable[nonside]:
        Wsheet = asheet1
    else:
        print(f"unknownn variable {argx} entered")
        break
    return Wsheet

writeSale(readSale(argx, Wsheet)), writeAxis(readYAxis(argx, Wsheet)), writeAxis(readYaxis(argx, Wsheet)))        
#creating a variable to check if true or false
mc = correct(titlemaker(mhm, mcolumns), campaignform), sc = correct(titlemaker(shm, scolumns), salesform)

#creating variables that contain string to the desired locations
sales = "C:\\Sales.xlsx", campaign = "C:\\Campaign.xlsx"

# load workbook, files must be called 'Sales.xlsx' and 'campaign.xlsx' on the c drive
sa = openpyxl.load_workbook(sales), md = openpyxl.load_workbook(campaign)

# creating object of the sheet i want to work on
shs = sa[0], shm = md[0]

#find rows & coiumns
srows = shs.max_row, scolumns = shs.max_column
mrows = shm.max_row, mcolumns = shm.max_column

campaignform = ["Store", "Start", "Finish"]
salesform = ["Store", "Week", "Sales"]

#making the data readable for me the user
if mc not True:
    print(f"ERROR: file {campaign} doesn't appear to be formatted correctly")
    break
elif sc not True:
    print(f"ERROR: file {sales} doesn't appear to be formatted correctly")
    break
else:
    readable = check(liststore(shs, srows, scolumns), liststore(shm, mrows, mcolumns))
    mediaside = 0, nonside = 1

#preparing my sheets and naming them
wk = openpyxl.Workbook()
asheet = wk.active
asheet.title = "Media Performance"

#naming extra sheets in range of 3 because i am creating 3 new sheets.
for sheetnum in range (3):
    if shhetnum == 0:
        sheetTitle = "Analysis"
        sheet0 = sheetTitle
    elif sheetnum == 1:
        sheetTitle = "Non-Media Performance"
        sheet1 = sheetTitle
    else:
        sheetTitle = "Visulisation"
        sheet2 = sheetTitle
    wk.create_sheet = (title = sheetTitle)


                                                                                             
brandname = input("please enter the name of the product you intend to be using here : ")
#saving file
wk.save(f"C:\\{brand}anaylsis.xlsx")
