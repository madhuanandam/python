import xlrd


loc=("C:\Temp\Deployment Tracker_31-Dec-18.xls")

distsysCsid, distuatCsid, distoatCsid, distliveCsid,setDate  =set(), set(), set(), set(),set()
listsysCsid,listDate, listDate2=[],[],[]
setsysCsid, setuatCsid, setoatCsid, setliveCsid=set(),set(),set(),set()
sysCount, uatCount, oatCount, liveCount=0,0,0,0

start=int(input("Enter the start date :"))
end=int(input("Enter the end date :"))

book = xlrd.open_workbook(loc)
sheet = book.sheet_by_index(0)

livecsid,syscsid,uatcsid,oatcsid=[],[],[],[]
syscomponent,uatcomponent,oatcomponent,livecomponent=[],[],[],[]


def csidCount(listRowStart):

            #print(listRowStart)
            if (sheet.cell_value(listRowStart,0) is not None):
                #print(sheet.cell_value(listRowStart,0))

                #print('setDate')
                if 'SYS' in sheet.cell_value(listRowStart, 4):
                    syscsid.append(sheet.cell_value(listRowStart, 2))
                    #syscomponent.append(sheet.cell_value(listRowStart, 3))
                    print('SYS is ',syscsid)

                if 'UAT' in sheet.cell_value(listRowStart, 4):
                    uatcsid.append(sheet.cell_value(listRowStart, 2))
                    #uatcomponent.append(sheet.cell_value(listRowStart, 3))
                    print('UAT is ',uatcsid)

                if 'OAT' in sheet.cell_value(listRowStart, 4):
                    oatcsid.append(sheet.cell_value(listRowStart, 2))
                    #oatcomponent.append(sheet.cell_value(listRowStart, 3))
                    print('OAT is ',oatcsid)

                if 'LIVE' in sheet.cell_value(listRowStart, 4):
                    livecsid.append(sheet.cell_value(listRowStart, 2))
                    #livecomponent.append(sheet.cell_value(listRowStart, 3))
                    print('Live is ', livecsid)


                #print(len(syscomponent))


            else:
                endvalue=lisRowStart
                #print('break')

            return len(syscsid), len(uatcsid), len(oatcsid), len(livecsid)

def component_Count(rowidx,setDate):
    print('in fuct')
    for sheet in book.sheets():
        print('in 1st for')
        for rowidx in range(sheet.nrows):
            print('in 2nd for')
            row = sheet.row(rowidx)
            for colidx, cell in enumerate(row):
                print('in 3rd for')#
                print(rowidx,colidx)


                if cell.value in setDate:
                    if 'SYS' in sheet.cell_value(rowidx, 4):
                        #syscsid.append(sheet.cell_value(rowidx, 2))
                        syscomponent.append(sheet.cell_value(rowidx, 3))

                    if 'UAT' in sheet.cell_value(rowidx, 4):
                        #uatcsid.append(sheet.cell_value(rowidx, 2))
                        uatcomponent.append(sheet.cell_value(rowidx, 3))

                    if 'OAT' in sheet.cell_value(rowidx, 4):
                        #oatcsid.append(sheet.cell_value(rowidx, 2))
                        oatcomponent.append(sheet.cell_value(rowidx, 3))

                    if 'LIVE' in sheet.cell_value(rowidx, 4):
                        #livecsid.append(sheet.cell_value(rowidx, 2))
                        livecomponent.append(sheet.cell_value(rowidx, 3))
                print(len(syscomponent))
                print(len(uatcomponent))
                print(len(oatcomponent))
                print(len(livecomponent))





        #return setsysCsid,setuatCsid,setoatCsid,setliveCsid

print(start)
print(end)
for rowidx in range(start,end+1):
    #print('a')
    if (sheet.cell_value(rowidx, 0)):
        listDate.append(sheet.cell_value(rowidx, 0))
        if 'Date' in listDate:
            listDate.remove('Date')
        setDate = list(filter(None,listDate))
        print(setDate)

        if 'SYS' in sheet.cell_value(rowidx, 4):
            syscsid.append(sheet.cell_value(rowidx, 2))
            # syscomponent.append(sheet.cell_value(listRowStart, 3))
            #setsysCsid=set(syscsid)
            print('SYS is ', syscsid)
        if 'UAT' in sheet.cell_value(rowidx, 4):
            uatcsid.append(sheet.cell_value(rowidx, 2))
            # syscomponent.append(sheet.cell_value(listRowStart, 3))
            #setsysCsid=set(syscsid)
            print('UAT is ', syscsid)
        if 'OAT' in sheet.cell_value(rowidx, 4):
            oatcsid.append(sheet.cell_value(rowidx, 2))
            # syscomponent.append(sheet.cell_value(listRowStart, 3))
            #setsysCsid=set(syscsid)
            print('OAT is ', syscsid)
        if 'LIVE' in sheet.cell_value(rowidx, 4):
            livecsid.append(sheet.cell_value(rowidx, 2))
            # syscomponent.append(sheet.cell_value(listRowStart, 3))
            #setsysCsid=set(syscsid)
            print('LIVE is ', syscsid)



    else:
        print(sheet.cell_value(rowidx, 0))
        print(rowidx)
        #
    print(set(setDate))
    print(listsysCsid)



#print(setDate)

for rowidx in range(start, end + 1):
    # print('a')
    if(sheet.cell_value(rowidx,0)=='Date'):
        newStart=rowidx+1
        for row in range(newStart,end):
            if(sheet.cell_value(row,0) is  None):
                print('not none')
                print(row)
            else:
                newEnd=row
                print(newStart, newEnd)
        listStart=rowidx+1
        #print(rowidx)
        sysCount, uatCount, oatCount, liveCount=csidCount(listStart)
    else:
        print('else')
print('execute func 2')
component_Count(start,setDate)

print("SYS      UAT         OAT         LIVE")
print(sysCount,'\t\t', uatCount,'\t\t\t', oatCount, '\t\t\t', liveCount)
print(len(syscomponent),'\t\t', len(uatcomponent),'\t\t\t', len(oatcomponent),'\t\t', len(livecomponent))

