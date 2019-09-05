import xlrd


loc=("C:\Temp\Deployment-Tracker.xls")


book = xlrd.open_workbook(loc)
sheet = book.sheet_by_index(0)

livecsid,syscsid,uatcsid,oatcsid=[],[],[],[]
syserrcsid, uaterrcsid, oaterrcsid, liveerrcsid=[],[],[],[]

syscomponent,uatcomponent,oatcomponent,livecomponent=[],[],[],[]
syserrcomp, uaterrcomp, oaterrcomp, liveerrcomp=[],[],[],[]

start=int(input("Enter the start value for the row :"))
end=int(input("Enter the ending value of the row :"))

print(start)
print(end)
for rowidx in range(start,end+1):

    #print('a')
    if 'SYS' in sheet.cell_value(rowidx, 4):
        syscsid.append(sheet.cell_value(rowidx, 2))
        syscomponent.append(sheet.cell_value(rowidx, 3))

    if 'UAT' in sheet.cell_value(rowidx, 4):
        uatcsid.append(sheet.cell_value(rowidx, 2))
        uatcomponent.append(sheet.cell_value(rowidx, 3))

    if 'OAT' in sheet.cell_value(rowidx, 4):
        oatcsid.append(sheet.cell_value(rowidx, 2))
        oatcomponent.append(sheet.cell_value(rowidx, 3))

    if 'LIVE' in sheet.cell_value(rowidx, 4):
        livecsid.append(sheet.cell_value(rowidx, 2))
        livecomponent.append(sheet.cell_value(rowidx, 3))

    if 'Error in sys' in sheet.cell_value(rowidx, 4):
        #syserrcsid.append(sheet.cell_value(rowidx, 2))
        syserrcomp.append(sheet.cell_value(rowidx, 3))

    if 'Error in uat' in sheet.cell_value(rowidx, 4):
        #uaterrcsid.append(sheet.cell_value(rowidx, 2))
        uaterrcomp.append(sheet.cell_value(rowidx, 3))

    if 'Error in oat' in sheet.cell_value(rowidx, 4):
        #oaterrcsid.append(sheet.cell_value(rowidx, 2))
        oaterrcomp.append(sheet.cell_value(rowidx, 3))

    if 'Error in live' in sheet.cell_value(rowidx, 4):
        #liveerrcsid.append(sheet.cell_value(rowidx, 2))
        liveerrcomp.append(sheet.cell_value(rowidx, 3))



distsysCsid=list(filter(None,syscsid))
distuatCsid=list(filter(None,uatcsid))
distoatCsid=list(filter(None,oatcsid))
distliveCsid=list(filter(None,livecsid))
print("SYS      UAT         OAT         LIVE")
print(len(distsysCsid),'\t\t', len(distuatCsid),'\t\t\t', len(distoatCsid),'\t\t\t', len(distliveCsid))
print(len(syscomponent),'\t\t', len(uatcomponent),'\t\t\t', len(oatcomponent),'\t\t', len(livecomponent))

print("Printing Pending components")
print("SYS      UAT         OAT         LIVE")
#print(len(syserrcsid),'\t\t', len(uaterrcsid),'\t\t\t', len(oaterrcsid),'\t\t\t', len(liveerrcomp))
print(len(syserrcomp),'\t\t', len(uaterrcomp),'\t\t\t', len(oaterrcomp),'\t\t\t', len(liveerrcomp))
