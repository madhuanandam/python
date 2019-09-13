#Importing the Libraries
import xlrd, datetime, sys, os
from datetime import timedelta
from pandas import DataFrame

#Local Variables declaration

startdatelist = []
distsysCsid, distuatCsid, distoatCsid, distliveCsid, setDate = set(), set(), set(), set(), set()
listsysCsid, listDate, listDate2 = [], [], []
setsysCsid, setuatCsid, setoatCsid, setliveCsid = set(), set(), set(), set()
sysCount, uatCount, oatCount, liveCount = 0, 0, 0, 0
finalsyscountcsid, finaluatcountcsid, finaloatcountcsid, finallivecountcsid = 0, 0, 0, 0
livecsid, syscsid, uatcsid, oatcsid = set(), set(), set(), set()
syscomponent, uatcomponent, oatcomponent, livecomponent = [], [], [], []

# Pending components count
pendSyscsid, pendUatcsid, pendOatcsid, pendLivecsid = set(), set(), set(), set()
pendSyscomponent, pendUatcomponent, pendOatcomponent, pendLivecomponent = [], [], [], []
SQLPLSQlCount, D2KCount, UnixCount, XMLCount, ADFCount, APPSCount, PortalCount, DiscovererCount, SOACount, SDFCount, MuleCount, OthersCount = [],[],[],[],[],[],[],[],[],[],[],[]

#Input the values

file_loc=input('Enter the Tracker location: ')
loc = (file_loc+'\Deployment Tracker.xls')
reportName =str(input('Press 1 for MSR; 2 for WSR; 3 for DSR: ' ))
lastrow = int(input('Enter the last row number for the excel : '))
st1 = str(input("Enter the start date in format dd/mm/yy: "))
start_date=datetime.datetime.strptime(st1, '%d/%m/%y')

if reportName == '1':
    for i in range(0, 31):
        modified_date = datetime.datetime.strftime(start_date, "%d/%m/%y")
        startdatelist.append(modified_date)
        start_date = start_date + timedelta(days=1)
        #print(datetime.datetime.strftime(modified_date, "%d/%m/%y"))
elif reportName== '2':
    for i in range(0, 5):
        modified_date = datetime.datetime.strftime(start_date, "%d/%m/%y")
        startdatelist.append(modified_date)
        start_date = start_date + timedelta(days=1)
        #print(datetime.datetime.strftime(modified_date, "%d/%m/%y"))
elif reportName== '3':
    for i in range(0, 1):
        modified_date = datetime.datetime.strftime(start_date, "%d/%m/%y")
        startdatelist.append(modified_date)
        start_date = start_date + timedelta(days=1)
else:
    print('You entered a wrong report name')
    exit()




book = xlrd.open_workbook(loc)
sheet = book.sheet_by_index(0)
#print(sorted(startdatelist))

for date in sorted(startdatelist):
    for rowx in range(0,lastrow ):
        datecell = sheet.cell_value(rowx, colx=0)
        try:
            # print(datecell)
            datecell_as_datetime = datetime.datetime(*xlrd.xldate_as_tuple(datecell, book.datemode))
            #print((sheet.cell_value(rowx, 2)))
            # print(datetime.date(startdate))

            if (str(date) == str(datecell_as_datetime.strftime('%d/%m/%y'))):
                #print(str('datetime: %s' % datecell_as_datetime.strftime('%d/%m/%y')))
                #print(str(date))
                # print(sheet.cell_value(rowx, colx=3))
                if 'Deployed in SYS' in sheet.cell_value(rowx, 4):
                    #print(type(sheet.cell_value(rowx, 2)))
                    #print(sheet.cell_value(rowx, 2))

                    #print(csid)
                    for csidSplit in str(sheet.cell_value(rowx, 2)).split('/'):
                        syscsid.add(csidSplit)
                    syscomponent.append(sheet.cell_value(rowx, 3))
                    #print('SYS is ', (syscsid))

                elif 'Deployed in UAT' in sheet.cell_value(rowx, 4):
                    for csidSplit in str(sheet.cell_value(rowx, 2)).split('/'):
                       uatcsid.add(csidSplit)
                    uatcomponent.append(sheet.cell_value(rowx, 3))
                    # print('UAT is ', uatcsid)

                elif 'Deployed in OAT' in sheet.cell_value(rowx, 4):
                    for csidSplit in str(sheet.cell_value(rowx, 2)).split('/'):
                       oatcsid.add(csidSplit)
                    oatcomponent.append(sheet.cell_value(rowx, 3))
                    # print('OAT is ', oatcsid)

                elif 'Deployed in LIVE' in sheet.cell_value(rowx, 4):
                    for csidSplit in str(sheet.cell_value(rowx, 2)).split('/'):
                       livecsid.add(csidSplit)
                    livecomponent.append(sheet.cell_value(rowx, 3))
                    #print('Live is ', livecsid)

                elif sheet.cell_value(rowx, 4) in ['Deployed in SYS, OAT' , 'Deployed in SYS,OAT'] :

                    for syscsidSplit in str(sheet.cell_value(rowx, 2)).split('/'):
                       syscsid.add(syscsidSplit)
                    syscomponent.append(sheet.cell_value(rowx, 3))
                    for oatcsidSplit in str(sheet.cell_value(rowx, 2)).split('/'):
                       oatcsid.add(syscsidSplit)
                    oatcomponent.append(sheet.cell_value(rowx, 3))
                    # print('Live is ', livecsid)


                elif sheet.cell_value(rowx, 4) in ['Deployed in SYS, UAT' ,'Deployed in SYS,UAT']:
                    for syscsidSplit in str(sheet.cell_value(rowx, 2)).split('/'):
                       syscsid.add(syscsidSplit)
                    syscomponent.append(sheet.cell_value(rowx, 3))

                    for uatcsidSplit in str(sheet.cell_value(rowx, 2)).split('/'):
                       uatcsid.add(uatcsidSplit)
                    uatcomponent.append(sheet.cell_value(rowx, 3))
                     # print('Live is ', livecsid)



                elif sheet.cell_value(rowx, 4) in ['Deployed in UAT, OAT', 'Deployed in UAT,OAT']:
                    for uatcsidSplit in str(sheet.cell_value(rowx, 2)).split('/'):
                       uatcsid.add(uatcsidSplit)
                    uatcomponent.append(sheet.cell_value(rowx, 3))

                    for oatcsidSplit in str(sheet.cell_value(rowx, 2)).split('/'):
                       oatcsid.add(syscsidSplit)
                    oatcomponent.append(sheet.cell_value(rowx, 3))
                     # print('Live is ', livecsid)


                elif sheet.cell_value(rowx, 4) in ['Deployed in SYS,UAT,OAT', 'Deployed in SYS, UAT, OAT']:

                    for syscsidSplit in str(sheet.cell_value(rowx, 2)).split('/'):
                       syscsid.add(syscsidSplit)
                    syscomponent.append(sheet.cell_value(rowx, 3))

                    for uatcsidSplit in str(sheet.cell_value(rowx, 2)).split('/'):
                       uatcsid.add(uatcsidSplit)
                    uatcomponent.append(sheet.cell_value(rowx, 3))

                    for oatcsidSplit in str(sheet.cell_value(rowx, 2)).split('/'):
                       oatcsid.add(syscsidSplit)
                    oatcomponent.append(sheet.cell_value(rowx, 3))
                     # print('Live is ', livecsid)


                elif sheet.cell_value(rowx, 4) in ['Pending in SYS', 'Error in SYS']:
                    for pendingsyscsidSplit in str(sheet.cell_value(rowx, 2)).split('/'):
                       pendSyscsid.add(pendingsyscsidSplit)
                    pendSyscomponent.append(sheet.cell_value(rowx, 3))
                    # print('SYS is ', syscsid)

                elif sheet.cell_value(rowx, 4) in ['Pending in UAT', 'Error in UAT']:
                    for pendinguatcsidSplit in str(sheet.cell_value(rowx, 2)).split('/'):
                       pendUatcsid.add(pendinguatcsidSplit)
                    pendUatcomponent.append(sheet.cell_value(rowx, 3))
                    # print('UAT is ', uatcsid)

                elif sheet.cell_value(rowx, 4) in ['Pending in OAT', 'Error in OAT']:
                    for pendingoatcsidSplit in str(sheet.cell_value(rowx, 2)).split('/'):
                       pendOatcsidatcsid.add(pendingoatcsidSplit)
                    pendOatcomponent.append(sheet.cell_value(rowx, 3))
                    #print('OAT is ', oatcsid)

                elif sheet.cell_value(rowx, 4) in ['Pending in LIVE', 'Error in LIVE']:
                    for pendinglivecsidSplit in str(sheet.cell_value(rowx, 2)).split('/'):
                       pendLivecsid.add(pendinglivecsidSplit)
                    pendLivecomponent.append(sheet.cell_value(rowx, 3))

            if (str(date) == str(datecell_as_datetime.strftime('%d/%m/%y'))):


                if sheet.cell_value(rowx, 5) in ['Type']:
                    print()

                elif sheet.cell_value(rowx, 5).lower() in [('SQL').lower(), ('PL/SQL').lower(), ('PLSQL').lower()] and 'Deployed in LIVE' in sheet.cell_value(rowx, 4):
                    SQLPLSQlCount.append(sheet.cell_value(rowx, 5))


                elif sheet.cell_value(rowx, 5).lower() in [('D2K').lower(), ('FMB').lower(), ('RDF').lower()] and 'Deployed in LIVE' in sheet.cell_value(rowx, 4):
                    D2KCount.append(sheet.cell_value(rowx, 5))

                elif sheet.cell_value(rowx, 5).lower() in [('Unix').lower()] and 'Deployed in LIVE' in sheet.cell_value(rowx, 4):
                    UnixCount.append(sheet.cell_value(rowx, 5))

                elif sheet.cell_value(rowx, 5).lower() in [('RTF').lower()] and 'Deployed in LIVE' in sheet.cell_value(rowx, 4):
                    XMLCount.append(sheet.cell_value(rowx, 5))

                elif sheet.cell_value(rowx, 5).lower() in [('ADF').lower()] and 'Deployed in LIVE' in sheet.cell_value(rowx, 4):
                    ADFCount.append(sheet.cell_value(rowx, 5))

                elif sheet.cell_value(rowx, 5).lower() in [('APPS').lower(), ('config').lower()] and 'Deployed in LIVE' in sheet.cell_value(rowx, 4):
                    APPSCount.append(sheet.cell_value(rowx, 5))

                elif sheet.cell_value(rowx, 5).lower() in [('Portal').lower()] and 'Deployed in LIVE' in sheet.cell_value(rowx, 4):
                    PortalCount.append(sheet.cell_value(rowx, 5))

                elif sheet.cell_value(rowx, 5).lower() in [('Discoverer').lower()] and 'Deployed in LIVE' in sheet.cell_value(rowx, 4):
                    DiscovererCount.append(sheet.cell_value(rowx, 5))

                elif sheet.cell_value(rowx, 5).lower() in [('SOA').lower()] and 'Deployed in LIVE' in sheet.cell_value(rowx, 4):
                    SOACount.append(sheet.cell_value(rowx, 5))

                elif sheet.cell_value(rowx, 5).lower() in [('SDF').lower()] and 'Deployed in LIVE' in sheet.cell_value(rowx, 4):
                    SDFCount.append(sheet.cell_value(rowx, 5))

                elif sheet.cell_value(rowx, 5).lower() in [('Mule').lower()] and 'Deployed in LIVE' in sheet.cell_value(rowx, 4):
                    MuleCount.append(sheet.cell_value(rowx, 5))

                elif sheet.cell_value(rowx, 4) in ['Deployed in LIVE']:
                    OthersCount.append(sheet.cell_value(rowx, 5))

                else:

                    nothing



        # startdate = startdate + datetime.timedelta(days=1)
        except:
            print('', end='')
    finalsyscountcsid = finalsyscountcsid + len(syscsid)
    finaluatcountcsid = finaluatcountcsid + len(uatcsid)
    finaloatcountcsid = finaloatcountcsid + len(oatcsid)
    finallivecountcsid = finallivecountcsid + len(livecsid)
    # print('SYS new is ', finalsyscountcsid)
    # syscsid=[]
    syscsid.clear()
    uatcsid.clear()
    oatcsid.clear()
    livecsid.clear()

finalsyscountcsidList, finaluatcountcsidList, finaloatcountcsidList, finallivecountcsidList=[],[],[],[]
finalsyscountcsidList.append(finalsyscountcsid)
finaluatcountcsidList.append(finaluatcountcsid)
finaloatcountcsidList.append(finaloatcountcsid)
finallivecountcsidList.append(finallivecountcsid)
# print(len(livecsid),len(syscsid),len(uatcsid),len(oatcsid))
# print(syscsid)

os.chdir("C:\\Temp\\report\\") #Changing the current directory

sys.stdout=open("test.txt","w") #Creating a file to store the output data

print("Deployed Components", end='\n\n')
print("SYS "'\t\t'"UAT"'\t\t'"OAT"'\t\t'"LIVE")
print(finalsyscountcsid, '\t\t', finaluatcountcsid, '\t\t', finaloatcountcsid, '\t\t', finallivecountcsid)
# print(len(syscsid),'\t\t', len(uatcsid),'\t\t\t',len(oatcsid), '\t\t\t', len(livecsid))
print(len(syscomponent), '\t\t', len(uatcomponent), '\t\t', len(oatcomponent), '\t\t', len(livecomponent))

print("--------------------------------------------------------")
print("Pending Components", end='\n\n')
print("SYS "'\t\t'"UAT"'\t\t'"OAT"'\t\t'"LIVE")
# print(len(pendSyscsid),'\t\t', len(pendUatcsid),'\t\t\t',len(pendOatcsid), '\t\t\t', len(pendLivecsid))
print(len(pendSyscomponent), '\t\t', len(pendUatcomponent), '\t\t', len(pendOatcomponent), '\t\t',
      len(pendLivecomponent))



print("--------------------------------------------------------")
print("Technology Components deployed in LIVE", end='\n\n')
#print("SQl/PLSQL      D2K         Unix         XML      ADF     Oracle APPS     Portal      Discoverer      SOA     SDF     Mule        Others")
# print(len(pendSyscsid),'\t\t', len(pendUatcsid),'\t\t\t',len(pendOatcsid), '\t\t\t', len(pendLivecsid))
print('SQL/PLSQL Count is       :',len(SQLPLSQlCount))
print('D2K Count is             :',len(D2KCount))
print('Unix Count is            :',len(UnixCount))
print('XML Count is             :',len(XMLCount))
print('ADF Count is             :',len(ADFCount))
print('Oracle APPS Count is     :',len(APPSCount))
print('Portal Count is          :',len(PortalCount))
print('Discoverer Count is      :',len(DiscovererCount))
print('SOA Count is             :',len(SOACount))
print('SDF Count is             :',len(SDFCount))
print('Mule Count is            :',len(MuleCount))
print('Others Count is          :',len(OthersCount))
print('Total count is           :', (len(SQLPLSQlCount)+len(D2KCount)+len(UnixCount)+len(XMLCount)+len(ADFCount)+len(APPSCount)+
                                     len(PortalCount)+len(DiscovererCount)+len(SOACount)+len(SDFCount)+len(MuleCount)+len(OthersCount)))


#dataframe = DataFrame({'SYS':[finalsyscountcsidList,len(syscomponent)], 'UAT':[finaluatcountcsidList,len(uatcomponent)],'OAT':[finaloatcountcsidList,len(oatcomponent)], 'LIVE':[finallivecountcsidList,len(livecomponent)]})
#dataframe = DataFrame([final_list_service])
#print(dataframe)
#print(len(final_list_service))
#print(final_list_service[1][1])
#dataframe.to_excel('test.xlsx', sheet_name='sheet1', index=False)

sys.stdout.close()