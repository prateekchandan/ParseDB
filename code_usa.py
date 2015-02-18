#!/usr/bin/python

import xlrd , datetime
import csv
import os.path,sys

mymap = ["application_date","first_name","last_name","ssn","dob","drivers_license_number","drivers_license_state","gender","military_active","amount_requested","residence_type","move_in_date","address1","address2","city","state","zip","phone_home","phone_cell","contact_time","email","ip_address","pay_frequency","net_income","employment_status","employer_name","job_title","hire_date","phone_work","direct_deposit","bank_name","account_type","routing_number","account_number"]
tmp_map = []

def bring_in_order(row):
    global mymap
    global tmp_map
    tm_row = {}
    i = 0
    #print tmp_map
    #print row
    while i < 34:
        #print str(tmp_map.index(mymap[i])) + " : " + row[tmp_map.index(mymap[i])]
        tm_row[i] = row[tmp_map.index(mymap[i])]
        i = i + 1
    return tm_row

def from_excel(file_name):
    global tmp_map
    wb = xlrd.open_workbook(file_name)
    sh = wb.sheet_by_name('Sheet1')

    i=0
    for rownum in xrange(sh.nrows):
        if i == 0:
            i=1
            tmp_map = sh.row_values(rownum)
            continue
        r = sh.row_values(rownum)
        for j in [0,4,27]:
            j = tmp_map.index(mymap[j])
            r[j] = str(datetime.datetime(*xlrd.xldate_as_tuple(r[j], wb.datemode)).strftime("%m/%d/%Y %H:%M"))
        for j in [3,5,8,9,11,16,17,18,19,23,28,29,33]:
            j = tmp_map.index(mymap[j])
            try:
                r[j] = int(r[j])
            except Exception, e:
                pass
            

        r = [str(v) for v in r]
        create_data(r)
        

def from_csv(file_name):
    global tmp_map
    i=0
    with open(file_name, 'rb') as csvfile:
        spamreader = csv.reader(csvfile, delimiter=',', quotechar='"')
        for row in spamreader:
            if i == 0:
                tmp_map = row
                i=1
                continue
            create_data(row)

def create_data(oldrow):
    row = bring_in_order(oldrow)
    filename = row[1]+" "+row[2]+"-"+row[15]
    j = 0
    while os.path.isfile(filename+".txt"):
        j = j + 1
        filename = row[1]+" "+row[2]+"-"+row[15] + "-"+str(j)
       
    print "Created "+filename+".txt"
    f = open(filename+".txt",'w')
    f.write("Application:\t"+row[0]+"\r\n")
    f.write("Amount:\t\t"+row[9]+"\r\n\r\n")
    f.write("Name:\t\t"+row[1]+" "+row[2]+"\r\n")
    f.write("SSN:\t\t"+row[3]+"\r\n")
    f.write("DOB:\t\t"+row[4]+"\r\n\r\n")
    f.write("DL#:\t\t"+row[5]+"\r\n")
    f.write("DL state:\t"+row[6]+"\r\n\r\n")
    f.write("Address:\t"+row[12]+" "+row[13]+"\r\n\t\t"+row[14]+", "+row[15]+"\r\n\t\t"+row[16]+"\r\n\r\n")
    f.write("Phone:\t\t"+row[17]+"\r\n")
    f.write("Mobile:\t\t"+row[18]+"\r\n")
    f.write("Work:\t\t"+row[28]+"\r\n\r\n")
    f.write("Email:\t\t"+row[20]+"\r\n\r\n")
    f.write("Res Type:\t"+row[10]+"\r\n")
    f.write("Residence:\t"+row[11]+"\r\n\r\n")
    f.write("Status:\t\t"+row[24]+"\r\n\r\n")
    f.write("Employer:\t"+row[25]+"\r\n")
    f.write("Tittle:\t\t"+row[26]+"\r\n")
    f.write("Hired:\t\t"+row[27]+"\r\n\r\n")

    f.write("Income:\t\t"+row[23]+"\r\n")
    f.write("Paid:\t\t"+row[22]+"\r\n\r\n")

    f.write("Bank:\t\t"+row[30]+"\r\n")
    f.write("Routing:\t"+row[32]+"\r\n")
    f.write("Account:\t"+row[33]+"\r\n")
    f.write("Account Type:\t"+row[31]+"\r\n\r\n")

    f.write("IP:\t\t"+row[21]+"\r\n")



if len(sys.argv) < 3:
    print "Usage : code.py -c/-x <filename>\r\n-c = the file is csv\r\n-x = the file is excel"
    sys.exit(2)
elif sys.argv[1] != "-c" and sys.argv[1] != "-x":
    print "Usage : code.py -c/-x <filename>\r\n-c = the file is csv\r\n-x = the file is excel"
    sys.exit(2)
elif not os.path.isfile(sys.argv[2]):
    print "Invalid filename"
    sys.exit(2)

file_name = sys.argv[2]

if sys.argv[1] == '-x':
    try:
        from_excel(file_name)
    except Exception, e:
        print "Invalid or unrecognized XLSX format"
        print e.message
else:
    try:
        from_csv(file_name)
    except Exception, e:
        print "Invalid or unrecognized CSV format"
        print e.message