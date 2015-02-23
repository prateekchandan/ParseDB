#!/usr/bin/python

import xlrd , datetime
import csv
import os.path,sys

mymap = ["application date","requested_amount","first_name","last_name","birth_date","email","home_phone","work_phone","cell_phone","address","city","province","postcode","address_length_months","own_home","income_type","employer","job_title","employed_months","monthly_income","pay_frequency","application_date","bank_institution_number","bank_name","bank_branch_number","bank_account_number","bank_account_length_months","direct_deposit","bank_account_type","title","sin","employer_address","employer_city","employer_province","employer_postcode"]
tmp_map = []

def bring_in_order(row):
    global mymap
    global tmp_map
    tm_row = {}
    i = 0
    #print tmp_map
    #print row
    while i < 35:
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
        for j in [0,4,21]:
            j = tmp_map.index(mymap[j])
            r[j] = str(datetime.datetime(*xlrd.xldate_as_tuple(r[j], wb.datemode)).strftime("%m/%d/%Y"))
        for j in [1,6,7,8,10,13,14,18,19,22,24,25,26,27,30,34]:
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
    filename = row[2]+" "+row[3]+"-"+row[11]
    j = 0
    while os.path.isfile(filename+".txt"):
        j = j + 1
        filename = row[2]+" "+row[3]+"-"+row[11] + "-"+str(j)
       
    print "Created "+filename+".txt"
    f = open(filename+".txt",'w')
    f.write("Application:\t"+row[0]+"\r\n")
    f.write("Amount:\t\t"+row[1]+"\r\n\r\n")
    f.write("Name:\t\t"+row[29]+" "+row[2]+" "+row[3]+"\r\n")
    f.write("SIN:\t\t"+row[30]+"\r\n")
    f.write("DOB:\t\t"+row[4]+"\r\n\r\n")
    f.write("Address:\t"+row[9]+"\r\n\t\t"+row[10]+", "+row[11]+"\r\n\t\t"+row[12]+"\r\n\r\n")
    f.write("Phone:\t\t"+row[6]+"\r\n")
    f.write("Mobile:\t\t"+row[8]+"\r\n")
    f.write("Work:\t\t"+row[7]+"\r\n\r\n")
    f.write("Email:\t\t"+row[5]+"\r\n\r\n")
    f.write("Res Type:\t"+row[14]+"\r\n")
    f.write("Residence:\t"+row[13]+"\r\n\r\n")
    f.write("Status:\t\t"+row[15]+"\r\n\r\n")
    f.write("Employer:\t"+row[16]+"\r\n")
    f.write("Tittle:\t\t"+row[17]+"\r\n")
    f.write("Address:\t"+row[31]+"\r\n\t\t"+row[32]+", "+row[33]+"\r\n\t\t"+row[34]+"\r\n\r\n")
    f.write("Length:\t\t"+row[18]+"\r\n")
    f.write("Income:\t\t"+row[19]+"\r\n")
    f.write("Paid:\t\t"+row[20]+"\r\n\r\n")
    f.write("Bank:\t\t"+row[23]+"\r\n")
    f.write("Institution:\t"+row[22]+"\r\n")
    f.write("Transit:\t"+row[24]+"\r\n")
    f.write("Account:\t"+row[25]+"\r\n")
    f.write("Account Type:\t"+row[28]+"\r\n")
    f.write("Account Age:\t"+row[26]+"\r\n")


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