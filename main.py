# __author__ = 'josh'

import mechanize
import re
import xlrd
import xlwt
import datetime

OUR_PS = 'PL112BD'

def get_mileage(ps1, ps2):
    """
    :param ps1: The home post code
    :param ps2: Our current post code
    :return: The mileage between the two from AA routefinder
    """
    br = mechanize.Browser()
    br.set_handle_robots(False)
    br.addheaders = [('User-agent', 'Firefox')]
    br.addheaders.append( ['Accept-Encoding','gzip'] )
    br.open('http://www.theaa.com/route-planner/classic/planner_main.jsp')
    br.select_form(name="routePlanner")
    br["fromPlace"] = ps1
    br["toPlace"] = ps2
    br.submit()  # AA routeplanner returns a confidence check for the post codes

    # we know we are happy with the post codes so submit again
    br.select_form(name="routePlanner")
    response = br.submit()
    for y in response:
        match = re.search('miles', y)

        if match:
            y = y.lstrip() # returns without left chars
            y = round(float(y[:5]), -1) # round to closest 10
            rsp = y
            break
        else:
            rsp = '{}-ISBAD'.format(ps2)
    return rsp

def confidence(PS):
    """
    :param Check a postcode for confidence
    :return: True/False
    """
    POSTCODE_FORMATS = ['\w\w\d\w\d\w\w', '\w\d\w\d\w\w', '\w\d\d\w\w', '\w\d\d\d\w\w', '\w\w\d\d\w\w', '\w\w\d\d\d\w\w', '\w\w\d\d\w\w']

    for X in POSTCODE_FORMATS:
        match = re.match(X, PS.replace(" ", "")) # STRIP WHITESPACE BEFORE MATCHING
    if match:
        return True
    else:
        return False

def openbook(workbook):
    """
    :param workbook: filename of workbook (XLS)
    :return: Returns a dict of service number and postcodes
    """
    openedbook = xlrd.open_workbook(workbook)
    sheet = openedbook.sheet_by_name('Sheet1')
    row = 0
    index = dict()
    for sn in sheet.col(0):
        if confidence(sheet.cell(row, 1).value):
            index[sn.value] = sheet.cell(row, 1).value

        else:
            print 'Bad postcode {} at row {}'.format(sheet.cell(row, 1).value, row)
        row += 1
    return index

index = openbook('test.xls')
print 'from {} the following mileages are tested'.format(OUR_PS)
coll_output = []
for key in index:
    in_tuple = (key, index[key], get_mileage(OUR_PS, index[key]))
    coll_output.append(in_tuple)
    print 'key {} from {} is {}'.format(*in_tuple)

i = raw_input('Do you wish to dump the return to a file? (Y/N): ')

if i.upper() == 'Y':
    date = datetime.date.today()
    out_name = 'output-' + str(date.year) + str(date.month) + str(date.day) +'.xls'
    print 'dumping to XLS output-{}.xls'.format(out_name)
    target = xlwt.Workbook()
    sheet = target.add_sheet('OUTPUT')
    row = 0
    for x in range(0, len(coll_output)):
        sheet.write(row, 0, label = coll_output[x][0])
        sheet.write(row, 1, label = coll_output[x][1])
        sheet.write(row, 2, label = coll_output[x][2])
        row += 1
    print 'wrote {} rows'.format(row)
    target.save(out_name)






