# __author__ = 'josh'

import mechanize
import re
import xlrd

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
    br.addheaders.append( ['Accept-Encoding','gzip'] ) # AA page seems to be GZIPED
    br.open('http://www.theaa.com/route-planner/classic/planner_main.jsp')
    br.select_form(name="routePlanner")
    br["fromPlace"] = ps1
    br["toPlace"] = ps2
    br.submit()  # AA routeplanner returns a confidence check for the post codes

    # we know we are happy with the post codes so submit again
    br.select_form(name="routePlanner")
    response = br.submit()
    for x in response:
        match = re.search('miles', x)
        if match:
            break   # it's messy but me know the first match is the mileage
    x = x.lstrip() # returns with black left chars
    x = round(float(x[:5]), -1) # round to closest 10
   # print '{} to {} is {} miles'.format(ps1, ps2, x)
    return x

def matchset(our_ps, postcodelist):
    """
    :param our_ps: our current post code
    :param postcodelist: a list of postcodes to match
    :return:
    """
    for ps in postcodelist:
        get_mileage(our_ps, ps)

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
        index[sn.value] = sheet.cell(row, 7).value
        row += 1

    return index

index = openbook('/home/josh/Documents/test.xls')
print 'from {} the following mileages are tested'.format(OUR_PS)
for key in index:
    print '{} to home address postcode of : {} is {}'.format(key, index[key], get_mileage(OUR_PS, index[key]))



