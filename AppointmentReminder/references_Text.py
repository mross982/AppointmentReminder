
import datetime
import collections

#These are the headers and blank cells in the Centricity reports which I do not want to include in the data pull.
exceptions = [
'',
'Time',
'Site Name',
'Address',
'Patient First',
'Patient Last',
'Phone',
'Language',
'Phone Type',
'Scheduled Date for Site:',
'Cancer Screening',
'Site',
'Contact'
]

# The key is a data type and the value is the column number where you will find that data type from the source sheet (centricity report). Remember to subtract 1 from the column number because they start at 0.
ccat = collections.OrderedDict()
ccat['phone'] = 12
ccat['time'] = 0
ccat['fname'] = 3
ccat['lname'] = 8

ocat = collections.OrderedDict()
ocat['phone'] = 15
ocat['time'] = 0
ocat['fname'] = 3
ocat['lname'] = 9

#This is a way to order the dictionary of data so it can be properly recorded in the log. Again, this is type of data : column number.
logorder = {'record_id': 1, 'fname': 2, 'lname': 3, 'phone': 4, 'location': 5, 'date': 6, 'time': 7,'remindersent': 8}


# These are strings in the locations workbook which indicate there is no screening on the selected date. These will cause the program to end.
noscreen = [
'',
'No Screening',
'NO SCREENING',
'AA OUTREACH',
'AA Outreach',
'RV services'
]

# this removes specific text from the location description in locationsbook ln 85
deltext = [
'MAMMOGRAM SCREENING ',
'SCREENING MAMMOGRAM ',
'MMG SCREENING ',
'MMG SCREENINGS ',
'SCREENING MMG ',
'AA OUTREACH SCREENING ',
'(SCO) ',
'SCO/CORP/ ',
'SCO/CORP ',
'(CTHD) ',
'(CORP) ',
'(SCO/CORP) ',
'NONTRAVIS',
'NONTRAVIS ',
'(SCO/MAP) ',
'CERVICAL- ',
'(CERVICAL) ',
'SCO/MAP',
'RV SERVICES',
'rv services ',
'RV Services '
]
