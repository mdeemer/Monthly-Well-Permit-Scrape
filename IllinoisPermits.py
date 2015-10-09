# Importing modules
from bs4 import BeautifulSoup
import xlsxwriter
import re
import string

# Getting user input for the Permit file
xmlfile = raw_input('''Input file path of the Illinois permit xml file.
''')

# Reading xml file saved from Illinois Permit pdf
soup = BeautifulSoup(open(xmlfile), "lxml")
soup = soup.get_text()
soup = soup.replace('\n', ' ').replace('\r', ' ')
soup = soup.strip()

# Finds broken substrings and corrects them to match search critera
soup = string.replace(soup, 'RN GE:', "RNGE:")

# Defining list variables that data will be appended to
date_issued =[]
oper_name = []
address = []
city = []
state = []
zip_code = []
refnum = []
permit_no = []
well_name = []
location = []
sec = []
township = []
range = [] 
county = []
well_type = []
well_status = []
action = []
drill_oper = []
formation = []

#Create list to write to excel
new_permit_data = []
final_data = []

#Creating list to verify data.
fields_list = ["Oper Name:", "City:", "FORMATION", "WELL TYPE:", "PERMIT #:", "DATE ISSUED:", "REF #:", "OPER #:", "WELL NAME:", "LOCATION:", "SEC:", "TWNSHP:", "RNGE:", "COUNTY:", "WELL STATUS:", "ACTION:", "SURF-ELEV:", "TOOLS:", "DRILL-OPER:", "State:", "Zip:", "Address"]
fields_to_keep = [date_issued, oper_name, address, city, state, zip_code, refnum, permit_no, well_name, location, sec, township, range, county, well_type, well_status, action, drill_oper, formation]

# Creating new workbook and worksheet
workbook = xlsxwriter.Workbook('IllinoisNewPermits.xlsx')
worksheet = workbook.add_worksheet()

#Write headers for excel sheet
worksheet.write('A1', 'DATE_IS')		#date_issued
worksheet.write('B1', 'OPERATOR')		#oper_name
worksheet.write('C1', 'ADDRESS')		#address
worksheet.write('D1', 'CITY')			#city	
worksheet.write('E1', 'STATE')			#state
worksheet.write('F1', 'ZIP')			#zip_code
worksheet.write('G1', 'REF')			#refnum
worksheet.write('H1', 'PERMIT')			#permit_no
worksheet.write('I1', 'WELL_NAME')		#well_name
worksheet.write('J1', 'LOCATION')		#location
worksheet.write('K1', 'S')				#sec
worksheet.write('L1', 'T')				#township
worksheet.write('M1', 'R')				#range
worksheet.write('N1', 'COUNTY')			#county
worksheet.write('O1', 'WELL_TTYPE')		#type
worksheet.write('P1', 'WELL_STATU')		#well_status
worksheet.write('Q1', 'ACTION')			#action
worksheet.write('R1', 'DRILL_OP')		#drill_oper
worksheet.write('S1', 'FORMATION')		#formation

#Define starting location to write in excel
row = 1
col = 0

#Defining variables to append data to lists
append_to_fields = [date_issued, oper_name, address, city, state, zip_code, refnum, permit_no, well_name, location, sec, township, range, county, well_type, well_status, action, drill_oper, formation]
field_names = ["DATE ISSUED:", "Oper Name:", "Address :", "City:", "State:", "Zip:", "REF #:", "PERMIT #:", "WELL NAME:", "LOCATION:", "SEC:", "TWNSHP:", "RNGE:", "COUNTY:", "WELL TYPE:", "WELL STATUS:", "ACTION:", "DRILL-OPER:"]

#Function to pull data from text and appending data to lists
def append_data(lists, strings):
	count = 0
	for string in strings:
		for m in re.finditer(string, soup):
			lists[count].append((str(soup[m.end():(m.end() + 100)])).strip())
		count = count + 1

#Running append data function		
append_data(append_to_fields, field_names)

#Finding all iterations of "FORMATION" and writing them to formation list variable.
for m in re.finditer("FORMATION", soup):
	formation.append((str(soup[(m.end()+4): (m.end() + 100)])).strip())	
	
#removing headings for well type field.
if len(well_type) > len(date_issued):
	well_type.pop(0)
if len(well_status) > len(date_issued):
	well_status.pop(0)

# creating a list of lists to run cleaning loop
fields_to_keep = [date_issued, oper_name, address, city, state, zip_code, refnum, permit_no, well_name, location, sec, township, range, county, well_type, well_status, action, drill_oper, formation]
	
def del_trailing_fields (keepfields, comparefields):
	#removing extra trailing characters for each field
	loop_count = len(comparefields)	
	while loop_count > 0 :
		count = 0
		for list in keepfields:
			for value in list:
				value = value.strip()
				for f in comparefields:
					if f in value:
						for m in re.finditer(f, value):
							list[count] = ((str(value[0: m.start()])).strip())
						break	
					else: 
						continue
				count = count +1
			count = 0
		loop_count = loop_count - 1	

#Running delete trailing fields function		
del_trailing_fields(fields_to_keep, fields_list)

#Cleaning data and remov
welltype_count = 0
for welltype in well_type:
	if welltype.startswith("O") == True or welltype.startswith("G") == True and welltype.startswith("OB") == False and welltype.startswith("OT") == False:
		well_type[welltype_count] = welltype[0]
		welltype_count = welltype_count + 1
	else:
		welltype_count = welltype_count + 1
		continue

#Create lists for one record from list of types of data
list_count = len(date_issued)
count = 0

for date_is in date_issued:
	list = [date_is]
	list.append(oper_name[count])
	list.append(address[count])
	list.append(city[count])
	list.append(state[count])
	list.append(zip_code[count])
	list.append(refnum[count])
	list.append(permit_no[count])
	list.append(well_name[count])
	list.append(location[count])
	list.append(sec[count])
	list.append(township[count])
	list.append(range[count])
	list.append(county[count])
	list.append(well_type[count])
	list.append(well_status[count])
	list.append(action[count])
	list.append(drill_oper[count])
	list.append(formation[count])
	new_permit_data.append(list)
	count = count + 1


# deleting records that don't need to be pulled 
biglist_count = 0
for list in new_permit_data:
	if list[14] == "O" or list[14] == "G":
		final_data.append(list)
	else:
		print "Not oil or gas well.  Data not collected."
	biglist_count = biglist_count + 1	
	
#Write data to excel		
for DATE_IS, OPERATOR, ADDRESS, CITY, STATE, ZIP, REF, PERMIT, WELL_NAME, LOCATION, S, T, R, COUNTY, WELL_TTYPE, WELL_STATU, ACTION, DRILL_OP, FORMATION in (final_data):
	worksheet.write(row, col,     DATE_IS)
	worksheet.write(row, col + 1, OPERATOR)
	worksheet.write(row, col + 2, ADDRESS)
	worksheet.write(row, col + 3, CITY)
	worksheet.write(row, col + 4, STATE)
	worksheet.write(row, col + 5, ZIP)
	worksheet.write(row, col + 6, REF)
	worksheet.write(row, col + 7, PERMIT)
	worksheet.write(row, col + 8, WELL_NAME)
	worksheet.write(row, col + 9, LOCATION)
	worksheet.write(row, col + 10, S)
	worksheet.write(row, col + 11, T)
	worksheet.write(row, col + 12, R)
	worksheet.write(row, col + 13, COUNTY)
	worksheet.write(row, col + 14, WELL_TTYPE)
	worksheet.write(row, col + 15, WELL_STATU)
	worksheet.write(row, col + 16, ACTION)
	worksheet.write(row, col + 17, DRILL_OP)
	worksheet.write(row, col + 18, FORMATION)
	row = row + 1	
	
#Close workbook             
workbook.close()                                                            
print '''Done.
Check Desktop for IllinoisNewPermits.xlsx.'''
