from bs4 import BeautifulSoup
import re
import xlsxwriter
import string


# Getting user input for the Permit file
xmlfile = raw_input('''Input file path of the Arkansas permit xml file.
''')

soup = BeautifulSoup(open(xmlfile), "lxml")
soup = soup.get_text()
soup = soup.replace('\n', ' ').replace('\r', ' ')
soup = soup.strip()

# Defining list variables that data will be appended to
oper_name = []
address = []
city = []
state = []
zip_code = []
refnum = []
permit_no = []
well_name = []
field = []
location = []
sec = []
township = []
range = [] 
county = []
well_status = []
action = []
depth = []
formation = []
final_formation = []
latitude = []
longitude = []
date_issued =[]

# Creating new workbook and worksheet
workbook = xlsxwriter.Workbook('ArkansasNewPermits.xlsx')
worksheet = workbook.add_worksheet()

#Write headers for excel sheet
worksheet.write('A1', 'OPERATOR')		#oper_name
worksheet.write('B1', 'ADDRESS')		#address
worksheet.write('C1', 'CITY')			#city
worksheet.write('D1', 'ST')				#state	
worksheet.write('E1', 'ZIP')			#zip_code
worksheet.write('F1', 'REF')			#refnum
worksheet.write('G1', 'PERMIT')			#permit_no
worksheet.write('H1', 'WELL_NAME')		#well_name
worksheet.write('I1', 'FIELD')			#field
worksheet.write('J1', 'LOCATION')		#location
worksheet.write('K1', 'S')				#sec
worksheet.write('L1', 'T')				#township
worksheet.write('M1', 'R')				#range
worksheet.write('N1', 'COUNTY')			#county
worksheet.write('O1', 'WELL_STATU')		#well_status
worksheet.write('P1', 'ACTION')			#action
worksheet.write('Q1', 'FORMATION')		#final_formation
worksheet.write('R1', 'Latitude')		#latitude
worksheet.write('S1', 'Longitude')		#longitude
worksheet.write('S1', 'DATE_IS')		#date_issued

#Define starting location to write in excel
row = 1
col = 0


#lists used to pull and clean data
permit_data = []
data = []
final_data = []
new_permit_data = []
op_end = ["P.O. Box", " 1", " 2", " 3", " 4", " 5", " 6", " 7", " 8", " 9", " 0"]
address_end = ["1 ", "2 ", "3 ", "4 ", "5 ", "6 ", "7 ", "8 ", "9 ", "0 "]
wellname_end = ["/", "rd", "st", "nd", "th"]

#Pulling only Form 2 data
North_form2 = soup[(soup.index('Anticipated Zone of Completion') + (len('Anticipated Zone of Completion')+1)):(soup.index('NORTH ARKANSAS WELL COMPLETIONS (FORM 3)')-48)]
South_form2 = soup[(soup.index('SOUTH ARKANSAS DRILLING PERMITS (FORM 2)') + (len('SOUTH ARKANSAS DRILLING PERMITS (FORM 2)')+ 80)):(soup.index('SOUTH ARKANSAS WELL COMPLETIONS (FORM 3)')-1)]

#Cleaning up Form 2 text
North_form2 = North_form2.replace('\n', ' ').replace('\r', ' ')
South_form2 = South_form2.replace('\n', ' ').replace('\r', ' ')
North_form2 = North_form2.strip()
South_form2 = South_form2.strip()
North_form2 = North_form2.encode("ascii", "ignore")
South_form2 = South_form2.encode("ascii", "ignore")

# combining form data and splitting data up by permit
text = North_form2 + South_form2
count = text.count("   ")
while count > 0:
	data.append(text[0:text.index("   ")])
	text = text[(text.index("   ") + 2):]
	count = count - 1
data.append(text[0:])

# removing empty data in list
for d in data:
	if d != "":
		if "renewal" not in d and "amend" not in d:
			permit_data.append(str(d).strip())
	else:
		continue

# Pulling data from strings in data list and appending them to each attribute list
#getting permit number data
count = 0
for data in permit_data:
	permit_no.append(str(data[0:5]).strip())
	permit_data[count] = (str(data[5:].strip()))
	count = count + 1

#getting ref number data
count = 0
for data in permit_data:
	refnum.append(str(data[0:2] + "-" + data[3:6] + "-" + data[7:12]).strip())
	permit_data[count] = (str(data[12:].strip()))
	count = count + 1

#Getting Operator data
count = 0
for data in permit_data:
	if any(string in data for string in op_end) == True:
		first_string = ""
		string_position = 100000
		for string in op_end:
			if string in data:
				position = data.find(string)
				if position < string_position:
					first_string = string
					string_position = position
		oper_name.append(str(data[0:data.find(first_string)].strip())) 
		permit_data[count] = (str(data[data.find(first_string):].strip()))
	count = count + 1

#Getting address data
count = 0
for data in permit_data:
	if any(string in data for string in address_end) == True:
		first_string = ""
		string_position = 0
		if "Ste" in data:
			address_and_city = data[0: data.find(", ")] + " " + data[data.find(", ") + 1:].strip()[0:(data.index(", ")+1)]
			permit_data[count] = (str(data[(len(address_and_city) + 2):].strip()))
			count = count + 1
			for string in address_end:
				if string in data:
					position = address_and_city.rfind(string)
			address.append(str(address_and_city[0:(position + 1)].strip())) 
			city.append(str(address_and_city[(position + 1):].strip()))
		else:
			address_and_city = data[0: data.find(",")]
			permit_data[count] = (str(data[(len(address_and_city) + 1):].strip()))
			for string in address_end:
				if string in data:
					position = address_and_city.find(string)
					if position > string_position:
						string_position = position
			address.append(str(address_and_city[0:(string_position + 1)].strip())) 
			city.append(str(address_and_city[(string_position + 1):].strip()))
			count = count + 1

#getting State data
count = 0
for data in permit_data:
	state.append(str(data[0:3]).strip())
	permit_data[count] = (str(data[3:].strip()))
	count = count + 1
	
#getting zip code data
count = 0
for data in permit_data:
	zc = data[0:11]
	if "-" in zc:
		zip_code.append(str(data[0:11]).strip())
		permit_data[count] = (str(data[11:].strip()))
		count = count + 1
	else:
		zip_code.append(str(data[0:6]).strip())
		permit_data[count] = (str(data[6:].strip()))
		count = count + 1
		
#removing extra data from string
count = 0
for data in permit_data:
	if any(string in data for string in wellname_end) == True:
		first_string = ""
		string_position = 100000
		for string in wellname_end:
			if string in data:
				position = data.find(string)
				if position < string_position:
					first_string = string
					string_position = position
		if "/" in data:
			firststr = str(data[(data.find(first_string)-2):].strip())
			permit_data[count] = (str(firststr[firststr.find(" "):].strip()))
		else: 
			firststr = str(data[(data.find(first_string)-2):].strip())
			secondstr = str(firststr[(firststr.find(" ")+ 6):].strip())
			permit_data[count] = (str(secondstr[secondstr.find(" "):].strip()))
	count = count + 1

#Getting depth data
count = 0
for data in permit_data:
	firststring = (str(data[data.find("TVD:"):]).strip())
	depth.append(str(data[data.find("TVD:") + 5: ((data.find("TVD:") + firststring.find("' ")) +1)]).strip())
	permit_data[count] = (str(data[0:data.find("TVD:")]) + str(data[(data.find("TVD:") + 10):]))
	count = count + 1


print depth 
print "\n"			
print permit_data[0]
print "\n"
print permit_data[1]
print "\n"
print permit_data[2]
print "\n"
print permit_data[3]
print "\n"
print permit_data[4]
print "\n"
print permit_data[5]
print "\n"


#Create lists for one record from list of types of data
list_count = len(date_issued)
count = 0

for operator in oper_name:
	list = [oper_name]
	list.append(address[count])
	list.append(city[count])
	list.append(state[count])
	list.append(zip_code[count])
	list.append(refnum[count])
	list.append(permit_no[count])
	#list.append(well_name[count])
	#list.append(field[count])
	#list.append(location[count])
	#list.append(sec[count])
	#list.append(township[count])
	#list.append(range[count])
	#list.append(county[count])
	#list.append(well_status[count])
	#list.append(action[count])
	#list.append(final_formation[count])
	#list.append(latitude[count])
	#list.append(longitude[count])
	#list.append(date_issued[count])
	new_permit_data.append(list)
	count = count + 1
	
#Write data to excel		
for OPERATOR, ADDRESS, CITY, ST, ZIP, REF, PERMIT, WELL_NAME, FIELD, LOCATION, S, T, R, COUNTY, WELL_STATU, ACTION, FORMATION, Latitude, Longitude, DATE_IS in (final_data):
	worksheet.write(row, col,     OPERATOR)
	
	worksheet.write(row, col + 1, ADDRESS)
	worksheet.write(row, col + 2, CITY)
	worksheet.write(row, col + 3, ST)
	worksheet.write(row, col + 4, ZIP)
	worksheet.write(row, col + 5, REF)
	worksheet.write(row, col + 6, PERMIT)
	worksheet.write(row, col + 7, WELL_NAME)
	worksheet.write(row, col + 8, FIELD)
	worksheet.write(row, col + 9, LOCATION)
	worksheet.write(row, col + 10, S)
	worksheet.write(row, col + 11, T)
	worksheet.write(row, col + 12, R)
	worksheet.write(row, col + 13, COUNTY)
	worksheet.write(row, col + 14, WELL_STATU)
	worksheet.write(row, col + 15, ACTION)
	worksheet.write(row, col + 16, FORMATION)
	worksheet.write(row, col + 17, Latitude)
	worksheet.write(row, col + 18, Longitude)
	worksheet.write(row, col + 19, DATE_IS)
	row = row + 1	
	
#Close workbook             
workbook.close()                                                            
print '''Done.

Check Desktop for ArkansasNewPermits.xlsx.'''
