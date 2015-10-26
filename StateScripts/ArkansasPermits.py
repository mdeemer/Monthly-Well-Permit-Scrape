#importing modules
from bs4 import BeautifulSoup
import xlsxwriter
import re
import string

# Getting user input for the Permit file
htmlfile = raw_input('''Input file path of the Arkansas permit html file.
''')

#Creating new workbook and worksheet
workbook = xlsxwriter.Workbook("ArkansasNewPermits.xlsx")
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
worksheet.write('T1', 'DATE_IS')		#date_issued

# getting xml as text
soup = BeautifulSoup(open(htmlfile), "lxml")

columns1 = soup.findAll('div', {'style': lambda x : x.startswith('position:absolute; border: textbox 1px solid; writing-mode:lr-tb; left:26px')}) 
columns2 = soup.findAll('div', {'style': lambda x : x.startswith('position:absolute; border: textbox 1px solid; writing-mode:lr-tb; left:69px')}) 
columns3 = soup.findAll('div', {'style': lambda x : x.startswith('position:absolute; border: textbox 1px solid; writing-mode:lr-tb; left:242px')})
columns4 = soup.findAll('div', {'style': lambda x : x.startswith('position:absolute; border: textbox 1px solid; writing-mode:lr-tb; left:400px')})
columns5 = soup.findAll('div', {'style': lambda x : x.startswith('position:absolute; border: textbox 1px solid; writing-mode:lr-tb; left:498px')})
columns6 = soup.findAll('div', {'style': lambda x : x.startswith('position:absolute; border: textbox 1px solid; writing-mode:lr-tb; left:569px')})
table = [columns1, columns2, columns3, columns4, columns5, columns6]

#lists of data by type
permit_number = []
refnum = []
oper_name = []
address = []
city = []
state = []
zip_code = []
well_name = []
location = []
latitude =[]
longitude = []
township = []
sec = []
range = []
field = []
county = []
depth = []
formation = []
final_formation = []
date_issued = []
well_status = []
action = []

#temp/reference date lists
col_data = []
op_end = ["P.O. Box", " 1", " 2", " 3", " 4", " 5", " 6", " 7", " 8", " 9", " 0"]
address_end = ["1 ", "2 ", "3 ", "4 ", "5 ", "6 ", "7 ", "8 ", "9 ", "0 "]
countylist = ["Arkansas", "Ashley", "Baxter", "Benton", "Boone", "Bradley", "Calhoun", "Carroll", "Chicot", "Clark", "Clay", "Cleburne", "Cleveland", "Columbia", "Conway", "Craighead", "Crawford", "Crittenden", "Cross", "Dallas", "Desha", "Drew", "Faulkner", "Franklin", "Fulton", "Garland", "Grant", "Greene", "Hempstead", "Hot Spring", "Howard", "Independence", "Izard", "Jackson", "Jefferson", "Johnson", "Lafayette", "Lawrence", "Lee", "Lincoln", "Little River", "Logan", "Lonoke", "Madison", "Marion", "Miller", "Mississippi", "Monroe", "Montgomery", "Nevada", "Newton", "Ouachita", "Perry", "Phillips", "Pike", "Poinsett", "Polk", "Pope", "Prairie", "Pulaski", "Randolph", "St Francis", "St. Francis", "Saline", "Scott", "Searcy", "Sebastian", "Sevier", "Sharp", "Stone", "Union", "Van Buren", "Washington", "White", "Woodruff", "Yell"]
months = {"January" : "1", "February" : "2", "March" : "3", "April" : "4", "May" : "5", "June" : "6", "July" : "7", "August" : "8", "September" : "9", "October" : "10", "November" : "11", "December" : "12"}
new_permit_data =[]
final_data = []

#Cleaning up column data
def cleaning_columns (columns):
	for column in columns:
		count = 0
		for data in column:
			column[count] = data.get_text()
			column[count] = column[count].encode("ascii", "ignore")
			column[count] = column[count].replace('\n', ' ').replace('\r', ' ').replace("  ", " ")
			column[count] = column[count].strip()
			count = count + 1

cleaning_columns(table)		

# Getting permit numbers
data_breaks = {}	
count = 0
numofrecords = 0
breakcount = 1
for data in columns1:
	if "Permit Number" in data:
		data_breaks.update({breakcount:count} )
		breakcount = breakcount + 1
		count = count + 1
		numofrecords = numofrecords + 1
	else:
		count = count + 1

end = data_breaks[2]
start = data_breaks[1] + 1
count = data_breaks[1] + 1
while end > start:
	permit_number.append((columns1[count]).strip())
	
	count = count + 1
	end = end - 1
	
end = data_breaks[6]
start = data_breaks[5] + 1
count = data_breaks[5] + 1
while end > start:
	permit_number.append((columns1[count]).strip())
	count = count + 1
	end = end - 1	

# Getting second column data
data_breaks = {}	
count = 0
breakcount = 1
for data in columns2:
	if "API Number" in data:
		data_breaks.update({breakcount:count} )
		breakcount = breakcount + 1
		count = count + 1
	else:
		count = count + 1
		
end = data_breaks[2]
start = data_breaks[1] + 1
count = data_breaks[1] + 1
while end > start:
	col_data.append((columns2[count]).strip())
	count = count + 1
	end = end - 1
	
end = data_breaks[6]
start = data_breaks[5] + 1
count = data_breaks[5] + 1
while end > start:
	col_data.append((columns2[count]).strip())
	count = count + 1
	end = end - 1	

#getting refnum data	
count = 0
for data in col_data:
	refnum.append(str(data[0:2] + "-" + data[3:6] + "-" + data[7:12]).strip())
	col_data[count] = (str(data[12:].strip()))
	count = count + 1
	
#Getting Operator data
count = 0
for data in col_data:
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
		col_data[count] = (str(data[data.find(first_string):].strip()))
	count = count + 1

	
#Getting address and city data
count = 0
for data in col_data:
	if any(string in data for string in address_end) == True:
		first_string = ""
		string_position = 0
		if "Ste" in data:
			address_and_city = data[0: data.find(", ")] + " " + data[data.find(", ") + 1:].strip()[0:(data.index(", ")+1)]
			col_data[count] = (str(data[(len(address_and_city) + 2):].strip()))
			count = count + 1
			for string in address_end:
				if string in data:
					position = address_and_city.rfind(string)
			address.append(str(address_and_city[0:(position + 1)].strip())) 
			city.append(str(address_and_city[(position + 1):].strip()))
		else:
			address_and_city = data[0: data.find(",")]
			col_data[count] = (str(data[(len(address_and_city) + 1):].strip()))
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
for data in col_data:
	state.append(str(data[0:3]).strip())
	col_data[count] = (str(data[3:].strip()))
	count = count + 1

#getting zip code data
count = 0
for data in col_data:
	zc = data[0:11]
	if "-" in zc:
		zip_code.append(str(data[0:11]).strip())
		col_data[count] = (str(data[11:].strip()))
		count = count + 1
	else:
		zip_code.append(str(data[0:6]).strip())
		col_data[count] = (str(data[6:].strip()))
		count = count + 1	

		
#reset col_data
col_data = []

# Getting third column data
data_breaks = {}	
count = 0
breakcount = 1
for data in columns3:
	if "Well Name and Number" in data:
		data_breaks.update({breakcount:count} )
		breakcount = breakcount + 1
		count = count + 1
	else:
		count = count + 1
	
end = data_breaks[2]
start = data_breaks[1] + 1
count = data_breaks[1] + 1
while end > start:
	col_data.append((columns3[count]).strip())
	count = count + 1
	end = end - 1
	
end = data_breaks[6]
start = data_breaks[5] + 1
count = data_breaks[5] + 1
while end > start:
	col_data.append((columns3[count]).strip())
	count = count + 1
	end = end - 1	


#Getting well name data
count = 0
for data in col_data:
	well_name.append((data[0:(data.find("SHL:"))]).strip())
	col_data[count] = ((data[(data.find("SHL:")):]).strip())
	count = count + 1

	
	
#Getting location data
count = 0
for data in col_data:
	location_end = ["FWL", "FEL"]
	firststring = (str(data[data.find("SHL:"):]).strip())
	for l in location_end:
		if l in data:
			location.append(str(data[data.find("SHL:") + 5: ((data.find("SHL:") + firststring.find(l)) + len(l))]).strip())
			col_data[count] = str(data[0:(data.find("SHL:"))]) + str(data[((data.find("SHL:") + firststring.find(l)) + len(l)):])
			count = count + 1
			break

count = 0		
for data in col_data:
	if "of Sec. " in data:
		col_data[count] = data[0:(data.index("of Sec. "))] + data[(data.index("of Sec. ") + len("of Sec. ") + 3):]
		count = count + 1
	else:
		count = count + 1
		continue			

#getting T S R data
count = 0
for data in col_data:
	if data[2] == "-" or data[3] == "-" or data[4] == "-":
		sec.append((data[0:data.find("-")]).strip())
		data = data[(data.index("-") + 1):]
		township.append((data[0:data.find("-")]).strip())
		data = data[(data.index("-") + 1):]
		range.append((data[0:data.find(" ")]).strip())
		col_data[count] = ((data[data.index(" "):]).strip())
		count = count + 1
		
	else:
		sec.append("")
		township.append("")
		range.append("")
		count = count + 1
		continue		

#Getting Lattitude data
count = 0
for data in col_data:
	latitude.append((data[(data.index(" N") + len(" N")): data.find("Longitude")]).strip())
	longitude.append((data[(data.index(" W") + len(" W")):]).strip())

#reset col_data	
col_data = []

#Getting column 4 data
print columns4
data_breaks = {}	
count = 0
breakcount = 1
for data in columns1:
	if "Permit Number" in data:
		data_breaks.update({breakcount:count} )
		breakcount = breakcount + 1
		count = count + 1
	else:
		count = count + 1
	
end = data_breaks[2] - 1
start = data_breaks[1]
count = data_breaks[1]
while end > start:
	col_data.append((columns4[count]).strip())
	count = count + 1
	end = end - 1
	
end = data_breaks[6] - 6
start = data_breaks[5] - 5
count = data_breaks[5] - 5
while end > start:
	col_data.append((columns4[count]).strip())
	count = count + 1
	end = end - 1	

print col_data
# Getting county and field data
count = 0
for data in col_data:
	string_position = 0
	for c in countylist:		
		if c in data:
			position = data.index(c)
			if position > string_position:
				string_position = position
			else:
				continue
			county.append(data[string_position:].strip())
			field.append((data[0: string_position]).strip())	

#reseting col_data
col_data = []

# Getting column 5 data
data_breaks = {}	
count = 0
breakcount = 1
for data in columns5:
	if "Depth" in data:
		data_breaks.update({breakcount:count} )
		breakcount = breakcount + 1
		count = count + 1
	else:
		count = count + 1
#getting additional data if it was missed
len497px = soup.findAll('div', {'style': lambda x : x.startswith('position:absolute; border: textbox 1px solid; writing-mode:lr-tb; left:497px')})
count = 0
for data in len497px:
	len497px[count] = data.get_text()
	len497px[count] = len497px[count].encode("ascii", "ignore")
	len497px[count] = len497px[count].replace('\n', ' ').replace('\r', ' ').replace("  ", " ")
	len497px[count] = len497px[count].strip()
	count = count + 1

count = 0	
if len(len497px) > 0:
	for data in len497px:
		if "TVD:" in data:
			depth.append((data[(data.find ("TVD:") + len("TVD:")): data.index("'") + 1]))
			count = count + 1	
		else:
			count = count + 1
	

end = data_breaks[2]
start = data_breaks[1] + 1
count = data_breaks[1] + 1
while end > start:
	col_data.append((columns5[count]).strip())
	count = count + 1
	end = end - 1
	
end = data_breaks[6]
start = data_breaks[5] + 1
count = data_breaks[5] + 1
while end > start:
	col_data.append((columns5[count]).strip())
	count = count + 1
	end = end - 1	


#Getting depth data	
for data in col_data:
	depth.append((data[(data.find ("TVD:") + len("TVD:")): data.index("'") + 1]))

#Resetting col_data
col_data = []
	
# Getting column 6 data
data_breaks = {}	
count = 0
breakcount = 1
for data in columns6:
	if "Zone" in data:
		data_breaks.update({breakcount:count} )
		breakcount = breakcount + 1
		count = count + 1
	else:
		count = count + 1

end = data_breaks[2]
start = data_breaks[1] + 1
count = data_breaks[1] + 1
while end > start:
	formation.append((columns6[count]).strip())
	count = count + 1
	end = end - 1
	
end = data_breaks[6]
start = data_breaks[5] + 1
count = data_breaks[5] + 1
while end > start:
	formation.append((columns6[count]).strip())
	count = count + 1
	end = end - 1		

#compiling final formation data
count = 0
for form in formation:
	final_formation.append((depth[count] + " " + formation[count]).strip())
	count = count + 1

#cleaning soup data
soup = soup.get_text()	
soup = soup.encode("ascii", "ignore")
soup = soup.replace('\n', ' ').replace('\r', ' ').replace("  ", " ")
soup = soup.strip()	

#adding action, well_stat, and date_issued attributes

writtendate = ((soup[(soup.find("Permits Issued During Week Ending:") + len("Permits Issued During Week Ending:")): soup.find("NORTH ARKANSAS DRILLING PERMITS (FORM 2)")]).strip())
records = len(oper_name)
while records > 0:
	well_status.append("ORIG")
	action.append("PERMIT")
	count = 0
	for key in months:
		if writtendate[0: writtendate.find(" ")] == key:
			month = months[key]
			day = ((writtendate[writtendate.index(" "):writtendate.find(",")]).strip())
			year = ((writtendate[(writtendate.index(", ") + 1):]).strip())
			date_issued.append(month + "/" + day + "/" + year)
			count = count + 1
		else:
			count = count + 1
			continue
	records = records -1

#Create lists for one record from list of types of data
list_count = len(oper_name)
count = 0

for operator in oper_name:
	list = [operator]
	list.append(address[count])
	list.append(city[count])
	list.append(state[count])
	list.append(zip_code[count])
	list.append(refnum[count])
	list.append(permit_number[count])
	list.append(well_name[count])
	list.append(field[count])
	list.append(location[count])
	list.append(sec[count])
	list.append(township[count])
	list.append(range[count])
	list.append(county[count])
	list.append(well_status[count])
	list.append(action[count])
	list.append(final_formation[count])
	list.append(latitude[count])
	list.append(longitude[count])
	list.append(date_issued[count])
	new_permit_data.append(list)
	count = count + 1

# deleting records that don't need to be pulled 
biglist_count = 0
for list in new_permit_data:
	if "amend" in list[6] or "renewal" in list[6] or "reentry" in list[6]:
		print "Not a new oil or gas well.  Data not collected."
	else:
		final_data.append(list)
	biglist_count = biglist_count + 1		
	
#Define starting location to write in excel
row = 1
col = 0
	
#Write data to excel		
for OPERATOR, ADDRESS, CITY, ST, ZIP, REF, PERMIT, WELL_NAME, FIELD, LOCATION, S, T, R, COUNTY, WELL_STATU, ACTION, FORMATION, Latitude, Longitude, DATE_IS in (final_data):
	worksheet.write(row, col, OPERATOR)
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

print """Done.

Check Desktop for ArkansasNewPermits.xlsx"""
