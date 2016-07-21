import pandas
import numpy
import re
import datetime

# read excel file, create dataframes, and set index for positions
top100names = pandas.read_excel('Top100.xlsx', 'Data')
positions = pandas.read_excel('Top100.xlsx', 'Positions', index_col='Title')

top100names.convert_objects(convert_numeric=True)

# index set by positions list
# columns set by range of oldest first position start year and current year plus one
yr_range = range(min(set(top100names['s_01'])), datetime.datetime.now().year+1)
top100pos = pandas.DataFrame(index=positions.index, columns=list(yr_range))

# isolate headers and search for digits to use in position counter
headers = list(top100names.columns)
p_range = set([int((d.group(1))) for h in headers for d in (re.compile('(\d+)').search(h),) if d])

# main loop: populates top100pos
for p in p_range: # iterates through positions held
	p_n = 'p_0'+str(p) # p_n name
	s_n = 's_0'+str(p) # p_n start year
	e_n = 'e_0'+str(p) # p_n end year
	top100namesn = top100names[top100names[p_n].notnull()] # filters out NaN in p_n 
	for names in top100namesn.itertuples(): # iterates through names
		pos_name = top100namesn.loc[names[0], p_n] # finds p_n for current name
		attribs = tuple(names[1:4]) # collects attributes for current name
		strtyr = top100namesn.loc[names[0], s_n] # start year of p_n for current name
		endyr = top100namesn.loc[names[0], e_n] # end year of p_n for current name
		rangeyr = list(range(int(strtyr), int(endyr+1))) # current p_n range for current name
		for yr in rangeyr: # iterates through current p_n range for current name
			top100pos.loc[pos_name, yr] = attribs # writes attribs to relevant position row and year column

# supplemental loop: gives me gender and race counts
for y in list(yr_range): # iterates through years 
	tuplesyr = top100pos[[y]].dropna().stack() # dataframe indexed by positions for year y
	races = [t[2] for t in tuplesyr] # all races for year y
	genders = [t[1] for t in tuplesyr] # all genders for year y
	top100pos.loc['Total', y] = tuplesyr.count() # count written to total row column y
	top100pos.loc['Male', y] = genders.count('Male') # count written to male row column y
	top100pos.loc['Female', y] = genders.count('Female') # count written to female row column y 
	top100pos.loc['Black', y] = races.count('Black') # count written to black row column y
	top100pos.loc['Asian', y] = races.count('Asian') # count written to asian row column y
	top100pos.loc['White', y] = races.count('White') # count written to white row column y
	top100pos.loc['Jewish', y] = races.count('Jewish') # count written to jewish row column y
	top100pos.loc['Latino', y] = races.count('Latino') # count written to latino row column y

top100pos.to_csv('Top100pos.csv', encoding='utf-8') # data written to csv in UTF-8 encoding
