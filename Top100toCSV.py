import pandas
import numpy
import re
import datetime

# read excel file, create dataframes, and set index for positions
# code won't work without excel file now, need to update it to work with repository
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
for p in p_range:
	p_n = 'p_0'+str(p)
	s_n = 's_0'+str(p)
	e_n = 'e_0'+str(p)
	top100namesn = top100names[top100names[p_n].notnull()]
	for names in top100namesn.itertuples():
		pos_name = top100namesn.loc[names[0], p_n]
		attribs = tuple(names[1:4])
		strtyr = top100namesn.loc[names[0], s_n]
		endyr = top100namesn.loc[names[0], e_n]
		rangeyr = list(range(int(strtyr), int(endyr+1)))
		for yr in rangeyr:
			top100pos.loc[pos_name, yr] = attribs

# supplemental loop: gives me gender and race counts
for y in list(yr_range):
	tuplesyr = top100pos[[y]].dropna().stack()
	races = [t[2] for t in tuplesyr]
	genders = [t[1] for t in tuplesyr]
	top100pos.loc['Total', y] = tuplesyr.count()
	top100pos.loc['Male', y] = genders.count('Male')
	top100pos.loc['Female', y] = genders.count('Female')
	top100pos.loc['Black', y] = races.count('Black')
	top100pos.loc['Asian', y] = races.count('Asian')
	top100pos.loc['White', y] = races.count('White')
	top100pos.loc['Jewish', y] = races.count('Jewish')
	top100pos.loc['Latino', y] = races.count('Latino')

top100pos.to_csv('Top100pos.csv', encoding='utf-8') # data written to csv in UTF-8 encoding
