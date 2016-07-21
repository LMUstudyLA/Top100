import pandas
import numpy
import re
import datetime
import xlsxwriter


# read excel file, create dataframes, and set index for positions
top100names = pandas.read_excel('Top100.xlsx', 'Data')
positions = pandas.read_excel('Top100.xlsx', 'Positions', index_col='Title')
columns = pandas.read_excel('Top100.xlsx', 'Columns', index_col='Title')
rows = pandas.read_excel('Top100.xlsx', 'Rows', index_col='Title')

top100names.convert_objects(convert_numeric=True)

# index set by positions list
# columns set by range of oldest first position start year and current year plus one
yr_range = range(min(set(top100names['s_01'])), datetime.datetime.now().year+1)
top100pos = pandas.DataFrame(index=positions.index, columns=list(yr_range))

# isolate headers and search for digits to use in position counter
headers = list(top100names.columns)
p_range = set([int((d.group(1))) for h in headers for d in (re.compile('(\d+)').search(h),) if d])


workbook = xlsxwriter.Workbook('merge1.xlsx')
worksheet = workbook.add_worksheet()

# indexmerged = list(top100merged.index)
# columnsmerged = list(top100merged.columns)

fontsize = 6

afam_format = workbook.add_format({
	'font_name': 'Arial',
	'text_wrap': 1,
	'bold': 1,
	'border': 1,
	'align': 'center',
	'valign': 'vcenter',
	'fg_color': '#F3A135',
	'font_size': fontsize})
	
asam_format = workbook.add_format({
	'font_name': 'Arial',
	'text_wrap': 1,
	'bold': 1,
	'border': 1,
	'align': 'center',
	'valign': 'vcenter',
	'fg_color': '#D54A99',
	'font_size': fontsize})

white_format = workbook.add_format({
	'font_name': 'Arial',
	'text_wrap': 1,
	'bold': 1,
	'border': 1,
	'align': 'center',
	'valign': 'vcenter',
	'fg_color': '#4A65AE',
	'font_size': fontsize})

jewish_format = workbook.add_format({
	'font_name': 'Arial',
	'text_wrap': 1,
	'bold': 1,
	'border': 1,
	'align': 'center',
	'valign': 'vcenter',
	'fg_color': '#A6A9AB',
	'font_size': fontsize})

latino_format = workbook.add_format({
	'font_name': 'Arial',
	'text_wrap': 1,
	'bold': 1,
	'border': 1,
	'align': 'center',
	'valign': 'vcenter',
	'fg_color': '#8BC543',
	'font_size': fontsize})	

header_format = workbook.add_format({
	'font_name': 'Arial',
	'text_wrap': 1,
	'bold': 1,
	'border': 1,
	'align': 'left',
	'valign': 'vcenter',
	'fg_color': '#8BC543',
	'font_size': 12})		
	
# for p in p_range: 
	# p_n = 'p_0'+str(p) 
	# s_n = 's_0'+str(p) 
	# e_n = 'e_0'+str(p) 
	# top100namesn = top100names[top100names[p_n].notnull()] 
	# for names in top100namesn.itertuples(): 
		# pos_name = top100namesn.loc[names[0], p_n] 
		# attribs = names[1] 
		# gender = names[2]
		# race = names[3]
		# status = names[4]
		# strtyr = int(top100namesn.loc[names[0], s_n]) 
		# endyr = int(top100namesn.loc[names[0], e_n]) 
		# strtcll = str(columns.loc[strtyr, 'A'])+str(rows.loc[pos_name, 1])
		# endcll = str(columns.loc[endyr, 'A'])+str(rows.loc[pos_name, 1])
		# rangecll = str(strtcll)+':'+str(endcll) 
		# if race == 'Black':
			# merge_format = afam_format
		# elif race == 'Asian':
			# merge_format = asam_format
		# elif race == 'Jewish':
			# merge_format = jewish_format
		# elif race == 'White':
			# merge_format = white_format
		# elif race == 'Latino':
			# merge_format = latino_format		
		# if strtcll == endcll:
			# # fontsize = int(6)
			# worksheet.write(strtcll, attribs, merge_format)
		# elif strtcll != endcll:
			# fontsize = 14
			# worksheet.merge_range(str(rangecll), attribs, merge_format)

indexlist = list(top100pos.index)
columnslist = list (top100pos.columns)		
			
for p in p_range: 
	p_n = 'p_0'+str(p) 
	s_n = 's_0'+str(p) 
	e_n = 'e_0'+str(p) 
	top100namesn = top100names[top100names[p_n].notnull()] 
	for names in top100namesn.itertuples(): 
		pos_name = top100namesn.loc[names[0], p_n] 
		attribs = names[1] 
		gender = names[2]
		race = names[3]
		status = names[4]
		strtyr = int(top100namesn.loc[names[0], s_n]) 
		endyr = int(top100namesn.loc[names[0], e_n]) 
		strtcol = columnslist.index(strtyr) 
		endcol = columnslist.index(endyr)
		posrow = indexlist.index(pos_name)
		if race == 'Black':
			merge_format = afam_format
		elif race == 'Asian':
			merge_format = asam_format
		elif race == 'Jewish':
			merge_format = jewish_format
		elif race == 'White':
			merge_format = white_format
		elif race == 'Latino':
			merge_format = latino_format
		if strtcol == endcol:
			# fontsize = int(6)
			worksheet.write(posrow, strtcol, attribs, merge_format)
		elif strtcol != endcol:
			# fontsize = 14
			worksheet.merge_range(posrow, strtcol, posrow, endcol, attribs, merge_format)

workbook.close()
