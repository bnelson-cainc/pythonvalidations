# -*- coding: utf-8 -*-
#!/usr/bin/python

import sys
import xlrd  # sudo pip install xlrd
import os.path
from ireadydbmodule import *
import re
#import pdb

EXCEL_FILENAME = "/Users/bnelson/Downloads/Comprehension_lesson_reorder_ALL_K2_01_19_18_v5.xlsx"

FILE_KEY_1 = 'lesson'
FILE_KEY_2 = 'reorder'


# Sheet 1 values:
SUBJECT_VALUES = [ 'Reading', 'Math' ]
GRADE_VALUES = [ 0, 1, 2, 3, 4, 5, 6, 7, 8 ]
YEAR_LEVEL_VALUES = [ 'Early', 'Mid', 'Late', 'Extra' ]
LESSON_STATE_VALUES = [ 'Disabled', 'Enabled' ]
LESSON_TYPE_VALUES = [ 'HTML_LESSON', 'FLASH_LESSON' ]
SEQUENCE_PREFIXES = ['Before', 'After']
DOMAIN_VALUES = [ 'DI.MATH.NO', 'DI.MATH.AL', 'DI.MATH.MS', 'DI.MATH.GEO', 'DI.ELA.PA', 'DI.ELA.PH', 'DI.ELA.HFW', 'DI.ELA.VOC', 'DI.ELA.COM', 'DI.ELA.INSTR.CR' ]
EXTRA_ONLY_VALUES = [ 'Yes', 'No' ]


row_num = 1
problems = 0
lesson_ids_array_from_db = []

file_only = os.path.basename(EXCEL_FILENAME)

# Validate the file name is correct.
if (FILE_KEY_1 + '_' + FILE_KEY_2) not in file_only:
	print "The spreadsheet filename MUST contain 'lesson_reorder in the filename."
	print "Here's the filename found:  >" + file_only + "<"
	sys.exit()

# Validate first tab (main tab).

if not os.path.exists(EXCEL_FILENAME):
	print "Unable to open file named '" + EXCEL_FILENAME + "', please check the name and try again."
	sys.exit()


workbook = xlrd.open_workbook(EXCEL_FILENAME, encoding_override="cp1252")

if workbook.nsheets != 1:
	print "There must be exactly 1 sheet in the Excel spreadsheet."
	sys.exit()

# Read in the existing lesson id's from the db.
# Get the connection for our local iready db.
conn = getiReadyDbConnection()

rows = getRowsFromDbForQuery(conn, 'select id,lesson_id from lesson.iric_lesson_component')

# Now add all the lesson id's from the db into our list.
for row in rows:
#	print "Processing db row = " + row[0] + "," + row[1]
	id,comp_lesson_id = row
	lesson_ids_array_from_db.append(comp_lesson_id)

# Put lesson ids into set for faster lookup
lesson_ids_array_from_db_set = set(lesson_ids_array_from_db)

# In case the sheet name changes, use indices.
#worksheet1 = workbook.sheet_by_name('New Lessons')
worksheet1 = workbook.sheet_by_index(0)

for row in range(1, worksheet1.nrows):
	lesson_id, domain, subject, lesson_title, obj_text, grade, year_level, new_domain_order, lesson_state, extra_only, orig_lesson_id, all_cols_to_right, new_iready_domain_order, old_domain_order, old_grade, old_sy_level, notes, domain_sequence, lesson_type = worksheet1.row_values(row)
	
	if not lesson_id and not domain and not subject:
		continue

	row_num += 1

	if '_' in lesson_id:
		problems += 1
		print "Row " + str(row_num) + " has an underscore in lesson id with lesson id = " + lesson_id
		
	# Ensure a new lesson id is already associated with a component in the db.
	if lesson_state == "Enabled" and lesson_type == "HTML_LESSON" and lesson_id not in lesson_ids_array_from_db_set:
		problems += 1
		print "Row " + str(row_num) + " has a lesson id that does not appear in the db already with lesson id = " + lesson_id

	ends_in_dotvnum = re.search(r"\.v[0-9]+", lesson_id)
	
	if ends_in_dotvnum and not orig_lesson_id:
		problems += 1
		print "Row " + str(row_num) + " has a lesson id with '.v<num> but no original lesson id with lesson id = " + lesson_id

	if not ends_in_dotvnum and orig_lesson_id:
		problems += 1
		print "Row " + str(row_num) + " has a lesson id without '.v<num> but an original lesson id is specified with lesson id = " + lesson_id
	
	if lesson_id == orig_lesson_id:
		problems += 1
		print "Row " + str(row_num) + " has a lesson id equal to the original lesson id with lesson id = " + lesson_id

	if len(lesson_title) > 255:
		problems += 1
		print "Row " + str(row_num) + " has a lesson name with length greater than 255 with lesson id = " + lesson_id

	if len(obj_text) > 512:
		problems += 1
		print "Row " + str(row_num) + " has an objective text with length greater than 512 with lesson id = " + lesson_id
	
	domainLen = len(domain)

	# Validate that the domain is a perfect substr of the lesson id (it should be).	
	if lesson_id[0 : domainLen] != domain:
		problems += 1
		print "Row " + str(row_num) + " has incorrect domain in row with id = " + lesson_id + " and domain = " + domain + "."

	# Validate that the domain is a known one.
	if domain not in DOMAIN_VALUES:
		problems += 1
		print "Row " + str(row_num) + " has unknown domain in row with id = " + lesson_id + " and domain = " + domain + "."

	# Validate the subject value (Reading or Math).		
	if subject not in SUBJECT_VALUES:
		problems += 1
		print "Row " + str(row_num) + " with lesson id " + lesson_id + " has bad subject value [must be 'Reading' or 'Math']: " + subject
		
	# Validate the grade is a number (0-8).
	if int(grade) not in GRADE_VALUES:
		problems += 1
		print "Row " + str(row_num) + " with lesson id " + lesson_id + " has bad grade value [must be a number]: " + grade

	# Validate that the new domain order is starting at one and increasing monotonically.
	if new_domain_order != row_num - 1:
		problems += 1
		print "Row " + str(row_num) + " with lesson id " + lesson_id + " has bad new domain order: " + new_domain_order
	
	# Validate that the year_level is one of the predefined values.
	if year_level not in YEAR_LEVEL_VALUES:
		problems += 1
		print "Row " + str(row_num) + " with lesson id " + lesson_id + " has bad year level value: " + year_level

	# Validate the lesson state.	
	if lesson_state not in LESSON_STATE_VALUES:
		problems += 1
		print "Row " + str(row_num) + " with lesson id " + lesson_id + " has bad lesson state : " + lesson_state	

	# Validate the extra only column.
	if extra_only not in EXTRA_ONLY_VALUES:
		problems += 1
		print "Row " + str(row_num) + " with lesson id " + lesson_id + " has bad extra only value : " + extra_only

	# Validate the lesson type.
	if lesson_type not in LESSON_TYPE_VALUES:
		problems += 1
		print "Row " + str(row_num) + " with lesson id " + lesson_id + " has bad lesson type : " + lesson_type


print ""

# Show results.
if problems > 0:
	print "There were " + str(problems) + " bad rows found, please fix the Lesson Activation spreadsheet."
else:
	print "Lesson Activation spreadsheet (" + EXCEL_FILENAME + ") validated!"

# exit the program
sys.exit()
