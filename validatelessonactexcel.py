# -*- coding: utf-8 -*-
#!/usr/bin/python

import sys
import xlrd  # sudo pip install xlrd
import os.path
from ireadydbmodule import *
import re
import argparse
#import pdb

#EXCEL_FILENAME = "/Users/bnelson/Downloads/MasterLessonSequence_Math_lesson_reorder_031618.xlsx"

# Set up and parse command line arguments
parser = argparse.ArgumentParser()
parser.add_argument("filename", help = "Excel filename of reorder sheet to validate")
parser.add_argument("--nowarn", help = "Do not print warnings, only reveal total warnings", action = "store_true")
args = parser.parse_args()

EXCEL_FILENAME = args.filename
SUPPRESS_WARNINGS = args.nowarn

if SUPPRESS_WARNINGS:
	print "NOTE: Warning messages will be suppressed."

FILE_KEY_1 = 'lesson'
FILE_KEY_2 = 'reorder'

NUM_COLUMNS = 19

PHOENIX_ID_SUFFIX = '.phx'


# Sheet 1 values:
SUBJECT_VALUES = [ 'Reading', 'Math' ]
GRADE_VALUES = [ 0, 1, 2, 3, 4, 5, 6, 7, 8 ]
YEAR_LEVEL_VALUES = [ 'Early', 'Mid', 'Late', 'Extra' ]
LESSON_STATE_VALUES = [ 'Disabled', 'Enabled' ]
LESSON_TYPE_VALUES = [ 'HTML_LESSON', 'FLASH_LESSON' ]
SEQUENCE_PREFIXES = ['Before', 'After']
DOMAIN_VALUES = [ 'DI.MATH.NO', 'DI.MATH.AL', 'DI.MATH.MS', 'DI.MATH.GEO', 'DI.ELA.PA', 'DI.ELA.PH', 'DI.ELA.HFW', 'DI.ELA.VOC', 'DI.ELA.COM', 'DI.ELA.INSTR.CR' ]
EXTRA_ONLY_VALUES = [ 'Yes', 'No' ]

# Other domains can have exceptions outside of Phoenix, it turns out.  Ugh.
DOMAIN_EXCEPTIONS = [ 'DI.MATH.AL', 'DI.MATH.MS', 'DI.MATH.GEO' ]

# Phoenix domains that could potentially have exceptions to the rule that the domain is
# a perfect substring of the start of the lesson id.  Here, the domain is 'DI.MATH.GEO',
# but there are 5 old lessons with a lesson id that starts 'DI.MATH.GE'.
PHOENIX_DOMAIN_EXCEPTIONS = [ 'DI.MATH.GEO', 'DI.MATH.AL' ]

PHOENIX_MATH_DOMAIN_TOKEN = 'MATH'
PHOENIX_ELA_DOMAIN_TOKEN = 'ELA'
PHOENIX_MATH_SWF_FILE_PREFIX = '#/lesson/math/'
PHOENIX_ELA_SWF_FILE_PREFIX = '#/lesson/reading/'
COMPONENTS = [ 'Practice', 'Tutorial', 'Quiz', 'Close Reading' ]

# Returns True if this is a Phoenix lesson, False otherwise.
# Does this by simply looking at the lesson id for ".phx".
def isPhoenixLesson(lessonId):
	if PHOENIX_ID_SUFFIX in lessonId:
		return True
	
	return False

# Return True if this Phoenix domain is one of those that are an exception to the
# mismatch error we normally produce; in other words, if the Phoenix lesson has
# one of the domains in the exceptions array, issue a warning (instead of an error)
# if the domain isn't a perfect substring of the lesson id.
def isDomainMismatchWarning(lessonId, domain):
	if isPhoenixLesson(lessonId) and domain in PHOENIX_DOMAIN_EXCEPTIONS:
		return True
	
	return False


# Initialize some variables.
row_num = 1
problems = 0
warnings = 0
lesson_ids_array_from_db = []
curr_domain = ""
domain_var = 0

print "Starting validations on " + EXCEL_FILENAME

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

#rows = getRowsFromDbForQuery(conn, 'select id,lesson_id from lesson.iric_lesson_component')
rows = getRowsFromDbForQuery(conn, 'select id from lesson.iric_lesson')

# Now add all the lesson id's from the db into our list.
for row in rows:
	id = row[0]
#	print "Processing db row where lesson id = >" + id + "<"
	lesson_ids_array_from_db.append(id)

# Put lesson ids into set for faster lookup
#lesson_ids_array_from_db_set = set(lesson_ids_array_from_db)

#for row in lesson_ids_array_from_db:
#	print "Found >" + row + "< in list."

# In case the sheet name changes, use indices.
#worksheet1 = workbook.sheet_by_name('New Lessons')
worksheet1 = workbook.sheet_by_index(0)

for row in range(1, worksheet1.nrows):
	all_vals = worksheet1.row_values(row)

	# Only fetch the first 19 columns; we had a sheet with 20 values, but the
	# last column was just empty stuff.	
	lesson_id, domain, subject, lesson_title, obj_text, grade, year_level, new_domain_order, lesson_state, extra_only, orig_lesson_id, all_cols_to_right, new_iready_domain_order, old_domain_order, old_grade, old_sy_level, notes, domain_sequence, lesson_type = all_vals[:NUM_COLUMNS]
	
	if not lesson_id and not domain and not subject:
		continue

	row_num += 1
	
	if domain != curr_domain:
#		print "switched domains where domain =  " + domain + ", curr_domain = " + curr_domain
		domain_var = 1
		curr_domain = domain
	else:
#		print "same domain where domain =  " + domain + ", curr_domain = " + curr_domain
		domain_var += 1
	
	if '_' in lesson_id:
		problems += 1
		print "ERROR: Row " + str(row_num) + " has an underscore in lesson id with lesson id = " + lesson_id
		
	# Ensure a new lesson id is already associated with a component in the db.
	if lesson_state == "Enabled" and lesson_type == "HTML_LESSON" and lesson_id not in lesson_ids_array_from_db:
		problems += 1
		print "ERROR: Row " + str(row_num) + " has a lesson id that does not appear in the db already with lesson id = " + lesson_id

	# If there's a ".v<some number>", then make sure there's an original lesson as well.
	ends_in_dotvnum = re.search(r"\.v[0-9]+", lesson_id)
	prev_lesson = lesson_id.rsplit('.', 1)[0]
	
	#pdb.set_trace()

	if ends_in_dotvnum == "" and (not orig_lesson_id or prev_lesson not in lesson_ids_array_from_db):
		problems += 1
		print "ERROR: Row " + str(row_num) + " has a lesson id with '.v<num> but no original lesson id with lesson id = " + prev_lesson

	if not ends_in_dotvnum and orig_lesson_id:
		problems += 1
		print "ERROR: Row " + str(row_num) + " has a lesson id without '.v<num> but an original lesson id is specified with lesson id = " + lesson_id
	
	if lesson_id == orig_lesson_id:
		problems += 1
		print "ERROR: Row " + str(row_num) + " has a lesson id equal to the original lesson id with lesson id = " + lesson_id

	if len(lesson_title) > 255:
		problems += 1
		print "ERROR: Row " + str(row_num) + " has a lesson name with length greater than 255 with lesson id = " + lesson_id

	if len(obj_text) > 512:
		problems += 1
		print "ERROR: Row " + str(row_num) + " has an objective text with length greater than 512 with lesson id = " + lesson_id
	
	domainLen = len(domain)

	# There's a few Phoenix lessons that don't fit this rule, ignore them with warnings.
	if isDomainMismatchWarning(lesson_id, domain) and lesson_id[0 : domainLen] != domain:
		warnings += 1
		if not SUPPRESS_WARNINGS:
			print "WARNING: Row " + str(row_num) + " has a mismatched Phoenix domain in row with id = " + lesson_id + " and domain = " + domain + "."
	elif domain in DOMAIN_EXCEPTIONS:
		warnings += 1
		if not SUPPRESS_WARNINGS:
			print "WARNING: Row " + str(row_num) + " has a mismatched domain in row with id = " + lesson_id + " and domain = " + domain + "."
	# Validate that the domain is a perfect substr of the lesson id (it should be).	
	elif lesson_id[0 : domainLen] != domain:
		problems += 1
		print "ERROR: Row " + str(row_num) + " has incorrect domain in row with id = " + lesson_id + " and domain = " + domain + "."

	# Validate that the domain is a perfect substr of the lesson id (it should be).	
#	if lesson_id[0 : domainLen] != domain:
#		problems += 1
#		print "Row " + str(row_num) + " has incorrect domain in row with id = " + lesson_id + " and domain = " + domain + "."

	# Validate that the domain is a known one.
	if domain not in DOMAIN_VALUES:
		problems += 1
		print "ERROR: Row " + str(row_num) + " has unknown domain in row with id = " + lesson_id + " and domain = " + domain + "."

	# Validate the subject value (Reading or Math).		
	if subject not in SUBJECT_VALUES:
		problems += 1
		print "ERROR: Row " + str(row_num) + " with lesson id " + lesson_id + " has bad subject value [must be 'Reading' or 'Math']: " + subject
		
	# Validate the grade is a number (0-8).
	if int(grade) not in GRADE_VALUES:
		problems += 1
		print "ERROR: Row " + str(row_num) + " with lesson id " + lesson_id + " has bad grade value [must be a number]: " + grade

	# Validate that the new domain order is starting at one and increasing monotonically, until we get to a new domain.
	if new_domain_order != domain_var:
		problems += 1
		print "ERROR: Row " + str(row_num) + " with lesson id " + lesson_id + " has bad new domain order: " + str(int(new_domain_order)) + " should be: " + str(domain_var)
	
	# Validate that the year_level is one of the predefined values.
	if year_level not in YEAR_LEVEL_VALUES:
		problems += 1
		print "ERROR: Row " + str(row_num) + " with lesson id " + lesson_id + " has bad year level value: " + year_level

	# Validate the lesson state.	
	if lesson_state not in LESSON_STATE_VALUES:
		problems += 1
		print "ERROR: Row " + str(row_num) + " with lesson id " + lesson_id + " has bad lesson state : " + lesson_state	

	# Validate the extra only column.
	if extra_only not in EXTRA_ONLY_VALUES:
		problems += 1
		print "ERROR: Row " + str(row_num) + " with lesson id " + lesson_id + " has bad extra only value : " + extra_only

	# Validate the lesson type.
	if lesson_type not in LESSON_TYPE_VALUES:
		problems += 1
		print "ERROR: Row " + str(row_num) + " with lesson id " + lesson_id + " has bad lesson type : " + lesson_type


print ""

# Show results.
if warnings > 0:
	print "There were " + str(warnings) + " warnings found, please check the messages carefully."
	
if problems > 0:
	print "There were " + str(problems) + " bad rows found, please fix the Lesson Activation spreadsheet."
else:
	print "Lesson Activation spreadsheet (" + EXCEL_FILENAME + ") validated!"

# exit the program
sys.exit()
