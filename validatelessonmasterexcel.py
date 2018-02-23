# -*- coding: utf-8 -*-
#!/usr/bin/python

import sys
import xlrd  # sudo pip install xlrd
import os.path
from ireadydbmodule import *
#import pdb

EXCEL_FILENAME = "/Users/bnelson/Downloads/Release8-6_NewLesson Master_02_22_18_E.xlsx"

FILE_KEY_1 = 'Lesson'
FILE_KEY_2 = 'Master'
SPACE_CHAR = ' '

REQUIRED_SUBSTRING = FILE_KEY_1 + SPACE_CHAR + FILE_KEY_2

# Sheet 1 values:
SUBJECT_VALUES = [ 'Reading', 'Math' ]
GRADE_VALUES = [ 0, 1, 2, 3, 4, 5, 6, 7, 8 ]
YEAR_LEVEL_VALUES = [ 'Early', 'Mid', 'Late' ]
LESSON_STATE_DEFAULT = 'Disabled'
LESSON_TYPE_DEFAULT = 'HTML_LESSON'
SEQUENCE_PREFIXES = ['Before', 'After']
PHOENIX_ID_SUFFIX = '.phx'
# Phoenix domains that could potentially have exceptions to the rule that the domain is
# a perfect substring of the start of the lesson id.  Here, the domain is 'DI.MATH.GEO',
# but there are 5 old lessons with a lesson id that starts 'DI.MATH.GE'.
DOMAIN_EXCEPTIONS = [ 'DI.MATH.GEO' ]

# Sheet 2 values:
PLAYER_LINKS = [ '/instruction/math/', '/instruction/reading-comp/', '/instruction/phoenix/' ]
SWF_FILE_PREFIX = '#/lesson/'
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
	if isPhoenixLesson(lessonId) and domain in DOMAIN_EXCEPTIONS:
		return True
	
	return False

# Generates a Phoenix SWF file name from the lesson id.
# Steps:
# 1 - Strips off ".phx"
# 2 - Replaces all periods with underscores
# 3a - If a math lesson, adds in the Phoenix math SWF file prefix string
# 3b - If a reading lesson, adds in the Phoenix reading SWF file prefix string
def getPhoenixSwfFileName(lessonId):
	if not isPhoenixLesson(lessonId):
		return

	newSwfFileName = ''
	# Since we know it's a phoenix lesson ending in '.phx', strip off that suffix.
	newLessonId = lessonId[:-len(PHOENIX_ID_SUFFIX)]
	newLessonId = newLessonId.replace('.', '_')
	
	if PHOENIX_MATH_DOMAIN_TOKEN in lessonId:
		newSwfFileName = PHOENIX_MATH_SWF_FILE_PREFIX + newLessonId
	elif PHOENIX_ELA_DOMAIN_TOKEN in lessonId:
		newSwfFileName = PHOENIX_ELA_SWF_FILE_PREFIX + newLessonId

	return newSwfFileName
	
# Initialize some variables.

row_num = 1
problems = 0
warnings = 0
lesson_ids_from_spreadsheet_array = []
lesson_ids_array = []
sequence_array = []

# Examine just the file name without the path.
file_only = os.path.basename(EXCEL_FILENAME)

# If there's no space, we don't have to do any more checking.
if ' ' not in file_only:
	print "There's no space in the filename, there HAS to be a space between " + FILE_KEY_1 + " and " + FILE_KEY_2 + "."
	print "Here's the filename found:  >" + file_only + "<"
	sys.exit()

# Ensure the string "Lesson Master" (exactly like that) is in the file name somewhere.
if REQUIRED_SUBSTRING not in file_only:
	print "The substring of '" + REQUIRED_SUBSTRING + "' must exist in the filename, but doesn't; this file won't import: " + file_only
	sys.exit()

# Validate first sheet (main sheet).

# Oh yeah, and the path/file ought to exist!
if not os.path.exists(EXCEL_FILENAME):
	print "Unable to open file named '" + EXCEL_FILENAME + "', please check the name and try again."
	sys.exit()


# Crack open the workbook.
workbook = xlrd.open_workbook(EXCEL_FILENAME, encoding_override="cp1252")

# There better be 2 sheets inside!
if workbook.nsheets != 2:
	print "There must be exactly 2 sheets in the Excel spreadsheet."
	sys.exit()

# In case the sheet name changes, use indices; sheet 1 is the Lesson sheet.
#worksheet1 = workbook.sheet_by_name('New Lessons')
worksheet1 = workbook.sheet_by_index(0)

for row in range(1, worksheet1.nrows):
	lesson_id, domain, subject, lesson_name, obj_text, grade, year_level, new_domain_order, ed_notes, lesson_state, extra_only, sequence, lesson_type, corr_source_lesson = worksheet1.row_values(row)

	row_num += 1

	# Save lesson id and sequence so later we can validate the sequencing.
	lesson_ids_array.append(lesson_id)
	lesson_ids_from_spreadsheet_array.append(lesson_id)
	
	if SPACE_CHAR not in sequence:
		problems += 1
		print "Row " + str(row_num) + " has a sequence with no space, sequence text = " + sequence
	else:
		sequence_array.append([sequence, row_num])

	# Strip whitespace at right end of string for comparison below.
	lesson_id2 = lesson_id.rstrip()
	lesson_name2 = lesson_name.rstrip()
	
	# Validate that there aren't any spaces at the end of the id or name/title.
	if lesson_name != lesson_name2 or lesson_id != lesson_id2:
		problems += 1
		print "Row " + str(row_num) + " has a space at end of lesson name with id = " + lesson_id + " and name = >" + lesson_name + "<."
		
	if '_' in lesson_id:
		problems += 1
		print "Row " + str(row_num) + " has an underscore in lesson id with id = " + lesson_id
	
	if len(lesson_name) > 255:
		problems += 1
		print "Row " + str(row_num) + " has a lesson name with length greater than 255 with lesson id = " + lesson_id

	if len(obj_text) > 512:
		problems += 1
		print "Row " + str(row_num) + " has an objective text with length greater than 512 with lesson id = " + lesson_id
	
	domainLen = len(domain)

	# There's a few Phoenix lessons that don't fit this rule, ignore them with warnings.
	if isDomainMismatchWarning(lesson_id, domain) and lesson_id[0 : domainLen] != domain:
		warnings += 1
		print "Row " + str(row_num) + " appears to have an incorrect domain in row with id = " + lesson_id + " and domain = " + domain + "."
	# Validate that the domain is a perfect substr of the lesson id (it should be).	
	elif lesson_id[0 : domainLen] != domain:
		problems += 1
		print "Row " + str(row_num) + " has incorrect domain in row with id = " + lesson_id + " and domain = " + domain + "."

	# Validate the subject value (Reading or Math).		
	if subject not in SUBJECT_VALUES:
		problems += 1
		print "Row " + str(row_num) + " with lesson id " + lesson_id + " has bad subject value [must be 'Reading' or 'Math']: " + subject
		
	# Validate the grade is a number (0-8).
	if int(grade) not in GRADE_VALUES:
		problems += 1
		print "Row " + str(row_num) + " with lesson id " + lesson_id + " has bad grade value [must be a number]: " + grade

	# Validate that the year_level is one of the predefined values.
	if year_level not in YEAR_LEVEL_VALUES:
		problems += 1
		print "Row " + str(row_num) + " with lesson id " + lesson_id + " has bad year level value: " + year_level

	# Validate the lesson state.	
	if lesson_state != LESSON_STATE_DEFAULT:
		problems += 1
		print "Row " + str(row_num) + " with lesson id " + lesson_id + " has bad lesson state (should be " + LESSON_STATE_DEFAULT + ") : " + lesson_state	

	# Validate the lesson type.
	if lesson_type != LESSON_TYPE_DEFAULT:
		problems += 1
		print "Row " + str(row_num) + " with lesson id " + lesson_id + " has bad lesson type (should be " + LESSON_TYPE_DEFAULT + ") : " + lesson_type

# Show results.
if warnings > 0:
	print "There were " + str(warnings) + " warnings found, please check carefully and proceed if they're okay."
if problems > 0:
	print "There were " + str(problems) + " bad rows found, please fix the Lesson Master spreadsheet and rerun script to validate sequencing."
	sys.exit()

print "Spreadsheet columns validated, now validating the lesson id's sequencing...."

# Read in the existing lesson id's from the db.
# Get the connection for our local iready db.
conn = getiReadyDbConnection()

rows = getRowsFromDbForQuery(conn, 'SELECT id FROM lesson.iric_lesson')

seq_probs = 0

# Now add all the lesson id's from the db into our list.
for row in rows:
	#print "Processing db row = " + row[0]
	lesson_ids_array.append(row[0])


# Now go through all the sequences and ensure that all lesson id's referenced actually exist.
for sequenceObj in sequence_array:
	sequence = sequenceObj[0]
	row_num = sequenceObj[1]
	# The string is "Before <lesson id>" or "After <lesson id>".
	seq_tokens = sequence.split(' ')
	pref_text = seq_tokens[0]
	existing_lesson_id = seq_tokens[1]
	
	if pref_text not in SEQUENCE_PREFIXES:
		seq_probs += 1
		print "Prefix text in Sequence column must be 'Before' or 'After', found: " + pref_text + " in line " + str(row_num)
	
	if existing_lesson_id not in lesson_ids_array:
		seq_probs += 1
		print "Supposedly existing lesson id " + existing_lesson_id + " not found from sequence " + sequence + " in line " + str(row_num)

if seq_probs > 0:
	print "There were " + str(seq_probs) + " sequence issues found, please correct and try again."
	sys.exit()

# Validate sheet 2.

print "Validating 2nd sheet (components)...."

# Now start on sheet 2 (Components).
#worksheet1 = workbook.sheet_by_name('Components')
worksheet2 = workbook.sheet_by_index(1)

row_num = 1
problems2 = 0
new_lesson_ids_array = []

for row in range(1, worksheet2.nrows):
	lesson_id, player_link, swf_file_name, concat_url_swf, component_type, component_order, estimated_time = worksheet2.row_values(row)
	
	# Skip header row
	if component_type == "Component type":
		continue
		
	row_num += 1

	# Validate that the lesson id in this tab is also present in the first tab.
	if lesson_id not in lesson_ids_from_spreadsheet_array:
		problems2 += 1
		print "Row " + str(row_num) + " has a lesson id '" + lesson_id + "' from Components Tab that was not referenced in first tab!"

	# Save this lesson id so we can do the opposite validation below.	
	new_lesson_ids_array.append(lesson_id)
	
	# Check the player link value.
	if player_link not in PLAYER_LINKS:
		problems2 += 1
		print 'Row ' + str(row_num) + ' has an invalid player link: ' + player_link

	# Check the SWF file name value.  Phoenix lessons of course have to be different!  They use underscores in the SWF file name.
	if isPhoenixLesson(lesson_id):
		phxSwfFileName = getPhoenixSwfFileName(lesson_id)

		# With other lesson types, we can accurately predict what the SWF file name should be;
		# but not with Phoenix.  The best we can do is generate an approximation and see if
		# that approximation is in the real SWF file name.
		if phxSwfFileName not in swf_file_name:
			problems2 += 1
			print 'Row ' + str(row_num) + ' has an invalid Phoenix SWF file name: ' + swf_file_name + ' should contain: ' + phxSwfFileName
	elif SWF_FILE_PREFIX + lesson_id != swf_file_name:
		problems2 += 1
		print 'Row ' + str(row_num) + ' has an invalid SWF file name: ' + swf_file_name + ' should be: ' + SWF_FILE_PREFIX + lesson_id

	# Check the Concat. URL SWF value.		
	if concat_url_swf != '':
		problems2 += 1
		print 'Row ' + str(row_num) + ' has an invalid concat URL and SWF, MUST be empty: ' + concat_url_swf

	# Check the component type value.
	if component_type not in COMPONENTS:
		problems2 += 1
		print 'Row ' + str(row_num) + ' has an invalid component type: ' + component_type

	#pdb.set_trace()

	# First check if the column was empty or all whitespace.
	if not isinstance(component_order, float) and (len(component_order) == 0 or component_order.isspace()):
		problems2 += 1
		print 'Row ' + str(row_num) + ' has an empty component order, MUST be a number.'
	else:
		# Check the component order value.	Excel always reads in integers as floats,
		# so convert to int() to get rid of '.0', and then convert to string to check
		# if the result is an integer number.
		component_order = int(component_order)
		comp_order_str = str(component_order)	
		if not comp_order_str.isdigit():
			problems2 += 1
			print 'Row ' + str(row_num) + ' has an invalid component order, MUST be a number: ' + comp_order_str

	# First check if the column was empty or all whitespace.
	if not isinstance(estimated_time, float) and (len(estimated_time) == 0 or estimated_time.isspace()):
		problems2 += 1
		print 'Row ' + str(row_num) + ' has an empty estimated time, MUST be a number.'
	else:
		# Check the estimated time value. Excel always reads in integers as floats,
		# so convert to int() to get rid of '.0', and then convert to string to check
		# if the result is an integer number.
		estimated_time = int(estimated_time)
		est_time_str = str(estimated_time)
		if not est_time_str.isdigit():
			problems2 += 1
			print 'Row ' + str(row_num) + ' has an invalid estimated time, MUST be a number (and NOT blank): ' + est_time_str
		

# Lesson id's from the first tab must appear at least once in this tab; validate that here.
for lessonid in lesson_ids_from_spreadsheet_array:
	if lessonid not in new_lesson_ids_array:
		problems2 += 1
		print "Lesson id '" + lessonid + "' from first tab was not referenced in Components Tab!"

if problems2 > 0:
	print 'There were ' + str(problems2) + ' problems in the Components Sheet, please fix and retry.'
else:
	print "Lesson Master spreadsheet (" + EXCEL_FILENAME + ") validated!"

# exit the program
sys.exit()
