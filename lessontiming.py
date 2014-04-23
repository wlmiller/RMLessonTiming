from docx import *
import re
import sys,os,time
from comtypes.client import CreateObject
import comtypes.gen
import wave, contextlib
from Tkinter import Tk
from tkFileDialog import askopenfilename
from parseOSfile import parseOSfile
from lessonitemstats import getlessonitemstats

Tk().withdraw()
if len(sys.argv) > 1:
	filename = sys.argv[1]
else:
	filename = askopenfilename(**{'title':'Select the OS file'})

if not filename[-4:] == 'docx':
	try:
		raise Exception()
	except Exception as e:
		print 'OS file must be of type *.docx' 
		exit(3)
lesson = re.search('[0-9][0-9][0-9]',filename).group()
filepath = filename.replace(lesson + '.docx','')

paths = parseOSfile(filename)

allitems = []
for fn in [f for f in os.listdir(filepath + 'Scripts/') if 'docx' in f]:
	allitems += [fn.split('.doc')[0].encode('ascii')]

itemstats = {}
emptylesson = {
	'word count': 0,	# Word count
	'submit time': 0, 	# Total submit time
	'WTD count': 0, 	# "Write this down" count
	'next count': 0,	# "Next" count
	'dialogue time (total)': 0., 		# Total dialogue time estimate
	'dialogue time (main branch)': 0., 	# Main branch dialogue time estimate
	'dialogue time (NR branch)': 0., 	# NoResponse branch dialogue time estimate
	}

def predLessonLength(itemstats):
	return '45:00'

def predItemLength(itemstat):
	prediction = 21.565
	prediction += 0. * itemstat['word count']
	prediction += 0.269 * itemstat['submit time']
	prediction += 26.751 * itemstat['WTD count']
	prediction += 5.238 * itemstat['next count']
	prediction += 0.759 * itemstat['dialogue time (total)']
	prediction += 0.602 * itemstat['dialogue time (main branch)']
	prediction += 0. * itemstat['dialogue time (NR branch)']

	minutes = int(prediction/60)
	seconds = int(round(prediction-minutes*60))
	return str(minutes) + ':' + str(seconds)

print ''
for i in sorted(allitems):
	item = '-'.join(i.split('-')[1:])
	itemstats[item] = emptylesson
	itemfile = filepath + 'Scripts/' + i + '.docx'
	
	if os.path.exists(itemfile.replace('docx','doc')) and not os.path.exists(itemfile):
		print 'Warning: script for item ' + item + ' is in *.doc format, not *.docx; skipping.'
	elif not os.path.exists(itemfile):
		print '"' + lesson + '"'
		print '"' + item + '"'
		print itemfile
		print 'Warning: Scripts/' + lesson + '-' + item + '.docx not found!'
	else:
		itemstats[item] = getlessonitemstats(itemfile)
		print i.ljust(15) + predItemLength(itemstats[item]).rjust(10)

print '='*25

for path in ['weak + behind','weak + ontime']:
	pathstats = [itemstats[i] for i in paths[path]]
	print path.ljust(15) + predLessonLength(pathstats).rjust(10)

sys.stdin.readline()
