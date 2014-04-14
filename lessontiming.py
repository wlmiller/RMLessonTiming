from docx import *
import re
import sys,os,time
from comtypes.client import CreateObject
import comtypes.gen
import wave, contextlib
from Tkinter import Tk
from tkFileDialog import askopenfilename
from parseOSfile import parseOSfile


Tk().withdraw()
if len(sys.argv) > 1:
	filename = sys.argv[1]
else:
	filename = askopenfilename()

if not filename[-4:] == 'docx':
	try:
		raise Exception()
	except Exception as e:
		print 'OS file must be of type *.docx' 
		exit(3)
lesson = re.search('[0-9][0-9][0-9]',filename).group()
filepath = filename.replace(lesson + '.docx','')
print filepath
print lesson

paths = parseOSfile(filename)

for path in paths:
	print path + ':',
	print ', '.join(sorted(paths[path]))

allitems = []
for path in paths:
	allitems += paths[path]

allitems = set([i.encode('ascii') for i in allitems])

itemstats = {}
emptylesson = {
	'WC': 0, 			# Word count
	'submittime': 0, 	# Total submit time
	'WTDs': 0, 			# "Write this down" count
	'nexts': 0, 		# "Next" count
	'TTStime': 0., 		# Total dialogue time estimate
	'TTSmaintime': 0., 	# Main branch dialogue time estimate
	'TTSNRtime': 0., 	# NoResponse branch dialogue time estimate
	'timeestimate': 0.}	# Total predicted time

for item in sorted(allitems):
	itemstats[item] = emptylesson
	itemfile = filepath + 'Scripts/' + lesson + '-' + item + '.docx'
	
	if os.path.exists(itemfile.replace('docx','doc')) and not os.path.exists(itemfile):
		print 'Warning: script for item ' + item + ' is in *.doc format, not *.docx; skipping.'
	elif not os.path.exists(itemfile):
		print 'Warning: Scripts/' + lesson + '-' + item + '.docx not found!'
	else:
		itemstats[item] = getlessonstats(itemfile)

for item in allitems:
	print item + ':',
	print itemstats[item]
