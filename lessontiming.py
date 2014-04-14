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

for path in paths:
	print path + ':',
	print ', '.join(sorted(paths[path]))

allitems = []
for path in paths:
	allitems += paths[path]

allitems = set([i.encode('ascii') for i in allitems])

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
stats = ['word count', 'submit time', 'WTD count', 'next count','dialogue time (total)', 'dialogue time (main branch)', 'dialogue time (NR branch)']

def predLength(itemstat):
	prediction = 21.565
	prediction += 0. * itemstat['word count']
	prediction += 0.269 * itemstat['submit time']
	prediction += 26.751 * itemstat['WTD count']
	prediction += 5.238 * itemstat['next count']
	prediction += 0.759 * itemstat['dialogue time (total)']
	prediction += 0.602 * itemstat['dialogue time (main branch)']
	prediction += 0. * itemstat['dialogue time (NR branch)']

	return prediction

for item in sorted(allitems):
	itemstats[item] = emptylesson
	itemfile = filepath + 'Scripts/' + lesson + '-' + item + '.docx'
	
	if os.path.exists(itemfile.replace('docx','doc')) and not os.path.exists(itemfile):
		print 'Warning: script for item ' + item + ' is in *.doc format, not *.docx; skipping.'
	elif not os.path.exists(itemfile):
		print 'Warning: Scripts/' + lesson + '-' + item + '.docx not found!'
	else:
		itemstats[item] = getlessonitemstats(itemfile)
	print item + ':'
	for stat in stats:
		print '\t{0:30s} {1:4d}'.format(stat,int(itemstats[item][stat]))
	print '\t' + '='*35
	print '\t{0:30s} {1:4d}'.format('PREDICTED TIME',int(predLength(itemstats[item])))

print ''

for path in ['weak + behind','weak + ontime']:
	print path + ':',
	print '{0:.2f}'.format(sum([predLength(itemstats[i]) for i in paths[path]])/60.),
	print 'min.'

sys.stdin.readline()
