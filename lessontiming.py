from docx import *
import re
import sys,os,time
from comtypes.client import CreateObject
import comtypes.gen
import wave, contextlib
from Tkinter import Tk
from tkFileDialog import askopenfilename


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

osfile = Document(filename)

paths = {'weak': [], 'average': [], 'strong': []}
for par in osfile.paragraphs:
	if re.match('[0-9][0-9][0-9]?\. ',par.text):
		itemno = par.text.split('.')[0].zfill(3)
		for path in paths.keys():
			paths[path].append(itemno)

for tab in osfile.tables:
	defaultitemno = tab.columns[0].cells[1].paragraphs[0].text.split('.')[0].zfill(3)
	for col in tab.columns:
		try:
			colheader = col.cells[0].paragraphs[0].text
			if re.match('[0-9][0-9][0-9]?\. ',col.cells[1].paragraphs[0].text):
				itemno = col.cells[1].paragraphs[0].text.split('.')[0].zfill(3)
			else:
				itemno = defaultitemno
			if 'weak' in colheader.lower():
				paths['weak'].append(itemno)
			if 'average' in colheader.lower():
				paths['average'].append(itemno)
			if 'strong' in colheader.lower():
				paths['strong'].append(itemno)
		except IndexError:
			pass
		except: 
			print e
			pass

for path in paths:
	print path + ':',
	print ', '.join(sorted(paths[path]))

allitems = []
for path in paths:
	allitems += paths[path]

allitems = set([i.encode('ascii') for i in allitems])

itemtimes = {}
for item in sorted(allitems):
	itemtimes[item] = 0.
	itemfile = filepath + 'Scripts/' + lesson + '-' + item + '.docx'
	
	if os.path.exists(itemfile.replace('docx','doc')) and not os.path.exists(itemfile):
		print 'Warning: script for item ' + item + ' is in *.doc format, not *.docx; skipping.'
	elif not os.path.exists(itemfile):
		print 'Warning: Scripts/' + lesson + '-' + item + '.docx not found!'
	else:
		itemtimes[item] = 1.

for item in allitems:
	print item + ':',
	print itemtimes[item]
