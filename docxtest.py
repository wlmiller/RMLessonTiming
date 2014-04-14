from docx import *
import zipfile, lxml.etree
import os
import re

mainlinestyles = ['Line','Normal','DefaultStyle']
for lesson in ['001']:#os.listdir('scripts'):
	for fn in ['001-150.docx']:#[f for f in os.listdir('scripts/' + lesson) if '.docx' in f]:
		doc = Document('scripts/' + lesson + '/' + fn)
		wc = 0
		lc = 0
		submittime = 0
		wtdc = 0
		nextc = 0
		inNR = False
		for par in doc.paragraphs:
			text = par.text
			style = par.style
			if style == 'NoResponse' or style == 'SecondaryNoResponse': inNR = True
			elif style in mainlinestyles: inNR = False
			print style,
			print text.encode('ascii','ignore'),
			print inNR
			if re.search('[0-9]+ words',text):
				match = re.search('[0-9]+ words',text).group()
				wc += int(match.split()[0])

			if style in mainlinestyles or inNR:
				if 'submit' in text.lower():
					if re.search('[0-9]+:[0-9][0-9]',text):
						time = re.search('[0-9]+:[0-9][0-9]',text).group(0)
						time = time.split(':')
						submittime += int(time[0])*60+int(time[1])	
					elif re.search('[0-9]+ second',text.lower()):
						time = re.search('[0-9]+ second',text.lower()).group(0)
						submittime += int(time.split(' ')[0])
					elif re.search('[0-9]+ minute',text.lower()):
						time = re.search('[0-9]+ minute',text.lower()).group(0)
						submittime += int(time.split(' ')[0])*60
					elif 'long' in text.lower():
						submittime += 40
					elif 'medium' in text.lower(): 
						submittime += 20
					elif 'short' in text.lower():
						submittime += 10
				if 'wtd' in text.lower() and not 'disappears' in text.lower():
					wtdc += 1
				if '[next' in text.lower():
					nextc += 1

		print fn.split('.')[0] + '\t' + str(wc) + '\t' + str(submittime) + '\t' + str(wtdc) + '\t' + str(nextc)
