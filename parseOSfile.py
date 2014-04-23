from docx import *
import re

def parseOSfile(osfn):
	osfile = Document(osfn)

	paths = {'weak + behind': [], 'weak + ontime': []}
	for par in osfile.paragraphs:
		if re.match(' ?[0-9][0-9][0-9]?\.',par.text):
			itemno = par.text.split('.')[0].replace(' ','').zfill(3)
			for path in paths.keys():
				paths[path].append(itemno.encode('ascii'))

	osfile = Document(osfn)
	for tab in osfile.tables:
		if len(tab.columns[0].cells) > 1:
			defaultitemno = tab.columns[0].cells[1].paragraphs[0].text.split('.')[0].replace(' ','').zfill(3)
			for col in tab.columns:
				try:
					colheader = col.cells[0].paragraphs[0].text
					for par in col.cells[1].paragraphs:
						if re.match(' ?[0-9][0-9][0-9]?\.',par.text):
							itemno = par.text.split('.')[0].replace(' ','').zfill(3)
						elif re.match('^same',par.text.lower()):
							itemno = defaultitemno
						else: itemno = ''
						if not itemno == '':
							if 'weak' in colheader.lower() or (not 'average' in colheader.lower() and not 'strong' in colheader.lower()):
								if not ('skip if behind' in colheader.lower() or 'not behind' in colheader.lower() or colheader.lower == 'behind'):
									paths['weak + behind'].append(itemno.encode('ascii'))
								if not colheader.lower() == 'behind':
									paths['weak + ontime'].append(itemno.encode('ascii'))
				except IndexError:
					pass
				except Exception as e: 
					print e
					pass

	return paths
