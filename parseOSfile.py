from docx import *
import re

def parseOSfile(osfn):
	osfile = Document(osfn)

	paths = {'weak': [], 'average': [], 'strong': []}
	for par in osfile.paragraphs:
		if re.match('[0-9][0-9][0-9]?\. ',par.text):
			itemno = par.text.split('.')[0].zfill(3)
			for path in paths.keys():
				paths[path].append(itemno)

	osfile = Document(osfn)
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
			except Exception as e: 
				print e
				pass

	return paths
