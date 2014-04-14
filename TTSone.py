from docx import *
import re
import sys,os,time
from comtypes.client import CreateObject
import comtypes.gen
import wave, contextlib


def getLength(text,wavfn):
	engine = CreateObject("SAPI.SpVoice")
	stream = CreateObject("SAPI.SpFileStream")
	stream.Open(wavfn, comtypes.gen.SpeechLib.SSFMCreateForWrite)
	engine.AudioOutputStream = stream
	engine.speak(text)
	stream.Close()
	
	with contextlib.closing(wave.open(wavfn,'r')) as f:
   		frames = f.getnframes()
    	rate = f.getframerate()
    	duration = frames / float(rate)
    
	return duration

for lesson in ['035']:#os.listdir('scripts'):
	for fn in ['035-010.docx','035-020.docx']:#[f for f in os.listdir('scripts/' + lesson) if 'docx' in f]:
		wavfn = 'scripts/' + lesson + '/' + fn.replace('docx','wav')
		doc = Document('scripts/' + lesson + '/' + fn)
		doctext = ''
		for par in doc.paragraphs:
			text = par.text
			text = text.replace(u'\u2019',"'")
			text = text.encode('ascii','ignore')
			if re.match('[A-Z][0-9]+',text):
				text = re.sub('[A-Z][0-9,]+','',text).split('//')[0]
				
				bc = 0
				temp = ''
				for char in text:	# Regexes don't handle nested brackets well.
					if char == '[':
						bc += 1
					if bc == 0: temp += char
					if char == ']':
						bc -= 1
				text = temp

				text = text.replace('  ',' ')
				text = text.replace('#','')
				
				doctext += ' ' + text
					
		print fn.split('.')[0] + '\t' + str(getLength(doctext,wavfn))
		sys.stdout.flush()
