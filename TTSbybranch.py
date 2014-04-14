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

mainlinestyles = ['Line','Normal','DefaultStyle']
for lesson in os.listdir('scripts'):
	for fn in [f for f in os.listdir('scripts/' + lesson) if 'docx' in f]:
		wavfn = 'scripts/' + lesson + '/' + fn.replace('.docx','')
		doc = Document('scripts/' + lesson + '/' + fn)
		MLtext = ''
		NRtext = ''
		inNR = False
		for par in doc.paragraphs:
			text = par.text
			style = par.style
			if style == 'NoResponse': inNR = True
			elif style in mainlinestyles: inNR = False
			if style in mainlinestyles or inNR:
				text = text.replace(u'\u2019',"'")
				text = text.encode('ascii','ignore')
				if style in ['Line','BranchLine']:
				#re.match('[A-Z][0-9]+',text) and not '(copy of' in text:
					text = re.sub('[A-Z][0-9,]+','',text)
					
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
					
					if style == 'Line':
						MLtext += ' ' + text
					elif inNR:
						NRtext += ' ' + text
					
		print fn.split('.')[0] + '\t' + str(getLength(MLtext,wavfn + '-main.wav')) + '\t' + str(getLength(NRtext,wavfn + '-NR.wav'))
		sys.stdout.flush()
