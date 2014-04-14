from docx import *
import os
import re
from comtypes.client import CreateObject
import comtypes.gen
import wave, contextlib

mainlinestyles = ['Line','Normal','DefaultStyle']

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

def getlessonitemstats(itemfn):
	doc = Document(itemfn)
	wavfn = itemfn.replace('docx','wav')

	wc = 0
	lc = 0
	submittime = 0
	wtdc = 0
	nextc = 0
	inNR = False
	doctext = ''
	MLtext = ''
	NRtext = ''

	for par in doc.paragraphs:
		text = par.text
		style = par.style
		if style == 'NoResponse' or style == 'SecondaryNoResponse': inNR = True
		elif style in mainlinestyles: inNR = False
		
		if re.search('[0-9]+ words',text):
			match = re.search('[0-9]+ words',text).group()
			wc += int(match.split()[0])

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

		if style in mainlinestyles or inNR:
			branchtext = text.replace(u'\u2019',"'")
			branchtext = branchtext.encode('ascii','ignore')
			if style in ['Line','BranchLine'] or inNR:
					
				branchtext = re.sub('[A-Z][0-9,]+','',text)
					
				bc = 0
				temp = ''
				for char in branchtext:	# Regexes don't handle nested brackets well.
					if char == '[':
						bc += 1
					if bc == 0: temp += char
					if char == ']':
						bc -= 1
				branchtext = temp

				branchtext = branchtext.replace('  ',' ')
				branchtext = branchtext.replace('#','')
						
				if style == 'Line':
					MLtext += ' ' + branchtext
				elif inNR:
					NRtext += ' ' + branchtext

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

	TTStime = getLength(doctext,wavfn)
	MLtime = getLength(MLtext,wavfn.replace('.wav','-main.wav'))
	NRtime = getLength(NRtext,wavfn.replace('.wav','-NR.wav'))
	return {
	'word count': wc,
	'submit time': submittime, 
	'WTD count': wtdc,
	'next count': nextc,
	'dialogue time (total)': TTStime, 
	'dialogue time (main branch)': MLtime,
	'dialogue time (NR branch)': NRtime
	}
