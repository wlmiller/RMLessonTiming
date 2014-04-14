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

for lesson in os.listdir('scripts'):
	for fn in [f for f in os.listdir('scripts/' + lesson) if 'docx' in f]:
		wavfn = 'scripts/' + lesson + '/' + fn.replace('docx','wav')
		doc = Document('scripts/' + lesson + '/' + fn)
		doctext = ''
		branchcount = 0
		inBranch = False
		for par in doc.paragraphs:
			if par.style == 'Correct' and not inBranch:
				branchcount += 1
				inBranch = True
			elif par.style == 'Incorrect':
				inBranch = False
		
		print fn.split('.')[0] + '\t' + str(branchcount)
		sys.stdout.flush()
