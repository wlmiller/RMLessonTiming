from docx import *
import re
import sys,os,time
from comtypes.client import CreateObject
import comtypes.gen
import wave, contextlib
from Tkinter import Tk
from tkFileDialog import askopenfilename

Tk().withdraw()
filename = askopenfilename()
print filename
