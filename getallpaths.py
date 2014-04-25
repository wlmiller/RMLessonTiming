import os
from parseOSfile import parseOSfile
import re

contentdir = "C://Users/nmiller.RMCITY/Desktop/svn/"

for s in [f for f in os.listdir(contentdir) if re.match('0[0-5]-',f)]:
    for l in os.listdir(contentdir + '/' + s):
        lesson = l.split('-')[0]
        osfn = contentdir + '/' + s + '/' + l + '/3-OS/' +  lesson + '.docx'


        if os.path.exists(osfn):
            path = parseOSfile(osfn)["weak + ontime"]
            for item in sorted(set(path)):
                print "6-" + lesson + "\t" + lesson + "-" + item.encode('ascii')
