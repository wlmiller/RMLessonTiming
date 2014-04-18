from lessonitemstats import getlessonitemstats
import os,sys

print "item\twc\tsubmittime\twtd count\tnext count\ttotal time\tmain time\tNR time\tavg branch time\tonscreen wc\tshort count\tmedium count\tlong count\tnonstandard submit time"
for l in os.listdir("../scripts"):
	for fn in [f for f in os.listdir("../scripts/" + l) if '.docx' in f and '-' in f]:
		stats = getlessonitemstats("../scripts/" + l + "/" + fn)
		print fn.replace(".docx","") + "\t" + str(stats["word count"]) + "\t" + str(stats["submit time"]) + "\t" + str(stats["WTD count"]) + "\t" + str(stats["next count"]) + "\t" + str(stats["dialogue time (total)"]) + "\t" + str(stats["dialogue time (main branch)"]) + "\t" + str(stats["dialogue time (NR branch)"]) + "\t" + str(stats["average branch time"]) + "\t" + str(stats["onscreen text word count"]) + "\t" + str(stats['short count']) + "\t" + str(stats['medium count']) + "\t" + str(stats['long count']) + "\t" + str(stats['nonstandard submit time'])
		sys.stdout.flush()
