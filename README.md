RM Lesson Timing
======

The main code for calculating lesson time is [lessontiming.py](lessontiming.py).

This code estimates lesson lengths for RM lessons.  The code asks the user to select the OS file for the lesson; it requires the OS file name to include the lesson number and the lesson item scripts to be located in a directory called `Scripts/` in the same directory as the OS file.  The OS file and all lesson item scripts must be in the `docx` format.

__Python 2.7 is required.__  There is one optional command line argument; the path to the OS file can be supplied directly, skipping the file selection dialog.

The included Python scripts are:  

1. [lessontiming.py](lessontiming.py): Given an OS file, extracts the "weak + behind" and "weak + ontime" paths and estimates the length of each lesson item as well as the total length of these two paths.  The resulting data are placed in file called `[lesson]_timing.csv` in the same directory as the OS file, and this file is automatically opened for the user.
2. [parseOSfile.py](parseOSfile.py): Called by `lessontiming.py` and `getallpaths.py`, extracts the lesson paths from a given OS file.
3. [lessonitemstats.py](lessonitemstats.py): Called by `lessontiming.py` and `parseallitems.py`, extracts the relevant features in the given lesson item script.
4. [getallpaths.py](getallpaths.py): Loops through all lessons in the curriculum and extracts the "weak + behind" path using `parseOSfile.py`.
5. [parseallitems.py](parseallitems.py): Loops through all lesson _items_ in the curriculum and extracts the relevant statistics using `lessonitemstats.py`.

__Note__: _Some of the feature extraction methods may seem somewhat tortured.  This is because the format of OS files and scripts is almost, but not quite, standard._
