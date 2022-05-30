import glob
import re
import os
import win32com.client as win32
from win32com.client import constants
from pathlib2 import Path
from datetime import datetime
import time

# Create list of paths to .doc files
inital_paths = glob.glob('\\Users\\William\\Desktop\\CodeProjects\\MPR\\*[0-9]', recursive=True)

# Standardizes the paths so they can be sorted by version number, only the most recent MPR is needed
paths = []
for initial_path in inital_paths:
    presorted_paths = glob.glob(initial_path + '\\**\\*[0-9]*.doc', recursive=True)
    #print(presorted_paths)
    def pathevener(xpath):
        xpath = xpath.replace(" ","")
        xpath = xpath.replace("\Obsolete","")
        xseppath = xpath.split('\\')
        return(xseppath[7])
    sorted_paths = sorted(presorted_paths, key=pathevener, reverse=True)
    try:
        paths.append(sorted_paths[0])
    except:
        continue
    
#print(paths)

def save_as_docx(path, word):
    # Opening MS Word
    doc = word.Documents.Open(path)
    doc.Activate()

    # Rename path with .docx
    new_file_abs = os.path.abspath(path)
    new_file_abs = new_file_abs.replace('.doc', '.docx')

    # Save and Close
    word.ActiveDocument.SaveAs(
        new_file_abs, FileFormat=constants.wdFormatXMLDocument
    )
    doc.Close()


start = datetime.now()
current_time = start.strftime("%H:%M:%S")
print('start time =', current_time)

word = win32.Dispatch('Word.Application')

for path in paths:
    catnum = Path(path).parts[6]
    now = datetime.now()
    current_time = now.strftime("%H:%M:%S")
    print(catnum, 'Current time =', current_time)
    try:
        save_as_docx(path, word)
        time.sleep(2)
    except Exception as e:
        print('Error in', catnum, e)
        continue
    #print(path)

word.Quit()

end = datetime.now()
current_time = end.strftime("%H:%M:%S")
print('end time =', current_time)
