import sys
import glob
import os

directory_name = ""
files = []
try:
    directory_name=sys.argv[1]
    print(directory_name)
    files = glob.glob(directory_name + "/*.geojson")
    print(files)
    for file in files:
        os.system("python3 geojsonToExcel.py " + file)
except Exception as e: print(e)