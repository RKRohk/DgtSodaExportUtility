import json
import gzip
from os import error
import pandas as pd
import os

import simplejson

data = pd.read_excel("TasksListImport.xlsx",engine="openpyxl").to_dict(orient="records")

buf = {}


files = sorted(os.listdir(os.getcwd()), key=os.path.getmtime)
gzfiles = filter(lambda x: True if x.endswith('json.gz') else False, files)
gzfiles = list(gzfiles)
newestfile = gzfiles[-1]

try:
    with gzip.open(newestfile) as infile:
        buf = json.load(infile)
        buf["task"] = data
        buf = simplejson.dumps(buf, ignore_nan=True)

    with gzip.open(newestfile, "w") as outfile:
        outfile.write(buf.encode("utf-8"))
        outfile.close()
except error:
    print("There was an error")
    print(error)
