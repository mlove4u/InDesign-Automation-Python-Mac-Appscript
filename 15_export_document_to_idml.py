import os
from glob import glob
from appscript import *

indd = app("Adobe InDesign 2022")
indd.script_preferences.user_interaction_level.set(1699640946)  # NEVER

target_fld = "/***path/***to/***fld/"
files = glob(target_fld + "*.indd")
for file in files:
    if file.startswith("."):
        continue
    saveto = os.path.splitext(file)[0] + ".idml"  # save to the same folder
    doc = indd.open(file)
    doc.export(format=k.InDesign_markup, to=saveto)
    doc.close(saving=k.no)
