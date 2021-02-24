# about appscript
- HP: http://appscript.sourceforge.net/
- Installation: pip install appscript
# usage
```py
from appscript import *

indd = app("Adobe InDesign CC 2019")
# "Adobe InDesign 2020", "Adobe InDesign 2021"...
# print(indd.version())
```
# sample code
- [00_application](00_application.py)
- [01_document](01_document.py)