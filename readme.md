# About
Some example scripts for InDesign automation control on Mac.
# About appscript
- Installation: `pip install appscript`
- HP: http://appscript.sourceforge.net/

# Keyword conversion
## click [here](http://appscript.sourceforge.net/py-appscript/doc/appscript-manual/05_keywordconversion.html)
> - Characters a-z, A-Z, 0-9 and underscores (_) are preserved.
> - Spaces, hyphens (-) and forward slashes (/) are replaced with underscores.
> - Ampersands (&) are replaced by the word 'and'.
> - All other characters are converted to 0x00-style hexadecimal representations.
> - Names that match Python keywords or names reserved by appscript have an underscore appended. (※)
#### ※example: class--> class_ , properties--> properties_

# Usage
```py
from appscript import *

indd = app("Adobe InDesign CC 2019")
# print(indd.version())
doc = indd.active_document
# print(doc.name())
tfs = doc.text_frames
# print(len(tfs()))
first_tf = tfs[1] # 1-based indexing
characters = first_tf.characters
```
# Add new object
```py
# syntax: parent_object.make(new=k.type)
indd.make(new=k.document) # add a new document
doc.make(new=k.page) # add a new page to doc
doc.pages[1].make(new=k.text_frame) # add a new text frame to page 1
first_tf.make(new=k.table) # add a new talbe to first_tf
```
# Get type/class of an object
```py
doc.class_() # k.document
first_tf.class_() # k.text_frame
```

# Method with keyword
```py
# same to python keyword arguments
doc.close(saving=k.ask, saving_in="path/to/file.indd")
```

# Get value
```py
doc.name()
first_tf.contents()
```
or
```py
doc.name.get()
first_tf.contents.get()
```

# Set value
```py
first_tf.contents.set("sample contents")
first_tf.fill_color.set("Cyan")
first_tf.geometric_bounds.set([0,0,50,100])
```
or 
```py
# set properties at one time (this is faster)
first_tf.properties_.set({
        k.contents: "sample contents",
        k.fill_color: "Cyan",
        k.geometric_bounds: [0,0,50,100]
    })
```
# Get item
## by name
```py
tfs["abc"] # return the 1st text frame whose name is "abc"
# javascript: tfs.itemByName("abc")
```

## by id
```py
tfs.ID(1234) # return text frame whose ID is 1234
# javascript: tfs.itemByID(1234)
```

## by range
```py
characters[1:3] # note: include the 3rd item !!!!!
# javascript: characters.itemByRange(0, 2)
```

# Performance
Running at about 30% speed of javascript/applescript. InDesign document always redraw during script execution. It just runs faster only when we directly execute officially supported language of indesign (javascript/applescript/vbscript) from script panel.
```py
# This doesn't speed up processing
indd.script_preferences.enable_redraw.set(False)
```

# Sample code
- [00_application](00_application.py)
- [01_document](01_document.py)
- [02_text_frame](02_text_frame.py) --> [result example](data_files/02_text_frame.png)
- [03_set_ruby_to_textframe](03_set_ruby_to_textframe/readme.md)
- [04_layer](04_layer.py)
- [05_character](05_character.py) --> [result example](data_files/05_character.png)
- [06_paragraph](06_paragraph.py) --> [result example](data_files/06_paragraph.png)
- [07_font](07_font.py) --> [result example](data_files/07_font.png)
- [08_table/cell](08_table.py) --> [result example](data_files/08_table.png)
- [09_graphic](09_graphic.py) --> [result example](data_files/09_graphic.png)
- [10_pdf](10_pdf.py)
- [11_find_and_change_text](11_find_and_change_text.py) --> [example](data_files/11_find_and_change_text.png)