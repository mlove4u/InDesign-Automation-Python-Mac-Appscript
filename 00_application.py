from appscript import *

indd = app("Adobe InDesign CC 2019")

# current selected objects
sel = indd.selection()  # list
for obj in sel:
    print(obj.class_())  # type of object--> applescript: class of obj

# active document
doc = indd.active_document
doc.text_frames[1].select()  # select the 1st text frame: 1-based indexing

# fonts
fonts = indd.fonts()
for font in fonts:
    print(font.name())

# user_interaction_level
print(indd.script_preferences.user_interaction_level())

# do javascript
js = "alert(arguments[0] + arguments[1]);"
# js = "path/to/javascript/file"
indd.do_script(js, language=1246973031, with_arguments=[1, 2])

# preflight profiles
for pp in indd.preflight_profiles():
    print(pp.name())
