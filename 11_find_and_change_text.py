from appscript import *


indd = app("Adobe InDesign 2022")
doc = indd.active_document

# important: reset all preferences to nothing
indd.find_text_preferences.set(k.nothing)
indd.change_text_preferences.set(k.nothing)
#
indd.find_change_text_options.include_locked_layers_for_find.set(True)
indd.find_change_text_options.include_locked_stories_for_find.set(True)
indd.find_change_text_options.include_hidden_layers.set(False)
indd.find_change_text_options.include_master_pages.set(False)
indd.find_change_text_options.include_footnotes.set(True)
indd.find_change_text_options.case_sensitive.set(False)
indd.find_change_text_options.whole_word.set(False)
indd.find_change_text_options.kana_sensitive.set(True)
indd.find_change_text_options.width_sensitive.set(False)
#
indd.find_text_preference.find_what.set("begining")
indd.change_text_preference.change_to.set("beginning")
result = doc.find_text()  # 275, 298, 324
for x in result:
    print(x.contents())  # begining
    print(x.parent_text_frames()[0].id())
doc.change_text()

# ---------- find only special character style
char_style = doc.character_styles["greenChar"]
indd.find_text_preference.applied_character_style.set(char_style)
result = doc.find_text()  # only 275
for x in result:
    print(x.parent_text_frames()[0].id())
doc.change_text()

# ---------- find and change Grep
indd.find_grep_preference.find_what.set("^begining")  # RegEx
result = indd.find_grep()  # only 324
for x in result:
    print(x.parent_text_frames()[0].id())
