from appscript import *


indd = app("Adobe InDesign CC 2019")
fonts = indd.fonts  # all fonts
print("count of all fonts:", len(fonts()))

headers = ["allowEditableEmbedding",
           "allowOutlines",
           "allowPDFEmbedding",
           #
           "fontFamily",
           "fontStyleName",
           "fontStyleNameNative",
           "fontType",
           #
           "name",
           "fullName",
           "fullNameNative",
           "postscriptName"
           ]

with open("fonts.csv", "w", encoding="utf-8") as file:
    file.write("\t".join(headers) + "\n")
    #
    for font in fonts():
        try:
            a = font.allow_editable_embedding()
        except:
            a = ""
        #
        try:
            b = font.allow_outlines()
        except:
            b = ""
        #
        try:
            c = font.allow_PDF_embedding()
        except:
            c = ""
        #
        d = font.font_family()
        #
        try:
            e = font.font_style_name()
            f = font.font_style_name_native()
        except:
            e = f = ""
        #
        try:
            g = font.font_type()
        except:
            g = ""
        #
        h = font.name()
        #
        try:
            i = font.full_name()
            j = font.full_name_native()
        except:
            i = j = ""
        #
        try:
            k = font.postscript_name()
        except:
            k = ""
        #
        #
        file.write(f"{a}\t{b}\t{c}\t{d}\t{e}\t{f}\t{g}\t{h}\t{i}\t{j}\t{k}\n")
