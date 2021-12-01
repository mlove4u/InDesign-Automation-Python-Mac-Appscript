from appscript import *


def add_CYMK_color(doc, values: list, name: str):
    # add a CMYK color to document
    return doc.make(new=k.color,
                    with_properties={
                        k.model: k.process,
                        k.space: k.CMYK,
                        k.color_value: values,
                        k.name: name
                    })


indd = app("Adobe InDesign CC 2019")
doc = indd.make(new=k.document)
page = doc.pages[1]
tf = page.make(new=k.text_frame,
               with_properties={
                   k.contents: "DEFGHIJKLMNOPQRSTUVWXYZ",
                   k.geometric_bounds: [20, 20, 40, 85]
               })

# add characters
tf.insertion_points[1].contents.set("ABC\r")

characters = tf.characters

# set size
characters[1].point_size.set("20 pt")  # point
characters[2].point_size.set("40 pt")
characters[3].point_size.set("40 Q")  # quarter

# add color to document
red_color = add_CYMK_color(doc, [0, 100, 100, 0], "red")
orange_color = add_CYMK_color(doc, [0, 80, 100, 0], "orange")
# get swatches from document
cyan_color = doc.swatches["C=100 M=0 Y=0 K=0"]

# set fill color to characters
characters[2].fill_color.set(red_color)
characters[3].fill_color.set(orange_color)

# font
fonts = indd.fonts  # list all fonts
# for font in fonts():
#     print(font.name.get())
# apply font
Arial_font = fonts['Arial\tRegular']
characters[2].applied_font.set(Arial_font)

# set horizontal and vertical scale
characters[2].properties_.set({
    k.horizontal_scale: 80,
    k.vertical_scale: 150
})

# set stroke color, weight
characters[2].stroke_color.set(cyan_color)
characters[2].stroke_weight.set("0.5mm")

# set underline
characters[3].properties_.set({
    k.underline: True,
    k.underline_color: "Black"})

# create outline
tf2 = page.make(new=k.text_frame,
                with_properties={
                    k.contents: "test " * 5,
                    k.geometric_bounds: [50, 20, 60, 85]
                })
# all characters
# tf2.create_outlines(tf2, delete_original=True)
# only sixth character
tf2.create_outlines(tf2.characters[6], delete_original=True)
# 8,9th character. Error when special characters exists --> Space,Tab,CR,etc.
tf2.create_outlines(tf2.characters[8:9], delete_original=True)
