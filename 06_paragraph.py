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
                   k.contents: "autodtp.app\rHello World\rcontents",
                   k.geometric_bounds: [20, 20, 100, 100]
               })
print("length of paragraphs: ", len(tf.paragraphs()))  # 3
#
p1 = tf.paragraphs[1]
p2 = tf.paragraphs[2]
p3 = tf.paragraphs[3]

# set font
fonts = indd.fonts  # all fonts
p1.applied_font.set(fonts['Arial\tRegular'])
p2.applied_font.set(fonts['Apple Chancery\tChancery'])

# set justification
p2.justification.set(k.center_align)
p3.justification.set(k.right_align)

# set underline, tracking
p3.properties_.set({
    k.underline: True,
    k.tracking: 200
})

# set color
c = add_CYMK_color(doc, [100, 0, 0, 0], "c")
m = add_CYMK_color(doc, [0, 100, 0, 0], "m")
y = add_CYMK_color(doc, [0, 0, 100, 20], "y")
tf.paragraphs[1].fill_color.set(c)
tf.paragraphs[2].stroke_color.set(m)
tf.paragraphs[2].stroke_tint.set(50)
tf.paragraphs[3].fill_color.set(y)

# add a new paragraph
tf.insertion_points[1].contents.set("the first paragraph." * 10 + "\r\r")
tf.insertion_points[-1].contents.set("\rthe last paragraph")
tf.paragraphs[-1].justification.set(k.left_align)

# remove a paragraph
# p3.contents.set("")

# set drop cap character
p1.properties_.set({
    k.drop_cap_characters: 3,
    k.drop_cap_lines: 2
})
