from appscript import *


indd = app("Adobe InDesign 2022")
doc = indd.make(new=k.document)
page1 = doc.pages[1]

# add a rectangle with properties
rectangle = page1.make(new=k.rectangle,
                       with_properties={
                           k.geometric_bounds: [20, 20, 100, 100],
                           k.stroke_weight: 0.5
                       })
print(rectangle.class_())  # k.rectangle

# image path: must be a full path
img = "/Users/***/Desktop/test.png"
# img = "/Users/***/Desktop/test.pdf" # place a pdf file(1st page)

# place a image to rectangle
# rectangle.place(img)  # Error!
rectangle.place(mactypes.Alias(img))
# rectangle.place(mactypes.File(img))  # OK

# set fit option
rectangle.fit(given=k.proportionally)

# set opacity to image
rectangle.transparency_settings.blending_settings.opacity.set(80)

# rotate rectangle/image
# rectangle.rotation_angle.set(-45)  # rotate rectangle
# rectangle.graphics[1].rotation_angle.set(-45)  # rotate 1st graphic

# get file path
file_path = rectangle.graphics[1].item_link.file_path.get()
print("file_path: ", file_path)  # Macintosh HD:Users:***:Desktop:test.png

# get parent layer name
layer_name = rectangle.graphics[1].item_layer.name.get()
print("layer_name: ", layer_name)  # レイヤー 1/Layer 1


# add caption(file path) below image
b = rectangle.visible_bounds.get()
bounds = [b[2], b[1], b[2] + 10, b[3]]
page1.make(new=k.text_frame,
           with_properties={
               k.visible_bounds: bounds,
               k.contents: file_path
           })

# add a QR code
rectangle2 = page1.make(new=k.rectangle,
                        with_properties={
                            k.geometric_bounds: [20, 120, 50, 150],  # w=h=30
                            k.fill_color: "Paper",
                            k.stroke_color: "None",
                            k.stroke_weight: 0,
                        })
rectangle2.Create_Plain_Text_QR_Code(plain_text="https://www.autodtp.app")
