from appscript import *


indd = app("Adobe InDesign 2022")
doc = indd.make(new=k.document)
page1 = doc.pages[1]

# add a rectangle with properties
rectangle = page1.make(new=k.rectangle,
                       with_properties={
                           k.geometric_bounds: [0, 0, 100, 100],
                           k.stroke_weight: 0.1
                       })

# pdf file path: must be a full path
pdf = "/Users/******/Desktop/test.pdf"

# place a pdf to rectangle
indd.PDF_place_preferences.page_number.set(2)  # place the 2nd page
# rectangle.place(pdf)  # Error!
rectangle.place(mactypes.Alias(pdf))
# rectangle.place(mactypes.File(pdf))  # OK

# set fit option
rectangle.fit(given=k.frame_to_content)
rectangle.move(to=[10, 20])  # move to: x=10, y=20

# export indesign document as pdf
doc.export(format=k.PDF_type,  # k.interactive_PDF
           to="/Users/******/Desktop/mypdf.pdf",
           using="[PDF/X-1a:2001]")  # "[PDF/X-1a:2001 (日本)]"
