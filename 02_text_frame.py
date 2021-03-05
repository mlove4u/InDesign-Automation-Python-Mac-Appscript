from appscript import *


indd = app("Adobe InDesign CC 2019")
doc = indd.make(new=k.document)
page1 = doc.pages[1]

# add a new text frame with properties
tf1 = page1.make(new=k.text_frame,
                 with_properties={
                     k.contents: "some contents. " * 2,
                     k.geometric_bounds: [20, 20, 50, 80]
                 })
# set stroke weight, color, inner margin, opacity
green_color = doc.make(new=k.color,
                       with_properties={k.color_value: [75, 0, 100, 0]})
red_color = doc.make(new=k.color,
                     with_properties={k.color_value: [15, 100, 100, 0]})
tf1.properties_.set({k.stroke_weight: 1, k.stroke_color: green_color})
tf1.text_frame_preferences.inset_spacing.set([2, 2, 2, 2])
tf1.transparency_settings.blending_settings.opacity.set(80)

# add text frame to tf1(inline text frame)
colors = [red_color, green_color]
for i in range(10):
    obj = tf1.insertion_points[i + 1].make(new=k.text_frame,
                                           with_properties={
                                               k.contents: str(i),
                                               k.fill_color: colors[i % 2]})

tf2 = page1.make(new=k.text_frame,
                 with_properties={k.contents: "tf2 contents",
                                  k.fill_color: "Cyan",
                                  k.geometric_bounds: [0, 0, 20, 20]})
# group tf1 and tf2
page1.make(new=k.group, with_data={k.group_items: [tf1, tf2]})

# make a text frame and place a txt file
tf3 = page1.make(new=k.text_frame)
tf3.geometric_bounds.set([20, 100, 100, 200])
tf3.place(mactypes.Alias("path/to/textfile.txt"), showing_options=False)
tf3.text_frame_preferences.vertical_justification.set(k.center_align)
