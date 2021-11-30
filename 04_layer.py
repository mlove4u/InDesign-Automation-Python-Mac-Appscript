from appscript import *


indd = app("Adobe InDesign CC 2019")
doc = indd.make(new=k.document)

# set new name to 1st layer
doc.layers[1].name.set("Layer 1")

# add another 4 new layers
for i in range(2, 6):
    doc.make(new=k.layer, with_properties={k.name: f"Layer {i}"})

# list of layer objects
layers = doc.layers()

# count of layers
print(f"count of layers: {len(layers)}")  # 5

# lock the 1st, 3nd, 5th layers; make 2nd, 4th layers invisible
for i, layer in enumerate(layers):
    if (i + 1) in (1, 3, 5):
        layer.locked.set(True)
    elif (i + 1) in (2, 4):
        layer.visible.set(False)

# duplicate the 1st layer
new_layer = layers[0].duplicate()
new_layer.name.set("new_layer")

# unlock layer
new_layer.locked.set(False)

# activate layer
doc.active_layer.set(new_layer)

# move layer
target_layer = doc.layers["Layer 3"]
new_layer.move(to=target_layer.after)  # below target_layer
new_layer.move(to=target_layer.before)  # above target_layer

# most bottom layer of document: below 4 lines are the same
# synax: a_layer.move(to=any_layer.end)
new_layer.move(to=layers[0].end)  # index: (-)0~len(layers)
# new_layer.move(to=layers[1].end)
# new_layer.move(to=layers[-1].end)
# new_layer.move(to=target_layer.end)

# most top layer of document
# synax: a_layer.move(to=any_layer.beginning)
new_layer.move(to=layers[0].beginning)
