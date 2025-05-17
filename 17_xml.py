from appscript import app


indd = app("Adobe InDesign 2024")
doc = indd.active_document
tf = doc.text_frames[1]
xml_elements = tf.associated_XML_element.XML_elements()
#
#
# associate data with xml elements
origin_data = {
    # key:tag name
    "name": "John",
    "age": "30",
    "country": "USA"
}
for x in xml_elements:
    tag_name = x.markup_tag.name()
    x.contents.set(origin_data[tag_name])
#
#
# extract data from xml elements
extracted_data = {}
for x in xml_elements:
    tag_name = x.markup_tag.name()
    extracted_data[tag_name] = x.contents()
print("extracted_data: ", extracted_data)
