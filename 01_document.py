from appscript import *

indd = app("Adobe InDesign CC 2019")
# print(indd.version())  # 14.*.*.***

# open an existing document
# doc1 = indd.open("path/to/indesign/file.indd")

# add a new document
doc1 = indd.make(new=k.document)
# set document size
doc1.document_preference.page_height.set(100)
doc1.document_preference.page_width.set(100)
doc1.name.set("doc1")

# add another new document
doc2 = indd.make(new=k.document)
# set properties at one time
doc2.document_preference.properties_.set({k.page_height: 200,
                                          k.page_width: 200})
doc2.name.set("doc2")

documents = indd.documents  # collection of documents
print(f"count of documents: {len(documents())}")  # 2
print(indd.active_document.name())  # doc2
indd.active_document.set(doc1)  # set active document to doc1
print(indd.active_document.name())  # doc1

# add page to doc1
doc1.make(new=k.page)  # add to end
doc1.make(new=k.page, at=doc1.beginning)  # add to beginning
pages = doc1.pages  # collection of pages
print(f"count of pages: {len(pages())}")  # 3

# save and close document
doc1.save(to="path/to/indesign/file.indd")
doc1.close()

doc2.close(saving=1634954016)  # doc.close(saving=k.ask)
