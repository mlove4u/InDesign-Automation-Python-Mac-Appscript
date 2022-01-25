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
page1 = doc.pages[1]

# add a new text frame with properties
tf = page1.make(new=k.text_frame,
                with_properties={
                    k.contents: "",
                    k.geometric_bounds: [20, 20, 200, 150]
                })
tf.insertion_points[1].contents.set("This is a table:")

# add table with 10*3 cells
tbl = tf.make(new=k.table, with_properties={
    k.body_row_count: 10,
    k.column_count: 3
})
rows = tbl.body_row_count()
columns = tbl.column_count()
print("rows: ", rows)  # 10
print("columns: ", columns)  # 3
#
red_color = add_CYMK_color(doc, [0, 100, 100, 0], "red")
cyan_color = add_CYMK_color(doc, [100, 0, 0, 0], "cyan")

# add index to all cells as contents
tbl.contents.set([str(i) for i in range(rows * columns)])  # 0~29
# tbl.rows[1].contents.set(["a", "b", "c"])

# set width, height, auto grow
tbl.columns[1].width.set(20)
tbl.columns[-1].width.set(30)
tbl.rows[1].height.set(10)
tbl.rows[2].auto_grow.set(False)

# set storke weight and color
tbl.top_border_stroke_weight.set(1)
tbl.top_border_stroke_color.set(red_color)
tbl.properties_.set({
    k.bottom_border_stroke_weight: 1,
    k.bottom_border_stroke_color: cyan_color
})

#
# delete row, column
# tbl.rows[1].delete()
# tbl.columns[-1].delete()

# cells
cells = tbl.cells()  # all cells in table
# set cell edge stroke
for cell in cells:
    cell.properties_.set({
        k.left_edge_stroke_weight: 0.1,
        k.right_edge_stroke_weight: 0.1
    })
# add fill color to row
for r in range(0, rows, 2):
    for c in range(columns):
        cells[r * columns + c].properties_.set({
            k.fill_color: cyan_color,
            k.fill_tint: 20  # 0-100
        })
# set alignment
cells[0].vertical_justification.set(k.center_align)
cells[0].paragraphs[1].justification.set(k.center_align)
# merge cells
cells[1].merge(with_=cells[4])  # merge two cells
cells[-1].merge(with_=cells[-5])  # merge four cells
# create outline
cells[10].create_outlines(cells[10], delete_original=True)
cells[11].create_outlines(cells[11].characters[1], delete_original=True)
