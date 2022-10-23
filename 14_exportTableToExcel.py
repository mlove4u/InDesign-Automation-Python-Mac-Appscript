"""
NOTE: Reproduce only cell merges and contents, others 
(such as graphic/text_frame/table in cells, colors, borders, etc) are ignored.
"""
import xlwings as xw
from appscript import app, k
indd = app("Adobe InDesign 2022")


class Table():
    def __init__(self, row_count: int, column_count: int, cell_datas: list):
        self.row_count = row_count
        self.column_count = column_count
        self.cell_datas = cell_datas


def get_info_from_indesign_table():
    if len(indd.selection()) == 0:
        print("please select a table or a text_frame containing a table")
        return
    sel = indd.selection()[0]  # text_frame or table
    class_ = sel.class_()
    if class_ == k.table:
        tables = [sel]
    elif class_ == k.text_frame:
        tables = sel.tables()
    else:
        print("Not a table nor a text_frame")
        return
    table_objs = []
    for table in tables:
        cell_datas = []
        cells = table.cells
        for cell in cells():
            c, r = [int(x) for x in cell.name().split(":")]
            cell_datas.append({
                "column": c,
                "row": r,
                "column_span": cell.column_span(),
                "row_span": cell.row_span(),
                # indesign line_break to excel line_break
                "contents": cell.contents().replace("\r", "\n")  # TODO: overflow
            })
        table_objs.append(
            Table(len(table.rows()), len(table.columns()), cell_datas)
        )
    return table_objs


def generate_excel_table(table_objs, wb=None, sh=None, start_cell_row=1, start_cell_column=1):
    # start_cell_row/column: default(1,1) --> A1 cell
    if not wb:
        wb = xw.Book()  # add a new book
        # wb = xw.books[0]  # current active book
    wb.activate(steal_focus=True)
    if not sh:  # adcive sheet
        sh = wb.sheets.active
    for table in table_objs:
        offset_row = start_cell_row - table.cell_datas[0]["row"]
        offset_column = start_cell_column - table.cell_datas[0]["column"]
        for x in table.cell_datas:
            r, c = x["row"] + offset_row, x["column"] + offset_column
            sh.cells(r, c).value = x["contents"]
            sh.cells(r, c).api.wrap_text.set(True)
            r_span, c_span = x["row_span"], x["column_span"]
            if c_span != 1 or r_span != 1:  # cells that need to be merged
                sh.range((r, c), (r + r_span - 1, c + c_span - 1)).merge()
        # set border
        end_cell_rc = (start_cell_row + table.row_count - 1,
                       start_cell_column + table.column_count - 1)
        for border in [k.border_top, k.border_bottom, k.border_left, k.border_right]:
            sh.range((start_cell_row, start_cell_column), end_cell_rc
                     ).api.get_border(which_border=border).weight.set(2)
        #
        # generate talbes from up to down
        start_cell_row = start_cell_row + table.row_count + 1  # keep a blank row
        # generate talbes from left to right
        # start_cell_column = start_cell_column + table.column_count + 1


if __name__ == "__main__":
    table_objs = get_info_from_indesign_table()
    if table_objs:
        generate_excel_table(table_objs)
