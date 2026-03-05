from barcode import JAN  # pip3 install python-barcode
from appscript import app, k
id_app = app("Adobe InDesign 2024")


class InDesignJancodeMaker():
    def __init__(self):
        self.black_color = "Black"
        self.white_color = "Paper"

    def __find_true_segments(self, a: list):
        """
        Finds the true segments in a given list.
        Args:
            a (list): The input list.
        Returns:
            list: A list of tuples representing the true segments.
            Each tuple contains the start index and length of a true segment.
        """
        true_segments = []
        start = None
        length = 0
        for i, value in enumerate(a):
            if value:
                if start is None:
                    start = i
                length += 1
            else:
                if start is not None:
                    true_segments.append((start, length))
                    start = None
                    length = 0
        # handle the case where the list ends with a True segment
        if start is not None:
            true_segments.append((start, length))
        return true_segments

    def create_jancode(self, data: str, module_width=0.33,
                       module_height=22.85,
                       guard_bar_height=25.93,
                       left_quiet_zone=11, right_quiet_zone=7):
        """
        https://internationalbarcodes.com/ean-13-specifications/
        unit: mm
        """
        if len(data) != 13:  # TODO:EAN/JAN-8
            raise ValueError("JAN code must be 13 digits.", data)
        code_obj = JAN(data)  # , guardbar=True
        assert code_obj.get_fullcode() == data, "Invalid JAN code."
        bar_sequence = code_obj.build()[0]  # '10100010*******'
        # get longgest true segments (black bars)
        true_segments = self.__find_true_segments(int(x) for x in bar_sequence)
        assert len(true_segments) == 30, "Invalid JAN code length."
        # draw white background
        bg_rect_width = (left_quiet_zone + len(bar_sequence) + right_quiet_zone
                         ) * module_width
        doc = id_app.make(new=k.document)
        group_items = []
        rect = doc.make(new=k.rectangle,
                        with_properties={
                            k.name: "background",
                            k.stroke_weight: 0,
                            k.fill_color: self.white_color,  # white background
                            k.geometric_bounds: [0, 0, guard_bar_height, bg_rect_width],
                        })
        group_items.append(rect)
        # draw module rectangles
        guard_bar_pos = (0, 1, 14, 15, 28, 29)
        for i, (start_pos, length) in enumerate(true_segments):
            height = guard_bar_height if i in guard_bar_pos else module_height
            rect = doc.make(new=k.rectangle,
                            with_properties={
                                k.stroke_weight: 0,
                                k.fill_color: self.black_color,  # black bars
                                k.geometric_bounds: [0, (start_pos + left_quiet_zone) * module_width,
                                                     height, (start_pos + left_quiet_zone + length) * module_width],
                            })
            group_items.append(rect)
        # group background and bars together
        group_obj = doc.make(new=k.group,
                             with_properties={k.group_items: group_items})
        group_obj.name.set("JANCode")
        return doc


my_code = InDesignJancodeMaker()
my_code.create_jancode("4901550151296")
