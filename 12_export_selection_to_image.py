"""
Usage:
1. Select object(s) in InDesign document
2. Click "Export"
"""
import os
import wx  # pip3 install wxPython
from appscript import *

indd = app("Adobe InDesign 2022")
desktop = os.path.expanduser("~/Desktop/")  # default saveto = desktop


def selection_to_jpg(selection, saveto):
    indd.JPEG_export_preference.properties_.set(
        {
            # image
            k.JPEG_Quality: k.maximum,  # k.low / k.medium / k.high / k.maximum
            k.JPEG_Rendering_style: k.baseline_encoding,  # k.progressive_encoding
            k.export_resolution: 72,
            k.jpeg_color_space: k.RGB,  # k.RGB / k.CMYK / k.gray
            # option
            k.embed_color_profile: False,
            k.anti_alias: True,
            k.use_document_bleeds: False,
            k.simulate_overprint: False,
        }
    )
    indd.export(selection, format=k.JPG, to=saveto)


def selection_to_png(selection, saveto):
    indd.PNG_export_preference.properties_.set(
        {
            # image
            k.PNG_Quality: k.maximum,  # k.low / k.medium / k.high / k.maximum
            k.export_resolution: 72,
            k.PNG_color_space: k.RGB,  # k.RGB / k.gray
            # option
            k.transparent_background: True,
            k.anti_alias: True,
            k.use_document_bleeds: False,
            k.simulate_overprint: False,
        }
    )
    indd.export(selection, format=k.PNG_format, to=saveto)


class MyFrame (wx.Frame):

    def __init__(self):
        wx.Frame.__init__(self, None, id=-1, title="Export Image",
                          size=wx.Size(200, 150))

        bSizer = wx.BoxSizer(wx.VERTICAL)
        self.fmt = wx.RadioBox(self, -1, "Format", wx.DefaultPosition,
                               wx.DefaultSize, ["JPG", "PNG"], 1,
                               wx.RA_SPECIFY_ROWS)
        self.fmt.SetSelection(0)
        self.btn = wx.Button(self, -1, "Export")
        bSizer.Add(self.fmt, 0, wx.ALL, 5)
        bSizer.Add(self.btn, 0, wx.ALL, 5)

        self.SetSizer(bSizer)
        self.Layout()
        self.Centre(wx.BOTH)

        # Connect Events
        self.btn.Bind(wx.EVT_BUTTON, self.export)

    def export(self, event):
        fmt = self.fmt.GetStringSelection()
        if fmt == "JPG":
            fn = selection_to_jpg
        elif fmt == "PNG":
            fn = selection_to_png
        else:
            raise Exception("Unknown format")
        #
        for i, x in enumerate(indd.selection()):
            fn(x, f"{desktop}/_test_{i}.{fmt}")
        print("Exported")


if __name__ == '__main__':
    app = wx.PySimpleApp()
    MyFrame().Show()
    app.MainLoop()
