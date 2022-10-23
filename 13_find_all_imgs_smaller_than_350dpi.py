from appscript import *
from PIL import Image
from pathlib import Path
#
DPI = 300  # effective ppi
EXTS = (".jpg", ".jpeg", ".eps", ".png", ".tiff", ".tif", ".psd")
#


def get_eps_resolution(eps_file):
    # TODO: is Illustrator eps
    img = Image.open(eps_file)
    if "Illustrator" in img.info["Creator"]:
        print("----------Creator: ", img.info["Creator"], eps_file)
    w, h = img.size  # unit: pixel
    tmp = img.info["HiResBoundingBox"].split(" ")  # unit: point
    tmp = [float(x) for x in tmp]
    HiResWidth = tmp[2] - tmp[0]
    HiResHeight = tmp[3] - tmp[1]
    return round(w / HiResWidth * 72), round(h / HiResHeight * 72)


#
indd = app("Adobe InDesign 2022")
doc = indd.active_document
links = doc.links()
for link in links:
    hfs_path = link.file_path()  # HFS path
    # FIXME: HFS path to POSIX path. Not work in external hdd!
    file_path = hfs_path[hfs_path.index(":"):].replace(":", "/")
    ext = Path(file_path).suffix.lower()
    if ext not in EXTS:
        continue
    status = link.status()
    if status == k.link_missing:
        print("link missing: ", file_path)
        continue
    if status == k.out_of_date:
        link.update()
    # horizontal scale may not equal to vertical scale
    scale = max(link.parent.horizontal_scale(), link.parent.vertical_scale())
    # print(graphic())
    if ext == ".eps":
        resolution = get_eps_resolution(file_path)
        # use width resolution: percentage
        effective_ppi = resolution[0] / scale * 100
    else:
        graphic = link.parent
        effective_ppi = min(graphic.effective_ppi())
    #
    if effective_ppi < DPI:
        print(effective_ppi, file_path)
