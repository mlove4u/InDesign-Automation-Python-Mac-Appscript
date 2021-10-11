import re
import tkinter
import json
from urllib import request
from appscript import *


def fetch_ruby(query):
    # https://developer.yahoo.co.jp/webapi/jlp/furigana/v2/furigana.html
    APPID = "***** replace your app id here *****"
    URL = "https://jlp.yahooapis.jp/FuriganaService/V2/furigana"
    headers = {
        "Content-Type": "application/json",
        "User-Agent": "Yahoo AppID: {}".format(APPID)
    }
    param_dic = {
        "id": "1234-1",
        "jsonrpc": "2.0",
        "method": "jlp.furiganaservice.furigana",
        "params": {
            "q": query,
            "grade": 1
        }
    }
    params = json.dumps(param_dic).encode()
    req = request.Request(URL, params, headers)
    with request.urlopen(req) as res:
        body = res.read()
    lists = json.loads(body.decode())["result"]["word"]
    ruby_data = []  # --> [[漢字1, ルビ1], [漢字2, ルビ2]...]
    for x in lists:
        if "furigana" in x:
            if "subword" in x:
                for sub in x["subword"]:
                    if sub["surface"] != sub["furigana"]:
                        ruby_data.append([sub["surface"], sub["furigana"]])
            else:
                ruby_data.append([x["surface"], x["furigana"]])
    return ruby_data


def set_ruby():
    sel = indd.selection()
    for obj in sel:
        _class = obj.class_()
        if _class != k.text_frame:
            print(f"Not a text frame: {_class}")
            continue
        contents = obj.contents.get()
        ruby_data = fetch_ruby(re.sub(r'\r\n|\r', '', contents))
        if not ruby_data:
            continue
        start = 0
        characters, texts = obj.characters, obj.text
        for x in ruby_data:
            word, ruby = x
            pos1 = contents.index(word, start)  # first position: zero-indexed
            pos2 = pos1 + len(word)  # second position
            text_range = texts[characters[pos1 + 1]: characters[pos2]]
            text_range.ruby_string.set(ruby)
            text_range.ruby_type.set(1249011570)  # group ruby
            text_range.ruby_flag.set(True)
            start = pos2


if __name__ == "__main__":
    #
    indd = app('Adobe InDesign 2020')
    #
    root = tkinter.Tk()
    root.resizable(False, False)
    root.title("set ruby")
    root.geometry("200x100")
    Button = tkinter.Button(text='set ruby', width=15, height=5,
                            command=set_ruby)
    Button.pack(padx=10, pady=10)
    root.mainloop()
