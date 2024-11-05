#!/usr/bin/env python
# -*- coding: utf-8 -*-

######################################
# region    <IMPORT>

from pathlib import Path

from logging import getLogger
import tkinter as tk
from tkinterdnd2 import DND_FILES, TkinterDnD
import pyautogui
from PIL import Image, ImageTk
from rich.console import Console
console = Console(stderr=True)

import msg2eml

log = getLogger(__package__)
log_trace = getLogger("trace." + __package__)


class Core:
    WINDOW_SIZE_X = 300
    WINDOW_SIZE_Y = 300
    ANIME_RES = 30  # fps
    ANIME_SPAN = 1  # sec
    ANIME_TICK = 500 // ANIME_RES  # ms
    ANIME_TOTAL_TICKS = ANIME_RES * ANIME_SPAN

    DEFAULT_MSG = "D&D: Convert *.msg to *.eml. \nDouble click: Close window."

    def __init__(self, img_bg, img_ok, img_ng):
        # self.msg2eml = msg2eml.Converter()
        pass

        # ウィンドウの作成
        self.root = root = TkinterDnD.Tk()
        root.title("D&D *.eml")

        # サイズ設定・変更禁止
        x, y = pyautogui.position()
        w = self.WINDOW_SIZE_X
        h = self.WINDOW_SIZE_Y
        root.geometry(f"{w}x{h}")
        # root.geometry(f"+{x-w//2}+{y-h//2}")
        root.geometry(f"+{x}+{y}")
        root.resizable(False, False)

        # 位置変数の初期化
        self.x = 0
        self.y = 0
        self.root_x = 0
        self.root_y = 0

        # 透過表示・タイトルバー非表示
        # root.configure(bg="SystemButtonFace")
        root.attributes("-alpha", 0.8)
        root.overrideredirect(True)

        tpcolor = "#F0FEFF"
        root.configure(bg=tpcolor)
        root.wm_attributes("-transparentcolor", tpcolor)

        # 最前面表示
        root.lift()
        root.attributes("-topmost", True)

        # 背景画像の取得
        img_bg = Image.open(img_bg)
        img_ok = Image.open(img_ok)
        img_ng = Image.open(img_ng)

        self.tk_img_bg = ImageTk.PhotoImage(img_bg)
        self.tk_img_ok = ImageTk.PhotoImage(img_ok)
        self.tk_img_ng = ImageTk.PhotoImage(img_ng)

        # アニメーション用グラデーション画像の設定
        step = 1.0 / self.ANIME_TOTAL_TICKS
        blendimg = lambda i, ini: ImageTk.PhotoImage(
            Image.blend(ini, img_bg, (i + 1) * step)
        )

        r = list(range(self.ANIME_TOTAL_TICKS))
        self.anime_ok = [blendimg(i, img_ok) for i in r]
        self.anime_ng = [blendimg(i, img_ng) for i in r]

        # Labelを作成・配置
        self.label = label = tk.Label(
            root,
            bg=tpcolor,
            compound=tk.CENTER,
        )
        self.set_window_default()
        label.pack(fill=tk.BOTH, expand=True)

        # D&D イベントの登録
        root.drop_target_register(DND_FILES)
        root.dnd_bind("<<Drop>>", self.drop)

        # ドラッグによるウィンドウ移動
        root.bind("<Button-1>", self.on_click)
        root.bind("<B1-Motion>", self.on_motion)

        # ダブルクリックで終了
        root.bind("<Double-1>", lambda e: root.destroy())

    def on_click(self, event):
        # ドラッグによるウィンドウ移動: クリック時に初期位置を記憶
        self.x = event.x
        self.y = event.y
        self.root_x = self.root.winfo_x()
        self.root_y = self.root.winfo_y()

    def on_motion(self, event):
        # ドラッグによるウィンドウ移動: 初期位置からの移動量に応じてウィンドウを移動
        delta_x = event.x - self.x
        delta_y = event.y - self.y
        next_x = self.root.winfo_x() + delta_x
        next_y = self.root.winfo_y() + delta_y
        self.root.geometry(f"+{next_x}+{next_y}")

    def parse_dad_input(self, input):
        # D&D で入力されたファイルパスを解析
        log.debug(f"D&D input: {type(input)}")
        log.debug(f"D&D input: {input}")
        escaped = input
        escaped = escaped.replace(r"\{", "<L>")
        escaped = escaped.replace(r"\}", "<R>")
        escaped = escaped.replace(r"\ ", "<S>")
        splitted = [s.strip() for s in escaped.split(" ")]
        log.debug(f"Escaped: {splitted}")
        n = len(splitted)
        path_list = []
        index = 0
        while index < n:
            s = splitted[index]
            if s[0] == "{":
                s = s[1:]
                while index < n - 1:
                    index += 1
                    t = splitted[index]
                    s += " " + t
                    if len(t) > 0 and t[-1] == "}":
                        break
            s = s.rstrip("} ")
            path_list.append(s)
            index += 1
        decoded_list = []
        for p in path_list:
            dec = p
            dec = dec.replace("<L>", "{")
            dec = dec.replace("<R>", "}")
            dec = dec.replace("<S>", " ")
            decoded_list.append(dec)
        log.debug(f"Decoded: {decoded_list}")
        return decoded_list

    def drop(self, event):
        try:
            path_list = self.parse_dad_input(event.data)
            if type(path_list) == str:
                path_list = [path_list]
            self._run_msg2eml(path_list)
        except Exception as e:
            self.set_window_ng(f"Unknown error:\n{str(e)[15:]}...")
            console.print_exception()

    def _run_msg2eml(self, path_list):
        # D&D の処理
        for path in path_list:
            path = Path(path)
            log.info(f"msg2eml: {path}")
            if not path.is_file:
                self.set_window_ng(f"File not found:\n{path}")
                return
            if path.suffix != ".msg":
                self.set_window_ng(f"Invalid extention:\n{path.suffix}")
                return
            conv = msg2eml.Converter()
            conv.read_msg(path)
            log.info(conv.get_description(show_body=False))
            conv.save_as_eml()
        self.set_window_ok("Success.")

    def set_window_ok(self, msg=None):
        self.label.configure(image=self.tk_img_ok)
        self.label.configure(text=msg)
        self.label.configure(
            font=(
                "Arial",
                16,
                "bold",
                "roman",
                "underline",
                "normal",
            )
        )
        self.tick = 0
        self.root.after(500, self.anime_from_ok)

    def set_window_ng(self, msg=None):
        self.label.configure(image=self.tk_img_ng)
        self.label.configure(text=msg)
        self.label.configure(
            font=(
                "Arial",
                16,
                "bold",
                "roman",
                "underline",
                "normal",
            )
        )
        self.tick = 0
        self.root.after(2000, self.anime_from_ng)

    def set_window_default(self):
        self.label.configure(text=self.DEFAULT_MSG)
        self.label.configure(image=self.tk_img_bg)
        self.label.configure(
            font=(
                "Arial",
                11,
                "bold",
                "roman",
                "normal",
                "normal",
            )
        )

    def anime_from_ok(self):
        if self.tick >= self.ANIME_TOTAL_TICKS:
            self.set_window_default()
            return
        self.label.configure(image=self.anime_ok[self.tick])
        self.tick += 1
        self.root.after(self.ANIME_TICK, self.anime_from_ok)

    def anime_from_ng(self):
        if self.tick >= self.ANIME_TOTAL_TICKS:
            self.set_window_default()
            return
        self.label.configure(image=self.anime_ng[self.tick])
        self.tick += 1
        self.root.after(self.ANIME_TICK, self.anime_from_ng)

    def run(self):
        self.root.mainloop()
