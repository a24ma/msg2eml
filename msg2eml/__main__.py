#!/usr/bin/env python
# -*- coding: utf-8 -*-

######################################
# region    <IMPORT>

import pathlib
from pathlib import Path
from pprint import pformat

import tkinter as tk
from tkinterdnd2 import DND_FILES, TkinterDnD
import pyautogui
from PIL import Image, ImageTk

import msg2eml

# pip install tkinterdnd2 pyautogui

# endregion </IMPORT>
######################################
# region    <USER_SETTINGS>

log_package_name = __package__  # Avoid __package__==None on main
sub_lib = [tk, pyautogui, Image, ImageTk]

# endregion </USER_SETTINGS>
######################################
# region    <UTILITY>

import click
import logging
import subprocess
from rich import print
from rich.console import Console
from rich.logging import RichHandler

log_trace_name = log_package_name + ".trace"  # logger with different logLevel
suppress_lib_on_err = [click, pathlib] + sub_lib  # Suppressed lib.'s stacktrace

log = logging.getLogger(log_package_name)
log_trace = logging.getLogger(log_trace_name)
console = Console(stderr=True)


def set_logger(debug_mode, quiet_mode, trace_mode=False):
    if debug_mode or trace_mode:
        log_level = logging.DEBUG
    elif quiet_mode:
        log_level = logging.WARNING
    else:
        log_level = logging.INFO

    log_fmt = "[%(name)10s] %(message)s"
    log_datefmt = "[%X]"  # w/o date
    # log_datefmt = "[%Y-%m-%d %X]" # w/ date
    log_fmt = logging.Formatter(fmt=log_fmt, datefmt=log_datefmt)
    log_hdl = RichHandler(rich_tracebacks=True, console=console)
    log_hdl.setFormatter(log_fmt)

    log.setLevel(log_level)
    log.addHandler(log_hdl)
    # log.addHandler(logging.FileHandler("log.txt"))

    if trace_mode:
        log_trace.setLevel(logging.DEBUG)
    else:
        log_trace.setLevel(logging.ERROR)


def waitkey(wait_sec: int, output=print, **output_kwargs):
    output("Press any key:", **output_kwargs)
    cmd = f"read -t {wait_sec} -n 1"
    subprocess.run(f"bash -c '{cmd}'", shell=True, cwd="/")


# endregion </UTILITY>
######################################
# region    <MAIN>


def debug(obj):
    log.debug(pformat(obj))


class Core:
    WINDOW_SIZE_X = 300
    WINDOW_SIZE_Y = 260
    ANIME_RES = 30  # fps
    ANIME_SPAN = 2  # sec
    ANIME_TICK = 500 // ANIME_RES  # ms
    ANIME_TOTAL_TICKS = ANIME_RES * ANIME_SPAN

    DEFAULT_MSG = "D&D: Convert *.msg to *.eml. \nDouble click: Close window."

    def __init__(self, img_bg, img_ok, img_ng):
        self.msg2eml = msg2eml.Converter()

        # ウィンドウの作成
        self.root = root = TkinterDnD.Tk()
        root.title("D&D *.eml")

        # サイズ設定・変更禁止
        x, y = pyautogui.position()
        w = self.WINDOW_SIZE_X
        h = self.WINDOW_SIZE_Y
        root.geometry(f"{w}x{h}")
        root.geometry(f"+{x-w//2}+{y-h//2}")
        root.resizable(False, False)

        # 位置変数の初期化
        self.x = 0
        self.y = 0
        self.root_x = 0
        self.root_y = 0

        # 透過表示・タイトルバー非表示
        root.configure(bg="SystemButtonFace")
        root.attributes("-alpha", 0.7)
        root.overrideredirect(True)

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
            text=self.DEFAULT_MSG,
            image=self.tk_img_bg,
            compound=tk.CENTER,
        )
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
        escaped = input
        escaped = escaped.replace(r"\{", "<L>")
        escaped = escaped.replace(r"\}", "<R>")
        escaped = escaped.replace(r"\ ", "<S>")
        splitted = [s.strip() for s in escaped.split(" ")]
        log.debug(splitted)
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
                    if t[-1] == "}":
                        s += " " + t[:-1]
                        break
                    else:
                        s += " " + t
            path_list.append(s)
            index += 1
        decoded_list = []
        for p in path_list:
            dec = p
            dec = dec.replace("<L>", "{")
            dec = dec.replace("<R>", "}")
            dec = dec.replace("<S>", " ")
            decoded_list.append(dec)
        return decoded_list

    def drop(self, event):
        # D&D の処理
        debug(event.data)
        path_list = self. parse_dad_input(event.data)
        debug(path_list)
        if type(path_list) == str:
            path_list = [path_list]
        for path in path_list:
            path = Path(path)
            log.info(f"msg2eml: {path}")
            if not path.is_file:
                self.set_window_ng(f"file not found:\n{path}")
                return
            if path.suffix != ".msg":
                self.set_window_ng(f"invalid extention:\n{path.stem}")
                return
            self.msg2eml.read_msg(path)
            log.info(self.msg2eml.get_description(show_body=False))
            self.msg2eml.save_as_eml()
        self.set_window_ok()

    def set_window_ok(self, msg=None):
        self.label.configure(image=self.tk_img_ok)
        self.label.configure(text=msg)
        self.tick = 0
        self.anime_from_ok()

    def set_window_ng(self, msg=None):
        self.label.configure(image=self.tk_img_ng)
        self.label.configure(text=msg)
        self.tick = 0
        self.anime_from_ng()

    def set_window_default(self):
        self.label.configure(text=self.DEFAULT_MSG)
        self.label.configure(image=self.tk_img_bg)

    def anime_from_ok(self):
        if self.tick >= self.ANIME_TOTAL_TICKS:
            self.label.configure(image=self.tk_img_bg)
            return
        self.label.configure(image=self.anime_ok[self.tick])
        self.tick += 1
        self.root.after(self.ANIME_TICK, self.anime_from_ok)

    def anime_from_ng(self):
        if self.tick >= self.ANIME_TOTAL_TICKS:
            self.label.configure(image=self.tk_img_bg)
            return
        self.label.configure(image=self.anime_ng[self.tick])
        self.tick += 1
        self.root.after(self.ANIME_TICK, self.anime_from_ng)

    def run(self):
        self.root.mainloop()


def main(**kwargs):
    core = Core("./bg.png", "./ok.png", "./ng.png")
    core.run()
    return True


# endregion </MAIN>
######################################
# region    <CUI>


@click.command()
@click.option("--quiet", "-q", "quiet", is_flag=True)
@click.option("--debug", "-d", "debug", is_flag=True)
@click.option("--trace", "-dd", "trace", is_flag=True)
@click.option("--wait_ok", "-wok", "wait_ok", type=int, default=0)
@click.option("--wait_ng", "-wng", "wait_ng", type=int, default=180)
def cmd(**kwargs):
    set_logger(kwargs["debug"], kwargs["quiet"], kwargs["trace"])
    wait_time = kwargs["wait_ok"]
    wait_time_ng = kwargs["wait_ng"]
    try:
        res = main(**kwargs)
        if res:
            print(f"[bold][green]Complete.[/green][/bold]")
        else:
            print(f"[bold][red]Failed.[/red][/bold]")
            wait_time = wait_time_ng
    except Exception:
        wait_time = wait_time_ng
        # TODO: error process
        console.print_exception(
            show_locals=kwargs["trace"],
            suppress=suppress_lib_on_err,
        )
    finally:
        if wait_time > 0:
            waitkey(
                wait_time,
                output=lambda s: print(f"[bold][red]{s}[/red][/bold]", end=""),
            )


# endregion </CUI>
######################################

if __name__ == "__main__":
    cmd()
