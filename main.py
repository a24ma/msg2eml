#!/usr/bin/env python
# -*- coding: utf-8 -*-

######################################
# region    <IMPORT>

import pathlib
import tkinter as tk
import pyautogui
import PIL
import msg2eml

# pip install tkinterdnd2 pyautogui

# endregion </IMPORT>
######################################
# region    <USER_SETTINGS>

log_package_name = "msg2eml"  # Avoid __package__==None on main
sub_lib = [tk, pyautogui, PIL]

# endregion </USER_SETTINGS>
######################################
# region    <UTILITY>

import click
import logging
import subprocess
from rich import print
from rich.console import Console
from rich.logging import RichHandler

log_trace_name = "trace." + log_package_name  # logger with different logLevel
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

    log_trace.addHandler(log_hdl)
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


def main(**kwargs):
    runner = msg2eml.Msg2eml("./mat/bg.png", "./mat/ok.png", "./mat/ng.png")
    runner.run()
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
