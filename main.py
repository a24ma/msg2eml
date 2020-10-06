import msg2eml
import dndgui
import traceback
import sys
from logutil import *

log = setup_logger(WARNING)
# getLogger("dndgui").setLevel(NOTSET)
# getLogger("msg2eml").setLevel(NOTSET)


def cuimode(args):
    filename = args[1]
    log.info(f"Start to convert {filename}...")
    print(f"NOTE: Please check outlook application for authentication.")
    conv = msg2eml.Converter()
    conv.read_msg(filename)
    print(conv.get_description(show_body=False))
    conv.save_as_eml()
    log.info(f"Completed.")


def guimode():
    form = dndgui.MainForm()
    form.run()


if __name__ == "__main__":
    stop_on_error = False
    try:
        args = sys.argv
        print(args)
        if "-d" in args:
            args.remove("-d")
            log = setup_logger(DEBUG)
            log.debug("Debug mode")
            stop_on_error = True
            args = args[1:]
        if len(args) >= 2:
            cuimode(args)
        else:
            guimode()
    except Exception as e:
        log.error(f"""
===== [ERROR] =====
{traceback.format_exc()}
===================

""")
        if stop_on_error:
            input("Hit enter...")
