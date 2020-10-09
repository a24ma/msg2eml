# -*- coding: utf-8 -*-

import unicodedata
from logging import getLogger
from logging import StreamHandler
from logging import Formatter
from logging import NOTSET
from logging import DEBUG

def setup_logger(level=NOTSET):
    logger = getLogger()
    sh = StreamHandler()
    formatter = Formatter('%(asctime)s [%(levelname)s] %(name)s > %(message)s')
    sh.setFormatter(formatter)
    logger.addHandler(sh)
    logger.setLevel(level)
    sh.setLevel(level)
    return logger

    
def diff_str(s1, s2):
    w = unicodedata.east_asian_width
    out1 = ""
    out2 = ""
    diff = ""
    for cs in zip(s1, s2):
        c1, c2 = cs
        w1, w2 = (2 if w(c) in "WF" else 1 for c in cs)
        if w1 == w2:
            out1 += f"{c1}"
            out2 += f"{c2}"
            if c1 == c2:
                diff += "  "[:w1]
            else:
                out1 += f"[{ord(c1):04x}]"
                out2 += f"[{ord(c2):04x}]"
                diff += "^-"[:w1] + "------"
        else:
            out1 += f"{c1:{w2}}[{ord(c1):04x}]"
            out2 += f"{c2:{w1}}[{ord(c2):04x}]"
            diff += "^-------"
    return out1, out2, diff

def out_diff(s1, s2):
    return ("""
*   s1: %s
*   s2: %s
* diff: %s
""" % diff_str(s1, s2))
