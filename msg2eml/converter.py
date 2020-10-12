# -*- coding: utf-8 -*-

import os
import re
import sys
import time
import datetime
import win32com.client
import pathlib
import mimetypes
import unicodedata
from email import encoders
from email import generator
from email.mime.multipart import MIMEMultipart
from email.mime.audio import MIMEAudio
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText
from email.headerregistry import Address
from email.header import Header
from email.utils import format_datetime
from logging import getLogger

log = getLogger(__package__)


class Converter(object):
    @staticmethod
    def addr_id(addr_entry):
        ret = re.sub("[^a-zA-Z0-9]", "", addr_entry.address)
        return ret[:4]

    @staticmethod
    def create_addr(addr_entry):
        name = Header(addr_entry.name, "utf-8").encode()
        addr = addr_entry.address
        return f"{name} <{addr}>"

    @staticmethod
    def get_filtered_addrs(recipients, filter_exp):
        return ", ".join([
            Converter.create_addr(r) for r in recipients
            if r.name in filter_exp])

    @staticmethod
    def convert_to_jst_exp(dt):
        timezone = datetime.timezone(datetime.timedelta(hours=9))  # JSTのみ想定
        return format_datetime(dt.replace(tzinfo=timezone))

    @staticmethod
    def get_body_format_exp(format_type):
        olbodyformat_dict = {
            1: "plain",
            2: "html",
            3: "html"
        }
        return olbodyformat_dict.get(format_type, "plain")

    @staticmethod
    def simplify_name(name, max_len=20):
        name = unicodedata.normalize("NFKD", name)
        log.debug(f"simplify_name()  name: {name}")
        parts = name.rsplit(".", 1)
        suffix = ""
        if len(parts) > 1:
            name = parts[0]
            suffix = "." + parts[1]
        name = re.sub("re: ?", "", name, flags=re.I)
        name = re.sub("fwd?: ?", "", name, flags=re.I)
        name = name.strip()
        name = name.translate(str.maketrans({
            "?": "？",
            "!": "！",
            ":": "：",
            "<": "＜",
            ">": "＞",
            '"': "'",
            "*": "＊",
            "　": "_",
            " ": "_",
            "|": "_",
            "/": "／",
            "\\": "_",
            "(": "（",
            ")": "）",
        }))
        if suffix is None:
            suffix = ""
        return name[:max_len] + suffix

    def __init__(self):
        self.outlook = win32com.client.Dispatch(
            "Outlook.Application").GetNamespace("MAPI")
        self.path = None

    def read_msg(self, filepath):
        srcpath = pathlib.Path(filepath).absolute()
        log.debug(f"readmsg() filepath: '{filepath}'")
        log.debug(f"readmsg()  srcpath: '{srcpath}'")
        mail = self.outlook.OpenSharedItem(srcpath)
        self.subject = mail.subject
        sender = mail.sender
        self.sender = self.create_addr(sender)
        self.sender_id = self.addr_id(sender)
        recipients = mail.recipients
        self.recipients_to = self.get_filtered_addrs(recipients, mail.to)
        self.recipients_cc = self.get_filtered_addrs(recipients, mail.cc)
        self.bodyformat = self.get_body_format_exp(mail.bodyformat)
        self.body = mail.htmlbody if self.bodyformat == "html" else mail.body
        self.attachments = mail.attachments
        self._init_date(mail.receivedtime)
        self._init_path(srcpath)

    def _init_date(self, dt):
        timezone = datetime.timezone(datetime.timedelta(hours=9))  # JSTのみ想定
        self.date = dt.replace(tzinfo=timezone)
        self.date_exp = format_datetime(self.date)

    def _init_path(self, srcpath):
        parent = srcpath.parent
        # ディレクトリ中のファイルのAOL:Consult序数から最大のものを探す。
        i = 0
        for f in parent.iterdir():
            # if not f.name.endswith(".eml"):
            #     continue
            m = re.match(r"c(\d\d)", f.name)
            if m is not None:
                index = int(m.groups()[0])
                i = max(i, index)
            f.name.startswith("")
        self.aolc_index = i + 1
        d = self.date.strftime("%Y-%m%d_%H-%M")
        simple_sub = self.simplify_name(self.subject)
        self.mail_id = f"{d}_{self.sender_id}_{simple_sub}"
        filename = f"c{self.aolc_index:02d}_{self.mail_id}.eml"
        self.path = parent.joinpath(filename)

    def _attachment_filepath(self, att, index):
        simple_fname = self.simplify_name(att.filename)
        prefix = f"c{self.aolc_index+index:02d}_{self.mail_id}"
        filename = f"{prefix}_[{index:02d}]{simple_fname}"
        return self.path.parent.joinpath(filename)

    def get_description(self, show_body=False):
        if self.path is None:
            raise Exception("read msg first")
        msg = ""
        msg += f"Subject: {self.subject}\n"
        msg += f"   From: {self.sender}\n"
        msg += f"     To: {self.recipients_to}\n"
        msg += f"     CC: {self.recipients_cc}\n"
        msg += f"   Date: {self.date_exp}\n"
        msg += f" Format: {self.bodyformat}\n"
        msg += "添付ファイル: \n"
        i = 0
        for a in self.attachments:
            i += 1
            msg += f"  {i:4d}. {a.filename}\n"
        if show_body:
            msg += (f"""
----- 本文 -----
{self.body}
---------------
""")
        return msg

    def save_as_eml(self):
        if self.path is None:
            raise Exception("read msg first")
        msg = MIMEMultipart()
        msg["Subject"] = self.subject
        msg["From"] = self.sender
        msg["To"] = self.recipients_to
        msg["Cc"] = self.recipients_cc
        msg["Date"] = self.date_exp
        msg.attach(MIMEText(self.body, self.bodyformat))
        i = 0
        for a in self.attachments:
            i += 1
            path = self._attachment_filepath(a, i)
            log.debug(f"save_as_eml() Save '{path}'")
            a.SaveAsFile(str(path))
            msg.attach(self._read_attachment(path, a.filename))
        with open(str(self.path), "w") as f:
            gen = generator.Generator(f)
            gen.flatten(msg)

    def _read_attachment(self, pathobj, name):
        log.debug(f"read_attachment() path: '{pathobj}'")
        log.debug(f"read_attachment() name: '{name}'")
        path = str(pathobj)
        ctype, encoding = mimetypes.guess_type(path)
        log.debug(f"read_attachment() ctype: '{ctype}'")
        log.debug(f"read_attachment() encoding: '{encoding}'")
        if ctype is None or encoding is not None:
            log.warning(f"Cannot guess the type of '{path}', "
                + "or it is encoded (compressed).")
            ctype = 'application/octet-stream'
        maintype, subtype = ctype.split('/', 1)
        if maintype == 'text':
            with open(path) as fp:
                # Note: we should handle calculating the charset
                att = MIMEText(fp.read(), _subtype=subtype)
        elif maintype == 'image':
            with open(path, 'rb') as fp:
                att = MIMEImage(fp.read(), _subtype=subtype)
        elif maintype == 'audio':
            with open(path, 'rb') as fp:
                att = MIMEAudio(fp.read(), _subtype=subtype)
        elif maintype == 'application':
            with open(path, 'rb') as fp:
                att = MIMEApplication(fp.read(), _subtype=subtype)
        else:
            with open(path, 'rb') as fp:
                att = MIMEBase(maintype, subtype)
                att.set_payload(fp.read())
            encoders.encode_base64(att)
        # Set the filename parameter
        encoded_name = Header(name, "utf-8").encode()
        att.add_header('Content-Disposition',
                       'attachment', filename=encoded_name)
        return att
