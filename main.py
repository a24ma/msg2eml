import os
import re
import sys
import datetime
import win32com.client
import traceback
import pathlib
from email import generator
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.headerregistry import Address
from email.header import Header
from email.utils import format_datetime


class Msg2EmlConverter(object):
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
            Msg2EmlConverter.create_addr(r) for r in recipients
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
        parts = name.rsplit(".", 2)
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
            "\\": "_",
        }))
        if suffix is None: suffix = ""
        return name[:max_len] + suffix

    def __init__(self):
        self.path = None

    def read_msg(self, filepath):
        srcpath = pathlib.Path(filepath).absolute()
        outlook = win32com.client.Dispatch(
            "Outlook.Application").GetNamespace("MAPI")
        mail = outlook.OpenSharedItem(str(srcpath))
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
        d = self.date.strftime("%Y-%m-%d_%H-%M")
        simple_sub = self.simplify_name(self.subject)
        self.mail_id = f"{d}_{self.sender_id}_{simple_sub}"
        filename = f"c{self.aolc_index:02d}_{self.mail_id}.eml"
        self.path = parent.joinpath(filename)

    def _attachment_filename(self, att, index):
        simple_fname = self.simplify_name(att.filename)
        prefix = f"c{self.aolc_index+index:02d}_{self.mail_id}"
        filename = f"{prefix}_[{index:02d}]{simple_fname}"
        return str(self.path.parent.joinpath(filename))

    def show(self, show_body=False):
        if self.path is None: raise Exception("read msg first")
        print("Subject: ", self.subject)
        print("   From: ", self.sender)
        print("     To: ", self.recipients_to)
        print("     CC: ", self.recipients_cc)
        print("   Date: ", self.date_exp)
        print(" Format: ", self.bodyformat)
        print("添付ファイル: ")
        i = 0
        for a in self.attachments:
            i += 1
            print(f"{i:4d}. {a.filename}")
        if show_body:
            print(f"""
----- 本文 -----
{self.body}
---------------
""")

    def save_as_eml(self):
        if self.path is None: raise Exception("read msg first")
        msg = MIMEMultipart("alternative")
        msg["Subject"] = self.subject
        msg["From"] = self.sender
        msg["To"] = self.recipients_to
        msg["Cc"] = self.recipients_cc
        msg["Date"] = self.date_exp
        msg.attach(MIMEText(self.body, self.bodyformat))
        with open(str(self.path), "w") as f:
            gen = generator.Generator(f)
            gen.flatten(msg)
        i = 0
        for a in self.attachments:
            i += 1
            a.SaveAsFile(self._attachment_filename(a, i))


def main(filename):
    conv = Msg2EmlConverter()
    conv.read_msg(filename)
    conv.show(show_body=False)
    conv.save_as_eml()


if __name__ == "__main__":
    try:
        args = sys.argv
        if len(args) >= 2:
            filename = args[1]
            print(f"[INFO] Start to convert {filename}...")
            print(f"[INFO] NOTE: Please check outlook application for authentication.")
            main(filename)
            print(f"[INFO] Completed.")
            print()
        else:
            print(f"[INFO] USAGE: {args[0]} <email.msg>")
    except Exception as e:
        print(f"""
===== [ERROR] =====
{traceback.format_exc()}
===================

""")
