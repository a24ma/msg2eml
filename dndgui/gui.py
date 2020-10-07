import PySimpleGUIQt as sg
import pathlib
import msg2eml
import traceback
import dndgui
from datetime import datetime
from logging import getLogger
from time import sleep

log = getLogger(__package__)
log_timeout = getLogger(__package__+"_timeout")


class PathInputHandler(object):
    def __init__(self):
        self.path_list = []
        self.valid_map = []
        self.targets = []
        self.path_count = 0
        self.valid_count = 0
        self.is_valid = False

    def receive_input(self, values):
        self.path_list = self._get_input_path_list(values)
        self._validate_path_list()
        self.targets = list(zip(
            range(self.path_count), self.path_list, self.valid_map))

    def _validate_path_list(self):
        self.path_count = len(self.path_list)
        self.valid_map = [self._check_valid(p) for p in self.path_list]
        self.valid_count = self.valid_map.count(True)
        self.is_valid = self.valid_count > 0

    def _get_input_path_list(self, values):
        raw_data = values["_INPUT_"]
        url_list = [d for d in raw_data.split("\n") if len(d) > 0]
        pathexp_list = [u.replace(r"file:///", r"") for u in url_list]
        return [pathlib.Path(p) for p in pathexp_list]

    def _check_valid(self, path):
        try:
            return path.exists() and path.suffix == ".msg"
        except OSError:
            log.warning(f"Invalid path: '{path}'")
            return False

    def convert(self, notify):
        conv = msg2eml.Converter()
        for index, path, valid in self.targets:
            log.info(f"Convert [{index}]: {path}")
            if valid:
                conv.read_msg(str(path.absolute()))
                # conv.show(show_body=False)
                conv.save_as_eml()
            notify(index, self.path_count)


class MainForm(object):
    def __init__(self):
        log.info("Form start.")
        sg.theme("DarkBlue3")
        self.def_font = ("Yu Gothic", 10)
        self.path_input_hdr = PathInputHandler()

    def _bd_status_kw(self, index, height):
        return {
            "key": f"_STATUS_{index}_",
            "size": (self.width, height),
            "font": self.def_font,
            "text_color": sg.theme_element_text_color(),
            "background_color": sg.theme_element_background_color(),
        }

    def _build(self):
        self.width = 576
        self.layout = [
            [sg.Text("Starting...", **self._bd_status_kw(1, 24))],
            [sg.Multiline("", **self._bd_status_kw(2, 128))],
            [sg.Text("", **self._bd_status_kw(3, 24))],
            [sg.Multiline(
                "D&D *.msg file(s) here", key="_INPUT_",
                enable_events=True, size=(self.width, 256),
                font=self.def_font, text_color="#666666")],
        ]
        self.header = "MSG=>EML Converter"
        self.window = sg.Window(self.header).Layout(self.layout).Finalize()
        self._on_running = True
        self._values = []
        self._exec_count = 0
        self._skip_next_event = False
        self._event_handlers = {
            "_INPUT_": self._on_dnd_event,
            "_TIMEOUT_": self._on_timeout_event,
        }

    @ property
    def inputbox(self):
        return self.window.FindElement("_INPUT_")

    @ inputbox.setter
    def inputbox(self, value):
        self.window.FindElement("_INPUT_").Update(value)

    def get_statustext(self, index):
        if index == 1:
            return self.window.FindElement(f"_STATUS_1_")
        elif index == 2:
            return self.window.FindElement(f"_STATUS_2_")
        elif index == 3:
            return self.window.FindElement(f"_STATUS_3_")
        else:
            raise IndexError()

    @ inputbox.setter
    def statustext(self, value):
        if len(value) == 3:
            self.get_statustext(1).Update(value[0])
            self.get_statustext(2).Update("\n".join(value[1]))
            self.get_statustext(3).Update(value[2])
        else:
            self.get_statustext(1).Update(value)
            self.get_statustext(2).Update("")
            self.get_statustext(3).Update("")

    def _on_exit_event(self, event, values):
        if event is None or event == "Exit":
            self._on_running = False
            return True
        return False

    def _default_event(self, event, values):
        pass

    def run(self):
        self._build()
        while True:
            event, values = self.window.Read(
                timeout=1000, timeout_key='_TIMEOUT_')
            log_timeout.debug(f"run() (event,values)=({event},{values})")
            if self._on_exit_event(event, values):
                break
            # Inputboxのテキスト更新による無限ループ防止
            if self._skip_next_event:
                self._skip_next_event = False
                continue
            handler = self._event_handlers.get(event, self._default_event)
            handler(event, values)
            self._exec_count += 1
            if not self._on_running or self._exec_count > 10:
                log.error("Too many events occur within short time.")
                log.error("Force to exit window.")
                break

    def _on_timeout_event(self, event, values):
        self._exec_count = 0

    def _on_dnd_event(self, event, values):
        self._skip_next_event = True
        self.inputbox = "D&D *.msg file(s) here"
        self.path_input_hdr.receive_input(values)
        if not self.path_input_hdr.is_valid:
            self.statustext = "ERROR: Not found *.msg files."
            return
        self._update_status_text("NOTE: Outlook の権限確認が出るので許可してください.")
        self.path_input_hdr.convert(self._receive_notification)

    def _update_status_text(self, msg):
        hdr = self.path_input_hdr
        self.statustext = (
            f"Found {hdr.path_count} files ({hdr.valid_count} are valid).",
            [self._view_stat(t) for t in hdr.targets],
            msg
        )
        self.window.refresh()

    def _view_stat(self, target):
        index, path, valid = target
        index_exp = f"[{index+1:2d}]"
        name = str(path.name)
        abbr_path = f"{name[:10]}...{name[-10:]}" if len(name) >= 25 else name
        status = "valid" if valid else "invalid"
        return f"{index_exp} {abbr_path:25}: {status}"

    def _receive_notification(self, index, n):
        now = datetime.now().strftime('%Y/%m/%d %H:%M:%S')
        self._update_status_text(
            f"{index+1}/{n} が完了しました。 @{now}")
