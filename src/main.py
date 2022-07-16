import platform, os
if platform.system() == 'Windows':
    os.environ['KIVY_GL_BACKEND'] = 'angle_sdl2'

import csv
import typing
from datetime import datetime
from pathlib import Path

from kivy.app import App
from kivy.graphics import Rectangle, Color
from kivy.uix import textinput
from kivy.uix.button import Button
from kivy.uix.gridlayout import GridLayout
from kivy.uix.label import Label
from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.xml.constants import XLSM

MAIN_ZEN_FIELDS = [
    "date",
    "categoryName",
    "payee",
    "comment",
    "outcomeAccountName",
    "outcome",
    "outcomeCurrencyShortTitle",
    "incomeAccountName",
    "income",
    "incomeCurrencyShortTitle",
    "createdDate",
    "changedDate"
]

path = os.path.abspath("/Users/gmpolunin/Downloads/2022-07-09T21-32-26.xlsm")


class CSVSniffer:
    def __init__(self, file_startswith: str, main_path: str, fields: typing.List):
        self.__is_file: bool = os.path.isfile(main_path)
        self.__start_row_flag: bool = False
        self.__file_startswith: str = file_startswith
        self.__main_path: str = Path(main_path).parent.absolute() if self.__is_file else main_path
        self.__fields: typing.List = fields

    def get_files_with_lines_for_set_dir(self) -> typing.List:
        try:
            return list(self.__main_dir_sniffer())
        except Exception as e:
            raise CSVSnifferException(exc=e)

    @property
    def is_file(self) -> bool:
        return self.__is_file

    def __main_dir_sniffer(self) -> typing.Generator:
        # if not os.path.exists(self.__main_path) or (not os.path.isdir(self.__main_path) and not self.__is_file):
        if not os.path.exists(self.__main_path) or not self.__is_file:
            raise NotADirectoryError
        for dir_path, _, files_list in os.walk(self.__main_path):
            yield from self.__files_sniffer_by_pattern(dir_path, files_list)

    def __files_sniffer_by_pattern(self, curr_dir_path: str, files_list: typing.List[str]) -> typing.Generator:
        for file_name in files_list:
            if file_name.startswith(self.__file_startswith):
                self.__start_row_flag = False
                lines = self.__csv_file_parser(os.path.join(curr_dir_path, file_name))
                yield {file_name: list(lines)}

    def __csv_file_parser(self, abs_file_path: str) -> typing.Generator:
        with open(abs_file_path, "r", encoding="utf-8") as csv_file:
            for csv_line in csv.reader(csv_file):
                if csv_line == self.__fields:
                    self.__start_row_flag = True
                if self.__start_row_flag is True:
                    yield csv_line


class CSVSnifferException(Exception):
    def __init__(self, exc: typing.Optional[Exception] = None):
        self.exc = exc


class ExcelWorker:
    __workbook: typing.Optional[Workbook] = None

    def __init__(
            self,
            workbook_name: str,
            workbook_extension: str = ".xlsx",
            want_cleared: bool = True,
            sheets_to_create: typing.Tuple = (),
    ):
        self.__workbook_name: str = workbook_name
        self.__workbook_extension: str = workbook_extension
        self.__full_workbook_name: str = self.__workbook_name + self.__workbook_extension
        self.__want_cleared: bool = want_cleared
        self.__sheets_to_create: typing.Tuple = sheets_to_create
        self.__load_or_create_wb(self.__want_cleared, self.__sheets_to_create)

    def fill_workbook(self, all_data: typing.Dict[str, typing.List]) -> Workbook:
        try:
            for k, v in all_data.items():
                self.__create_and_fill_ws(sheet_name=k, data_to_fill=v)
            return self.__save_and_close_wb()
        except Exception as e:
            raise ExcelWorkerException(exc=e)

    @property
    def full_workbook_name(self) -> str:
        return self.__full_workbook_name

    def __load_or_create_wb(self, want_cleared: bool, sheets_to_create: typing.Tuple):
        if want_cleared:
            self.__workbook = Workbook()
            self.__workbook.template = XLSM
            self.__workbook.remove(worksheet=self.__workbook.active)

            if sheets_to_create:
                for sc in sheets_to_create:
                    self.__workbook.create_sheet(title=sc)
        else:
            self.__workbook = load_workbook(self.__full_workbook_name, keep_vba=True)

    def __save_and_close_wb(self):
        self.__workbook.save(self.__full_workbook_name)
        self.__workbook.close()
        return self.__workbook

    def __rename_and_pick_first_ws(self, sheet_name: str) -> Worksheet:
        ws: Worksheet = self.__workbook.worksheets[0]
        ws.title = sheet_name
        return ws

    def __create_named_ws_in_wb(self, sheet_name: str) -> Worksheet:
        return self.__workbook.create_sheet(title=sheet_name)

    @staticmethod
    def __ws_append_with_data(ws: Worksheet, data: typing.List):
        for data_row in data:
            ws.append(data_row)

    def __create_and_fill_ws(self, sheet_name: str, data_to_fill: typing.List):
        ws = self.__create_named_ws_in_wb(sheet_name=sheet_name[:30])
        self.__ws_append_with_data(ws=ws, data=data_to_fill)


class ExcelWorkerException(Exception):
    def __init__(self, exc: typing.Optional[Exception] = None):
        self.exc = exc


class ZenMoneyJob:
    def __init__(self, dir_path: str):
        self.__dir_path: str = dir_path

        self.__file_startswith: str = "zen_"
        self.__csv_hunter_fields: typing.List[str] = MAIN_ZEN_FIELDS

        self.__workbook_name: str = datetime.isoformat(datetime.now())[:19].replace(":", "-")

        self.__csv_hunter: CSVSniffer = CSVSniffer(
            file_startswith=self.__file_startswith,
            main_path=self.__dir_path,
            fields=self.__csv_hunter_fields,
        )
        if self.__csv_hunter.is_file:
            self.__excel_worker: ExcelWorker = ExcelWorker(
                workbook_name=self.__dir_path,
                workbook_extension="",
                want_cleared=False,
            )
        else:
            self.__excel_worker: ExcelWorker = ExcelWorker(
                workbook_name=f"{self.__dir_path}/{self.__workbook_name}",
                workbook_extension=".xlsm",
                sheets_to_create=("Total", "Config"),
            )

    def __prepare_data_with_dir_path(self) -> typing.List[typing.Dict]:
        return list(self.__csv_hunter.get_files_with_lines_for_set_dir())

    def __paste_prepared_data_at_workbook(self, prepared_data: typing.Dict[str, typing.List]) -> Workbook:
        return self.__excel_worker.fill_workbook(all_data=prepared_data)

    def find_csv_files_and_paste_lines_to_excel(self):
        try:
            prepared_data = self.__prepare_data_with_dir_path()
            for pd in prepared_data:
                self.__paste_prepared_data_at_workbook(prepared_data=pd)
            return self.__excel_worker.full_workbook_name
        except CSVSnifferException as e:
            raise ZenMoneyJobException(msg="Файловая ошибка. Директория/файл не найден.", exc=e.exc)
        except ExcelWorkerException as e:
            raise ZenMoneyJobException(msg="Внутренняя ошибка работы с Excel.", exc=e.exc)


class ZenMoneyJobException(Exception):
    def __init__(self, exc: typing.Optional[Exception] = None, msg: typing.Optional[str] = None):
        self.exc = exc
        self.msg = msg


class TextInput(textinput.TextInput):
    def __init__(self, **kwargs):
        super(TextInput, self).__init__(**kwargs)
        self.padding_x = [
            self.center[0] - self._get_text_width(max(self._lines, key=len), self.tab_width, self._label_cached) / 2.0,
            0,
        ] if self.text else [self.center[0], 0]
        self.padding_y = [self.height / 2.0 - (self.line_height / 2.0) * len(self._lines), 0]


VIOLET = .20, .06, .31, 1
YELLOW = .988, .725, .074, 1


class ZenMoneyLayout(GridLayout):
    def __init__(self, **kwargs):
        super(ZenMoneyLayout, self).__init__(**kwargs)

        with self.canvas.before:
            Color(*YELLOW, mode="rgba")
            self.rect = Rectangle(pos=self.pos, size=self.size)
        self.bind(size=self.update_rect)

        self.cols = 1
        self.height = self.minimum_height

        self.directory_input = GridLayout(
            cols=2,
            size_hint_y=.3,
        )
        self.directory_input.add_widget(
            Label(
                text="Путь до директории/файла:",
                color=VIOLET,
                bold=True,
                text_size=(None, None),
                font_size="20sp",
            )
        )
        self.directory = TextInput(
            multiline=True,
            hint_text="User/example/path/to/directory",
            is_focusable=True,
        )
        self.directory_input.add_widget(self.directory)
        self.add_widget(self.directory_input)

        self.submit = Button(
            text="Выполнить",
            background_normal="",
            background_color=VIOLET,
            size_hint_y=.3,
            bold=True,
            text_size=(None, None),
            font_size="20sp",
            color=YELLOW,
        )

        self.submit.bind(on_press=self.press)
        self.add_widget(self.submit)

        self.error = Label(
            bold=True,
            text_size=(None, None),
            font_size="20sp",
            padding=[100, 100],
        )
        self.add_widget(self.error)

    def press(self, instance):
        try:
            zmj = ZenMoneyJob(
                dir_path=self.directory.text,
            )
            new_excel = zmj.find_csv_files_and_paste_lines_to_excel()
            self.error.color = "green"
            self.error.text = f"Успешно:\n{new_excel}"
        except ZenMoneyJobException as e:
            self.error.color = "red"
            self.error.text = f"{e.msg}\n{e.exc}"

    def update_rect(self, *args):
        self.rect.pos = self.pos
        self.rect.size = self.size


class ZenMoneyApp(App):
    icon = "zen_ico.png"
    title = "Zen Money 💰"

    def build(self):
        return ZenMoneyLayout()


if __name__ == "__main__":
    ZenMoneyApp().run()
