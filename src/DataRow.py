from datetime import datetime, date, timedelta
from typing import Dict, Iterator, Union, List
from calendar import Calendar
import holidays
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.styles.builtins import styles
from yaml import load, SafeLoader
from pathlib import Path

date_types = ['normal', 'weekend', 'holiday', 'special']


class DataRow:
    date: date
    type: int  # key of data_types
    name: str


class Configurator:
    def __init__(self):
        with (Path.cwd() / 'config.yaml').open() as config_file:
            self.config: Dict = load(config_file.buffer, SafeLoader)


class DateHandler:
    _worksheet: Worksheet
    _config: Dict
    _calendar: Calendar
    _holidays: holidays
    _special_days: Dict

    def __init__(self, sheet: Union[Worksheet, None], config: Configurator) -> None:
        if sheet:
            self._worksheet = sheet
        self._config = config.config
        self.font = Font(size=16)
        self.side = Side(style='medium', color='00000000')

        self._calendar = Calendar(firstweekday=0)
        self._holidays: Dict = holidays.country_holidays(
            country=self._config.get('country', 'DE'),
            subdiv=self._config.get('subdiv', 'BY'),
            years=self._config.get('year', datetime.today().year)
        )
        self._special_days = {}
        for name in self._config.get('holidays', [{'name': ''}]):
            for date_record in name.get('dates', []):
                month_day = datetime.strptime(date_record, self._config.get('format', '%d/%m'))
                key_date = date(year=self._config.get('year', datetime.today().year),
                                month=month_day.month, day=month_day.day)
                self._special_days[key_date.isoformat()] = name.get('name', '')

    def year_iterator(self) -> Iterator[DataRow]:
        day = date(year=self._config.get('year', date.today().year), month=1, day=1)
        return_row = DataRow()
        for i in range(0, (date(year=self._config.get('year', date.today().year) + 1, month=1, day=1) -
                           date(year=self._config.get('year', date.today().year), month=1, day=1)).days):
            if day in self._holidays:
                return_row.date, return_row.type, return_row.name = day, 2, self._config.get(
                    'holiday', 'Holiday')
            elif day.weekday() > 4:
                return_row.date, return_row.type, return_row.name = day, 1, ''
            elif day.isoformat() in self._special_days.keys():
                return_row.date, return_row.type, return_row.name = day, 3, self._special_days[
                    day.isoformat()]
            else:
                return_row.date, return_row.type = day, 0
            yield return_row
            day += timedelta(days=1)

    def add_row(self, hours: List, stop: str = '') -> int:
        if len(hours) < 5:
            raise Exception('you have to provide hours for each weekday')

        dv = DataValidation(type='list', formula1='"krank,urlaub,kindkrank,fortbildung"')
        self._worksheet.add_data_validation(dv)
        header = [('A1', 'Datum'), ('B1', 'Wochentag'), ('C1', 'Soll'), ('D1', 'Ist'), ('E1', 'Saldo'),
                  ('F1', 'Abwesenheit')]
        for cell in header:
            self._worksheet[cell[0]] = cell[1]
            self._worksheet[cell[0]].style = styles['Title']
            self._worksheet.column_dimensions[cell[0][0:1]].width = 20
        if stop:
            month_day = datetime.strptime(stop, self._config.get('format', '%d/%m'))
        else:
            month_day = date(year=self._config.get('year', datetime.today().year), month=12, day=31)
        for index, day_row in enumerate(self.year_iterator()):
            if day_row.date > date(year=self._config.get('year', datetime.today().year),
                                   month=month_day.month,
                                   day=month_day.day):
                return 1

            style = ''
            if day_row.type == 1:
                style = '40 % - Accent1'
            if day_row.type == 2:
                style = '40 % - Accent2'
            if day_row.type == 3:
                style = '40 % - Accent6'
            if style:
                for c in range(1, 7):
                    self._worksheet.cell(row=index + 2, column=c).style = styles[style]
                self._worksheet.cell(row=index + 2, column=6, value=day_row.name)
            else:
                self._worksheet.cell(row=index + 2, column=3, value=hours[day_row.date.weekday()])
                self._worksheet.cell(
                    row=index + 2, column=5,
                    value=f'=IF(AND(A{index + 2}<TODAY()-2,C{index + 2}<>"",F{index + 2}=""),'
                          f'IF(C{index + 2}=0,D{index + 2}*1.5,D{index + 2}-C{index + 2}),"")')
                dv.add(self._worksheet.cell(row=index + 2, column=6))

            self._worksheet.cell(row=index + 2, column=1, value=day_row.date)
            self._worksheet.cell(row=index + 2, column=1).number_format = 'd/m'
            self._worksheet.cell(row=index + 2, column=2, value=day_row.date.strftime('%a'))
            self._worksheet.row_dimensions[index + 2].height = 25
            for c in range(1, 7):
                self._worksheet.cell(row=index + 2, column=c).font = self.font
                self._worksheet.cell(row=index + 2, column=c).alignment = Alignment(horizontal="center")

        return self._worksheet.max_row

    def summary(self, name: str):

        for i in ('G', 'H', 'I', 'J', 'K', 'L', 'M', 'N'):
            self._worksheet.column_dimensions[i].width = 20
        # header for name
        self._worksheet.merge_cells(start_row=8, start_column=8, end_row=8, end_column=16)
        self._worksheet.cell(row=8, column=8, value=f'Zusammenfassung {name}')
        self._worksheet.cell(row=8, column=8).style = styles['Output']
        self._worksheet.cell(row=8, column=8).alignment = Alignment(horizontal="center")
        self._worksheet.cell(row=8, column=8).font = self.font
        # self._worksheet.cell(row=8, column=8).border = Border(left=self.side, right=self.side)

        # header holiday
        self._worksheet.merge_cells(start_row=10, start_column=8, end_row=10, end_column=12)
        self._worksheet.cell(row=10, column=8, value=f'Urlaubstage')
        self._worksheet.cell(row=10, column=8).style = styles['40 % - Accent3']
        self._worksheet.cell(row=10, column=8).alignment = Alignment(horizontal="center")
        self._worksheet.cell(row=10, column=8).font = self.font

        # header overtime
        self._worksheet.merge_cells(start_row=10, start_column=14, end_row=10, end_column=16)
        self._worksheet.cell(row=10, column=14, value=f'Überstunden')
        self._worksheet.cell(row=10, column=14).style = styles['40 % - Accent3']
        self._worksheet.cell(row=10, column=14).alignment = Alignment(horizontal="center")
        self._worksheet.cell(row=10, column=14).font = self.font

        # holiday cells
        self._worksheet.cell(row=11, column=8, value=f"Rest {self._config.get('year', datetime.today().year) - 1}")
        self._worksheet.cell(row=11, column=8).style = styles['40 % - Accent3']
        self._worksheet.cell(row=11, column=8).alignment = Alignment(horizontal="center")
        self._worksheet.cell(row=11, column=8).font = self.font
        self._worksheet.cell(row=11, column=9, value=f"Soll {self._config.get('year', datetime.today().year)}")
        self._worksheet.cell(row=11, column=9).style = styles['40 % - Accent3']
        self._worksheet.cell(row=11, column=9).alignment = Alignment(horizontal="center")
        self._worksheet.cell(row=11, column=9).font = self.font
        self._worksheet.cell(row=11, column=10, value=f"Summe")
        self._worksheet.cell(row=11, column=10).style = styles['40 % - Accent3']
        self._worksheet.cell(row=11, column=10).alignment = Alignment(horizontal="center")
        self._worksheet.cell(row=11, column=10).font = self.font
        self._worksheet.cell(row=11, column=11, value=f"Genommen")
        self._worksheet.cell(row=11, column=11).style = styles['40 % - Accent3']
        self._worksheet.cell(row=11, column=11).alignment = Alignment(horizontal="center")
        self._worksheet.cell(row=11, column=11).font = self.font
        self._worksheet.cell(row=11, column=12, value=f"Offen")
        self._worksheet.cell(row=11, column=12).style = styles['40 % - Accent3']
        self._worksheet.cell(row=11, column=12).alignment = Alignment(horizontal="center")
        self._worksheet.cell(row=11, column=12).font = self.font
