from datetime import datetime, date, timedelta
from typing import Dict, Iterator, Union
from calendar import Calendar
import holidays
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.dimensions import RowDimension
from openpyxl.styles.builtins import styles
from yaml import load, SafeLoader
from pathlib import Path


class DateRow():
    _worksheet: Worksheet
    _config: Dict
    _calendar: Calendar
    _holidays: holidays
    _special_days: Dict

    def __init__(self, sheet: Union[Worksheet, None]) -> None:
        if sheet:
            self._worksheet = sheet
        with (Path.cwd() / 'config.yaml').open() as config_file:
            self._config = load(config_file.buffer, SafeLoader)

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
                self._special_days[key_date.isoformat()] = {
                    'name': name.get('name', ''),
                    'color': name.get('color', self._config.get('color_holiday', '#000000'))}

    def year_iterator(self) -> Iterator:
        day = date(year=self._config.get('year', date.today().year), month=1, day=1)
        for i in range(0, (date(year=self._config.get('year', date.today().year) + 1, month=1, day=1) -
                           date(year=self._config.get('year', date.today().year), month=1, day=1)).days):
            if day in self._holidays:
                yield [day, self._config.get('holiday', 'holiday'), self._config.get('color_holiday', '#000000')]
            elif day.weekday() > 4:
                yield [day, self._config.get('weekend', 'weekend'), self._config.get('color_weekend', '#000000')]
            elif day.isoformat() in self._special_days.keys():
                yield [day, self._special_days[day.isoformat()]['name'],
                       self._special_days[day.isoformat()]['color']]
            else:
                yield [day, '', '']
            day += timedelta(days=1)

    def add_row(self) -> int:
        header = [('A1', 'Datum'), ('B1', 'Wochentag'), ('C1', 'Soll'), ('D1', 'Ist'), ('E1', 'Saldo'),
                  ('F1', 'Abwesenheit')]
        for cell in header:
            self._worksheet[cell[0]] = cell[1]
            self._worksheet[cell[0]].style = styles['Title']
            self._worksheet.column_dimensions[cell[0][0:1]].width = 20

        for index, day in enumerate(self.year_iterator()):
            for c in range(1, 7):

                if day[2] == self._config.get('color_holiday', '#000000'):
                    self._worksheet.cell(row=index + 2, column=c).style = styles['Accent1']
                elif day[2] == self._config.get('color_weekend', '#000000'):
                    self._worksheet.cell(row=index + 2, column=c).style = styles['Accent2']
                elif day[2]:
                    self._worksheet.cell(row=index + 2, column=c).style = styles['Accent6']
                else:
                    self._worksheet.cell(row=index + 2, column=3, value=8.00)
                    self._worksheet.cell(
                        row=index + 2, column=5,
                        value=f'=IF(AND(A{index + 2}<TODAY()-2,C{index + 2}<>""),D{index + 2}-C{index + 2},"")')
                    # dv.add(self._worksheet.cell(row=index + 2, column=6))

            self._worksheet.cell(row=index + 2, column=1, value=day[0])
            self._worksheet.cell(row=index + 2, column=2, value=day[0].strftime('%a'))
            self._worksheet.row_dimensions[index + 2].height = 25

        return self._worksheet.max_row
