from openpyxl import Workbook

from src.DataRow import DateHandler, Configurator

wb = Workbook()
cfg = Configurator()
for member in cfg.config.get('members', []):
    sheet = wb.create_sheet(member.get('name', 'x'))
    dr = DateHandler(sheet=sheet, config=cfg)
    dr.add_row(member.get('hours', [8, 8, 8, 8, 7]), member.get('stop', ''), member.get('start', '01/01'))
    dr.summary(member.get('name', 'x'))
wb.save('test.xlsx')
