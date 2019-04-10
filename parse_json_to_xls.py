import json
import datetime
import xlwt


class SaveExcel:
    def __init__(self, _file):
        with open(_file, encoding='utf8') as json_file:
            self.data = json.load(json_file)
            self.data_record_excel()

    def validate(self, value):
        return str(value) or 'None'

    def data_record_excel(self):
        wb = xlwt.Workbook()
        ws = wb.add_sheet('Parsing json data', cell_overwrite_ok=True)
        row = 0
        for i, key in enumerate(self.data):
            ws.write(i, row, key)
            if key == 'description':
                ws.write(i, row + 1, self.validate(' '.join(self.data[key].split())))
            else:
                ws.write(i, row + 1, self.validate(self.data[key]))
        wb.save('docs/' + str(datetime.datetime.now().strftime('%H%M%S')) + '.xls')


file = SaveExcel('B07JG7ZZZ7.json.txt')
