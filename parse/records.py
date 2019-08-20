import xlrd
import re
from functools import reduce
from xlrd import XLRDError
from utils.logger import logger
from parse.errror import ErrorNotValidExcel, with_error_stack


class Records(object):

    def __init__(self, filename: str):
        self.filename = filename
        self.split_time = re.compile(r'(\d{2}:\d{2})')
        self.length = 0
        self.record_type = 'floor_3'
        self.workbook = xlrd.open_workbook(self.filename)
        try:
            self.worksheet = self.workbook.sheet_by_name('考勤记录')
        except XLRDError as e:
            logger.error(with_error_stack(e))
            try:
                self.worksheet = self.workbook.sheet_by_index(0)
                logger.info({
                    'sheet': self.worksheet,
                })
            except Exception as e:
                logger.error(with_error_stack(e))
                raise ErrorNotValidExcel()
            self.record_type = 'floor_18'

    def get_sheet(self):
        return self.worksheet

    def parse(self) -> list:
        logger.info({
            'parse': self.record_type,
        })
        if self.record_type == 'floor_3':
            return self.parse_floor_3()
        return self.parse_floor_18()

    def parse_floor_3(self) -> list:
        self.length = self.worksheet.ncols
        result = []
        for i in range(self.worksheet.nrows):
            records = self.worksheet.row_values(i)
            if records[0] != '工 号:':
                continue

            meta = list(filter(lambda x: len(x) > 0, records))
            check_in = {
                'name': meta[3],
                'department': meta[5],
                'times': []
            }

            data = self.worksheet.row_values(i + 1)
            for index, datum in enumerate(data):
                data_time = self.split_time.split(datum)
                check_in['times'].append((
                    list(filter(lambda x: len(x) > 0, data_time)),
                    index + 1,
                ))
            result.append(check_in)

        if len(result) <= 0:
            raise ErrorNotValidExcel()

        return result

    def parse_floor_18(self) -> list:
        id_index = -1
        name_index = -1
        date_index = -1
        off_work_index = -1

        check_in = {}
        for index, record in enumerate(self.worksheet.row_values(0)):
            if record == '考勤号码':
                id_index = index
            if record == '姓名':
                name_index = index
            elif record == '日期':
                date_index = index
            elif record == '签退时间':
                off_work_index = index
        if id_index < 0 or name_index < 0 or date_index < 0 or off_work_index < 0:
            raise ErrorNotValidExcel()

        for i in range(1, self.worksheet.nrows):
            records = self.worksheet.row_values(i)
            name = records[name_index] + str(id_index)
            if name not in check_in:
                check_in[name] = {
                    'name': records[name_index],
                    'length': 0,
                    'times': []
                }

            check_in[name]['length'] += 1
            check_in[name]['times'].append((
                [records[off_work_index]],
                records[date_index].split('/')[-1]
            ))

        max_length_item = reduce(lambda x, y: x if x['length'] > y['length'] else y, check_in.values())
        self.length = max_length_item['length']

        result = []
        for key, value in check_in.items():
            result.append({
                'name': value['name'],
                'department': '',
                'times': value['times'],
            })
        return result

    def get_cols_len(self):
        return self.length
