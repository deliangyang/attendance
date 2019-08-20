import xlwt
import re
from xlrd.sheet import Sheet
from utils.logger import logger
from parse.errror import with_error_stack


class Write(object):

    def __init__(self, filename: str, data: iter, length: int, origin_sheet: Sheet):
        self.data = data
        self.filename = filename
        self.length = length
        self.workbook = xlwt.Workbook(encoding="utf-8")
        self.origin_sheet = origin_sheet
        self.split_time = re.compile(r'(\d{2}:\d{2})')

        alignment = xlwt.Alignment()
        alignment.wrap = 1
        alignment.horz = 0x02
        self.style2 = xlwt.XFStyle()
        self.style2.alignment = alignment

    def save(self):
        logger.info('write origin data')
        self.write_origin_data()
        logger.info('write overtime')
        self.write_overtime()
        logger.info('write take taxi')
        self.write_take_taxi()
        self.workbook.save(self.filename)

    def write_origin_data(self):
        sheet = self.workbook.add_sheet('考勤记录', cell_overwrite_ok=True)
        for i in range(self.origin_sheet.nrows):
            record = self.origin_sheet.row_values(i)
            for j, rd in enumerate(record):
                try:
                    if rd and self.split_time.findall(str(rd)) is not None:
                        items = self.split_time.split(str(rd))
                        rd = "\n".join(
                            filter(lambda x: len(x) > 0, items)
                        )
                except Exception as e:
                    logger.error(with_error_stack(e))
                sheet.write(i, j, rd, self.style2)

    def write_overtime(self):
        sheet = self.workbook.add_sheet('加班', cell_overwrite_ok=True)
        sheet.write(0, 0, '姓名')
        sheet.write(0, 1, '部门')
        sheet.write(0, 2, '总次数')
        for i in range(self.length):
            sheet.write(0, i + 3, str(i + 1))

        for index, datum in enumerate(self.data):
            index += 1
            sheet.write(index, 0, datum['name'])
            sheet.write(index, 1, datum['department'])
            total = 0
            for times in datum['times']:
                time, idx = times
                if len(time) < 0:
                    continue
                time, ts = self.is_overtime(time)
                sheet.write(index, int(idx) + 2, " \n".join(time), self.style2)
                total += ts
            sheet.write(index, 2, total)

    def write_take_taxi(self):
        sheet = self.workbook.add_sheet('打车', cell_overwrite_ok=True)
        sheet.write(0, 0, '姓名')
        for i in range(self.length):
            sheet.write(0, i + 1, str(i + 1))

        for index, datum in enumerate(self.data):
            index += 1
            sheet.write(index, 0, datum['name'])
            for times in datum['times']:
                time, idx = times
                if len(time) < 0:
                    continue
                time, _ = self.is_take_a_taxi(time)
                sheet.write(index, int(idx) + 1, " \n".join(time), self.style2)

    @classmethod
    def is_overtime(cls, time: list) -> (list, int):
        return cls.calc_time(time, '21:00')

    @classmethod
    def is_take_a_taxi(cls, time: list) -> (list, int):
        return cls.calc_time(time, '21:30')

    @classmethod
    def calc_time(cls, time: list, bound_time: str) -> (list, int):
        valid_times = []
        times = 0
        t_21, t_00 = True, True
        for t in time:
            if t >= bound_time and t_21:
                valid_times.append(t)
                times += 1
                t_21 = False
            elif t_00 and '00:00' <= t <= '05:00':
                valid_times.append(t)
                times += 1
                t_00 = False
        return valid_times, times
