import xlwt
import re
from xlrd.sheet import Sheet
from utils.logger import logger
from parse.errror import with_error_stack


class Writer(object):

    def __init__(self, filename: str, data: iter, length: int, origin_sheet: Sheet, skip_days: list):
        self.data = data
        self.filename = filename
        self.length = length
        self.workbook = xlwt.Workbook(encoding="utf-8")
        self.origin_sheet = origin_sheet
        self.split_time = re.compile(r'(\d{2}:\d{2})')
        self.skip_days = skip_days

        alignment = xlwt.Alignment()
        alignment.wrap = 1
        alignment.horz = 0x02
        self.style2 = xlwt.XFStyle()
        self.style2.alignment = alignment

        self.re_int = re.compile(r'^\d+\.0$')

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
                if self.re_int.match(str(rd)):
                    rd = rd.replace('.0', '')
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
            yesterday = ([], -2)
            for times in datum['times']:
                time, idx = times
                if len(time) <= 0:
                    continue
                # 今天是排除的日子，昨天不是
                if idx in self.skip_days and idx - 1 not in self.skip_days:
                    ye_t, _ = yesterday
                    if not self.has_overtime(ye_t) and self.is_early_morning_over_time(time):
                        # 昨天没有显示加班, 今天凌晨有加班
                        sheet.write(index, int(idx) + 2, " \n".join([time[0]]), self.style2)
                        total += 1
                elif idx in self.skip_days:
                    pass
                else:
                    time, ts = self.is_overtime(time, idx, yesterday)
                    sheet.write(index, int(idx) + 2, " \n".join(time), self.style2)
                    total += ts
                yesterday = times
            sheet.write(index, 2, total)

    def write_take_taxi(self):
        sheet = self.workbook.add_sheet('打车', cell_overwrite_ok=True)
        sheet.write(0, 0, '姓名')
        for i in range(self.length):
            sheet.write(0, i + 1, str(i + 1))

        for index, datum in enumerate(self.data):
            index += 1
            sheet.write(index, 0, datum['name'])
            yesterday = ([], -2)
            for times in datum['times']:
                time, idx = times
                if len(time) < 0:
                    continue
                time, _ = self.is_take_a_taxi(time, idx, yesterday)
                yesterday = times
                sheet.write(index, int(idx), " \n".join(time), self.style2)

    @classmethod
    def is_overtime(cls, time: list, idx: int, yesterday) -> (list, int):
        return cls.calc_time(time, '21:00', idx, yesterday)

    @classmethod
    def is_take_a_taxi(cls, time: list, idx: int, yesterday) -> (list, int):
        return cls.calc_time(time, '21:30', idx, yesterday)

    @classmethod
    def calc_time(cls, time: list, bound_time: str, idx: int, yesterday) -> (list, int):
        valid_times = []
        times = 0
        t_21, t_00 = True, True
        yesterday_time, yesterday_idx = yesterday
        for t in time:
            if int(idx) - int(yesterday_idx) == 1 and len(yesterday_time) > 0:
                try:
                    if int(str(t).replace(':', '')) + 2400 - int(str(yesterday_time[-1]).replace(':', '')) <= 100:
                        continue
                except Exception as e:
                    logger.error({
                        'e': e,
                        'yesterday': yesterday,
                        'time': time,
                    })
            if t >= bound_time and t_21:
                valid_times.append(t)
                times += 1
                t_21 = False
            elif t_00 and '00:00' <= t <= '05:00':
                valid_times.append(t)
                times += 1
                t_00 = False
        return valid_times, times

    @classmethod
    def has_overtime(cls, date_times: list) -> bool:
        for date_time in date_times:
            if '21:00' <= date_time <= '23:59':
                return True
        return False

    @classmethod
    def is_early_morning_over_time(cls, date_times: list):
        for date_time in date_times:
            if '00:00' <= date_time <= '05:00':
                return True
        return False
