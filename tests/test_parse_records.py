import unittest
from parse.records import Records
from parse.writer import Writer


class TestParseRecords(unittest.TestCase):

    def test_parse(self):
        records = Records("./testdata/bbbbbb.xlsx")

        for data in records.parse_floor_18():
            print(data)
        write = Writer("./testdata/save.xls",
                       records.parse_floor_18(),
                       records.get_cols_len(),
                       records.get_sheet())
        write.save()

    def test_parse_floor_3(self):
        records = Records("./testdata/3楼考勤机.xls")
        for data in records.parse():
            print(data)
        write = Writer(
            "./testdata/save.xls",
            records.parse(),
            records.get_cols_len(),
            records.get_sheet())
        write.save()

    def test_split(self):
        import re
        skip_days = '1,2,3,4,5,123,12312,3141241'
        self.skip_days = re.compile(r'\d+').findall(skip_days)
        print(self.skip_days)

    def test_calc_time(self):
        yesterday = (['09:12', '23:49'], 2)
        time = ['00:12', '11:12', '19:00']
        idx = 3
        data = Writer.calc_time(time=time, bound_time='21:30', idx=idx, yesterday=yesterday)
        self.assertEqual(([], 0), data)

        yesterday = (['09:12', '23:49'], 2)
        time = ['03:12', '11:12', '19:00']
        idx = 3
        data = Writer.calc_time(time=time, bound_time='21:30', idx=idx, yesterday=yesterday)
        self.assertEqual((['03:12'], 1), data)

        yesterday = (['09:12', '23:49'], 2)
        time = ['11:12', '23:00']
        idx = 3
        data = Writer.calc_time(time=time, bound_time='21:30', idx=idx, yesterday=yesterday)
        self.assertEqual((['23:00'], 1), data)
