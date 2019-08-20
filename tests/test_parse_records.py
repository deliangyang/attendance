import unittest
import pandas
import xlrd
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
