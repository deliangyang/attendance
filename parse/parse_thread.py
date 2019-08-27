import threading
import wx
import os
import re
from parse.records import Records
from parse.writer import Writer
from utils.logger import logger
from parse.errror import with_error_stack
from xlrd import XLRDError


class ParseThread(threading.Thread):

    def __init__(self, filename: str, dist: str, callback, skip_days: str):
        threading.Thread.__init__(self)
        self.filename = filename
        self.dist = dist
        self.callback = callback
        self.skip_days = self.parse_skip_days(skip_days)
        logger.info("start parse")

    def run(self) -> None:
        ok = False
        try:
            records = Records(self.filename)
            writer = Writer(
                self.dist,
                records.parse(),
                records.get_cols_len(),
                records.get_sheet(),
                skip_days=self.skip_days
            )
            writer.save()
            logger.info("end parse and wrote")
            os.system("start explorer %s" % os.path.dirname(self.dist))
            ok = True
        except XLRDError as _:
            dlg = wx.MessageDialog(None, '如果是18楼的考勤，请复制保存至另外一个文件', 'Excel表格不符合规范', wx.YES_NO)
            if dlg.ShowModal() == wx.ID_YES:
                pass
            dlg.Destroy()
        except UnicodeDecodeError as _:
            dlg = wx.MessageDialog(None, '如果是18楼的考勤，请复制保存至另外一个文件', 'Excel表格不符合规范', wx.YES_NO)
            if dlg.ShowModal() == wx.ID_YES:
                pass
            dlg.Destroy()
        except Exception as e:
            logger.error(with_error_stack(e))
        finally:
            if callable(self.callback):
                self.callback(ok)

    @classmethod
    def parse_skip_days(cls, skip_days) -> list:
        skip_days = re.compile(r'\d+').findall(skip_days)
        skip_days = list(
            filter(lambda x: 0 < x < 32,
                   map(lambda x: int(x), skip_days)))
        logger.info(skip_days)
        return skip_days
