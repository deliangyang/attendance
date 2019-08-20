import threading
import wx
import os
from parse.records import Records
from parse.writer import Writer
from utils.logger import logger
from parse.errror import with_error_stack


class ParseThread(threading.Thread):

    def __init__(self, filename: str, dist: str, callback):
        threading.Thread.__init__(self)
        self.filename = filename
        self.dist = dist
        self.callback = callback
        logger.info("start parse")

    def run(self) -> None:
        ok = False
        try:
            records = Records(self.filename)
            writer = Writer(
                self.dist,
                records.parse(),
                records.get_cols_len(),
                records.get_sheet()
            )
            writer.save()
            logger.info("end parse and wrote")
            os.system("start explorer %s" % os.path.dirname(self.dist))
            ok = True
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
