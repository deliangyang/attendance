import wx
import os
import datetime
from utils.logger import logger
from parse.parse_thread import ParseThread


class MainPanel(wx.Panel):

    def __init__(self, parent):
        wx.Panel.__init__(self, parent, -1)
        self.btn = wx.Button(self, 1, "添加考勤记录", size=(100, 40), pos=(350, 200))
        self.sizer = wx.BoxSizer(wx.VERTICAL)
        self.status_bar = parent.status_bar
        self.sizer.Add(self.btn, 0, wx.CENTER, border=10)
        self.Bind(wx.EVT_BUTTON, self.on_get_file, self.btn)
        self.filename = ''
        self.dirname = ''

    def on_get_file(self, e):
        self.dirname = ''
        dlg = wx.FileDialog(self, "Choose a file", self.dirname, "", "*.xlsx;*.xls", wx.FD_OPEN)
        if dlg.ShowModal() == wx.ID_OK:
            self.filename = dlg.GetFilename()
            self.dirname = dlg.GetDirectory()
            logger.info({
                'filename': self.filename,
                'dirname': self.dirname,
            })
        dlg.Destroy()
        self.btn.Disable()
        self.status_bar.SetStatusText(u"状态：处理中...", 0)
        thread = ParseThread(
            self.dirname + os.sep + self.filename,
            self.get_storage_path(),
            callback=self.after_parse)
        thread.start()

    @classmethod
    def get_desktop_path(cls):
        return os.path.join(os.path.expanduser("~"), 'Desktop')

    def get_storage_path(self):
        dirname = os.path.join(
            self.get_desktop_path(),
            '考勤统计'
        )
        if not os.path.exists(dirname):
            os.makedirs(dirname)

        return os.path.join(
            dirname,
            (str(datetime.datetime.now()) + '.xls').replace(' ', '-').replace(':', '')
        )

    def after_parse(self):
        self.btn.Enable()
        self.status_bar.SetStatusText(u"状态：处理完毕", 0)
        self.status_bar.SetStatusText(u"文件路径：%s" % self.get_storage_path(), 1)
