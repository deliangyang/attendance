import wx
import os
import datetime
from parse.parse_thread import ParseThread


class MainPanel(wx.Panel):

    def __init__(self, parent):
        wx.Panel.__init__(self, parent, -1)
        self.status_bar = parent.status_bar
        self.filename = ''
        self.dirname = ''

        self.st_tips = wx.StaticText(self, 0, "不计算加班的日期(任意字符分隔):", style=wx.TE_LEFT)
        top_box_sizer = wx.BoxSizer(wx.VERTICAL)
        top_box_sizer.Add(self.st_tips, proportion=0, flag=wx.TOP, border=50)

        b_sizer_all = wx.BoxSizer(wx.VERTICAL)

        center_box_sizer = wx.BoxSizer(wx.HORIZONTAL)
        btn_box_sizer = wx.BoxSizer(wx.HORIZONTAL)
        self.btn = wx.Button(self, 1, "添加考勤记录")
        self.not_overtime = wx.TextCtrl(self, 1, style=wx.TE_LEFT | wx.TE_MULTILINE, size=(1000, 160))

        center_box_sizer.Add(self.not_overtime, proportion=1, flag=wx.ALL | wx.ALIGN_CENTER_VERTICAL,
                             border=5)
        btn_box_sizer.Add(self.btn, proportion=0, flag=wx.ALL|wx.CENTER, border=5)

        b_sizer_all.Add(top_box_sizer, proportion=0, flag=wx.EXPAND | wx.ALL, border=5)
        b_sizer_all.Add(center_box_sizer, proportion=0, flag=wx.EXPAND | wx.ALL, border=5)
        b_sizer_all.Add(btn_box_sizer, proportion=0, flag=wx.CENTER, border=20)
        self.SetSizer(b_sizer_all)
        self.Bind(wx.EVT_BUTTON, self.on_get_file, self.btn)

    def on_get_file(self, e):
        dlg = wx.FileDialog(self, "Choose a file", self.dirname, "", "*.xlsx;*.xls", wx.FD_OPEN)
        if dlg.ShowModal() == wx.ID_CANCEL:
            dlg.Destroy()
            return
        self.filename = dlg.GetFilename()
        self.dirname = dlg.GetDirectory()
        skip_days = self.not_overtime.GetValue()
        dlg.Destroy()
        self.btn.Disable()
        self.init_status_bar()
        self.status_bar.SetStatusText(u"状态：处理中...", 0)
        thread = ParseThread(
            self.dirname + os.sep + self.filename,
            self.get_storage_path(),
            callback=self.after_parse,
            skip_days=skip_days)
        thread.start()

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

    def after_parse(self, ok):
        self.btn.Enable()
        if ok:
            self.status_bar.SetStatusText(u"状态：处理完毕", 0)
            self.status_bar.SetStatusText(u"文件路径：%s" % self.get_storage_path(), 1)
        else:
            self.status_bar.SetStatusText(u"状态：处理失败", 0)
            self.status_bar.SetStatusText(u"文件路径：", 1)

    @classmethod
    def get_desktop_path(cls):
        return os.path.join(os.path.expanduser("~"), 'Desktop')

    def init_status_bar(self):
        self.status_bar.SetStatusText(u"状态：添加考勤记录", 0)
        self.status_bar.SetStatusText(u"文件路径：", 1)
