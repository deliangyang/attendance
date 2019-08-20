import wx
from ui.windows import MainWindows


app = wx.App(False)
frame = MainWindows(None, "考勤统计")
app.MainLoop()

