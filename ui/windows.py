import wx
from ui.pannel import MainPanel


class MainWindows(wx.Frame):

    def __init__(self, parent, title):
        wx.Frame.__init__(self, parent, title=title, size=(800, 600))
        self.status_bar = self.CreateStatusBar()
        self.status_bar.SetFieldsCount(2)
        self.status_bar.SetStatusWidths([-1, -3])
        MainPanel(self)
        self.Show(True)

