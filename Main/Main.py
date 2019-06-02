import wx, openpyxl, os, random, datetime


class MyFrame(wx.Frame):
    def __init__(self):
        super().__init__(parent=None, title='课堂随机点名软件', size=(700, 500), style=wx.DEFAULT_FRAME_STYLE | wx.STAY_ON_TOP)
        self.Center()
        self.panel = wx.Panel(self)
        self.lst = []
        self.number = 1000
        self.num = 1000
        self.path = ''
        vbox = wx.BoxSizer(wx.VERTICAL)
        hbox1 = wx.BoxSizer(wx.HORIZONTAL)
        hbox2 = wx.BoxSizer(wx.HORIZONTAL)
        hbox3 = wx.BoxSizer(wx.HORIZONTAL)

        Select = wx.Button(self.panel, label='选择文件', id=1)
        self.Bind(wx.EVT_BUTTON, self.GetExcel, id=1)

        Call = wx.Button(self.panel, label='顺序点名', id=2)
        self.Bind(wx.EVT_BUTTON, self.Call, id=2)

        Alive = wx.Button(self.panel, label='在席', id=4)
        Absent = wx.Button(self.panel, label='缺席', id=5)
        Stop = wx.Button(self.panel, label='结束', id=6)
        self.Bind(wx.EVT_BUTTON, self.Set, id=4, id2=6)
        self.Tip = wx.TextCtrl(self.panel, style=wx.TE_READONLY, size=(200, -1))
        self.StatusTip = wx.TextCtrl(self.panel, style=wx.TE_READONLY, size=(200, -1))

        RandomCall = wx.Button(self.panel, label='随机点名', id=3)
        self.Bind(wx.EVT_BUTTON, self.RandowCall, id=3)
        GetA = wx.Button(self.panel, label='A', id=7)
        GetB = wx.Button(self.panel, label='B', id=8)
        GetC = wx.Button(self.panel, label='C', id=9)
        GetD = wx.Button(self.panel, label='D', id=10)
        self.Bind(wx.EVT_BUTTON, self.GetGrage, id=7, id2=10)

        hbox1.Add(Select, flag=wx.ALL | wx.ALIGN_CENTER)
        hbox1.Add(self.Tip, flag=wx.ALL | wx.ALIGN_CENTER)
        hbox1.Add(self.StatusTip, flag=wx.ALL | wx.ALIGN_CENTER)

        hbox2.Add(Call, flag=wx.ALL | wx.ALIGN_CENTER)
        hbox2.Add(Alive, flag=wx.ALL | wx.ALIGN_CENTER)
        hbox2.Add(Absent, flag=wx.ALL | wx.ALIGN_CENTER)
        hbox2.Add(Stop, flag=wx.ALL | wx.ALIGN_CENTER)

        hbox3.Add(RandomCall, flag=wx.ALL | wx.ALIGN_CENTER)
        hbox3.Add(GetA, flag=wx.ALL | wx.ALIGN_CENTER)
        hbox3.Add(GetB, flag=wx.ALL | wx.ALIGN_CENTER)
        hbox3.Add(GetC, flag=wx.ALL | wx.ALIGN_CENTER)
        hbox3.Add(GetD, flag=wx.ALL | wx.ALIGN_CENTER)

        vbox.Add(hbox1, flag=wx.ALIGN_CENTER)
        vbox.Add(hbox2, flag=wx.ALIGN_CENTER)
        vbox.Add(hbox3, flag=wx.ALIGN_CENTER)
        self.panel.SetSizer(vbox)
        self.status_bar = self.CreateStatusBar()
        self.status_bar.SetStatusText('准备就绪')

    # 选择文件
    def GetExcel(self, event):
        file = '*.xlsx'
        dlg = wx.FileDialog(self, message='选择一个文件', defaultDir=os.getcwd(), style=wx.FD_OPEN, wildcard=file)
        if dlg.ShowModal() == wx.ID_OK:
            self.path = dlg.GetPath()  # 绝对路径
            filename = os.path.relpath(self.path)  # 文件名
            self.Tip.SetLabelText(filename)
            self.status_bar.SetStatusText('选择文件成功，文件：' + self.path)
            self.ReadExcel(self.path)
        dlg.Destroy()

    # 读取文件
    def ReadExcel(self, path):
        self.status_bar.SetStatusText('读取文件成功，文件：' + path)
        self.lst = []
        self.number = 1000
        self.num = 1000
        self.wb = openpyxl.load_workbook(path)
        self.sheet = self.wb.active
        self.lastRow = self.sheet.max_row + 1  # 最大行数
        self.lastCol = self.sheet.max_column + 1  # 最大列数
        self.firstRow = 2  # 初始行数
        self.firstCol = 2  # 初始列数
        for i in range(self.firstRow, self.lastRow):
            self.lst.append(self.sheet.cell(row=i, column=self.firstCol).value)

        self.date = datetime.datetime.now().strftime('%Y-%m-%d')
        print(self.lst)

    # 顺序点名
    def Call(self, event):
        if self.number >= 1000:
            self.lastRow = self.sheet.max_row + 1  # 重新获取最大行数
            self.lastCol = self.sheet.max_column + 1  # 重新获取最大列数
            self.number = 0
            if len(self.lst) != 0:
                self.StatusTip.SetLabelText('开始进行顺序点名')
                # self.Next()
                self.Tip.SetLabelText(self.lst[self.number])
                self.number += 1
                self.sheet.cell(row=self.firstRow - 1, column=self.lastCol).value = self.date
                self.wb.save(self.path)
            else:
                self.StatusTip.SetLabelText('顺序点名失败')
                self.Tip.SetLabelText('请先导入正确的文件')

    # 顺序点名出勤情况
    def Set(self, event):
        id = event.GetId()
        try:
            self.lst[self.number - 1]
        except:
            id = 0
        if id == 4:
            self.StatusTip.SetLabelText(self.lst[self.number - 1] + '状态录入成功，状态：在席')
            self.sheet.cell(row=self.firstRow + self.number - 1, column=self.lastCol).value = '在席'
            self.wb.save(self.path)
            self.Next()
        elif id == 5:
            self.StatusTip.SetLabelText(self.lst[self.number - 1] + '状态录入成功，状态：缺席')
            self.sheet.cell(row=self.firstRow + self.number - 1, column=self.lastCol).value = '缺席'
            self.wb.save(self.path)
            self.Next()
        elif id == 6:
            self.Tip.SetLabelText('顺序点名中止')
            self.number = 1000
        else:
            self.StatusTip.SetLabelText('状态录入失败，状态：未知')

    # 顺序点名下一位
    def Next(self):
        if self.number >= len(self.lst):
            self.Tip.SetLabelText('顺序点名完毕')
            self.number = 1000
        else:
            self.Tip.SetLabelText(self.lst[self.number])
            self.number += 1

    # 随机点名
    def RandowCall(self, event):
        if self.number >= 1000:
            if len(self.lst) != 0:
                if self.sheet.cell(row=self.firstRow - 1, column=self.lastCol).value != self.date + '得分':
                    # print(self.sheet.cell(row=self.firstRow - 1, column=self.lastCol).value)
                    if self.sheet.cell(row=self.firstRow - 1, column=self.lastCol).value != None:
                        self.lastRow = self.sheet.max_row + 1  # 重新获取最大行数
                        self.lastCol = self.sheet.max_column + 1  # 重新获取最大列数
                    self.sheet.cell(row=self.firstRow - 1, column=self.lastCol).value = self.date + '得分'
                    # self.wb.save(self.path)
                self.StatusTip.SetLabelText('开始随机点名')
                for i in range(30):
                    self.num = random.randrange(0, len(self.lst))
                    for j in range(self.firstCol, self.lastCol):
                        if self.sheet.cell(row=self.firstRow + self.num, column=j).value == 'A':
                            break
                    else:
                        break
                else:
                    self.StatusTip.SetLabelText('所有人获得最高分，之后将不再录入分数')
                    self.Tip.SetLabelText(self.lst[self.num])
                    self.num = 1001
                if self.num != 1001:
                    self.Tip.SetLabelText(self.lst[self.num])
            else:
                self.StatusTip.SetLabelText('随机点名失败')
                self.Tip.SetLabelText('请先导入正确的文件')
        else:
            if len(self.lst) != 0:
                self.StatusTip.SetLabelText('正在顺序点名中，请勿进行其他操作')
        # print(self.num)
        # print(self.lst[self.num])

    # 随机点名打分
    def GetGrage(self, event):
        id = event.GetId()
        try:
            self.lst[self.num]
        except:
            if self.num == 1001:
                id = 1
            else:
                id = 0
        if id == 7:
            self.StatusTip.SetLabelText(self.lst[self.num] + '成绩录入成功，成绩：A')
            self.sheet.cell(row=self.firstRow + self.num, column=self.lastCol).value = 'A'
            self.wb.save(self.path)
            self.num = 1001
        elif id == 8:
            self.StatusTip.SetLabelText(self.lst[self.num] + '成绩录入成功，成绩：B')
            self.sheet.cell(row=self.firstRow + self.num, column=self.lastCol).value = 'B'
            self.wb.save(self.path)
            self.num = 1001
        elif id == 9:
            self.StatusTip.SetLabelText(self.lst[self.num] + '成绩录入成功，成绩：C')
            self.sheet.cell(row=self.firstRow + self.num, column=self.lastCol).value = 'C'
            self.wb.save(self.path)
            self.num = 1001
        elif id == 10:
            self.StatusTip.SetLabelText(self.lst[self.num] + '成绩录入成功，成绩：D')
            self.sheet.cell(row=self.firstRow + self.num, column=self.lastCol).value = 'D'
            self.wb.save(self.path)
            self.num = 1001
        elif id == 1:
            self.StatusTip.SetLabelText('成绩录入失败，用户成绩已录入')
        else:
            self.StatusTip.SetLabelText('成绩录入失败，用户：未知')


class App(wx.App):
    def OnInit(self):
        frame = MyFrame()
        frame.Show()
        return True

    def OnExit(self):
        print('应用程序退出')
        return 0


if __name__ == '__main__':
    app = App()
    app.MainLoop()
