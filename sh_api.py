import wx
import wx.grid
from wx.lib import plot as wxplot
#import time
from datetime import datetime
import requests
from urllib.request import urlopen
import pandas as pd
import json
import threading
import os.path
import shutil

import shioaji as sj
from shioaji import TickFOPv1, Exchange

"""
    Frame -> Panel -> Control (StaticText, TextCtrl ...)
               |   -> Sizer (BoxSizer , GridSizer)
                               |            |--> FlexGridSizer -> GridBabSizer
                               |-->StaticBox, wrapSizer,
"""

name_font = 9
font_size = 10

class MainFrame(wx.Frame):
    def __init__(self, parent, title):
        super(MainFrame, self).__init__(parent, title = title, size = (1130,730))
        self.time_center = datetime.now().strftime("%Y/%m/%d %H:%M:%S")
        self.cat_rowLabels = ["類股", "量增", "漲跌"]
        self.stock_colLabels = ["代號", "名稱", "增量", "漲跌%", "成交價", "量比", "估量", "昨量"]
        self.chart_xData = [i for i in range(1145)]
        self.chart_yyData = [0 for i in range(1145)]
        self.chart_xyy = list(zip(self.chart_xData, self.chart_yyData))
        self.chart2_xData = [i for i in range(3600)]
        self.chart2_yBase = [0 for i in range(3600)]
        self.chart2_xyBase = list(zip(self.chart2_xData, self.chart2_yBase))
        self.baseline = wxplot.PolySpline(
                self.chart2_xyBase,
                colour=wx.Colour(255, 0, 0),   # Color: red
                width=1,
            )

        # 永豐 Api
        
        self.option_category_w = ''
        self.option_code_w = list()
        self.option_data_w = list()

        self.option_category_m = ''
        self.option_code_m = list()
        self.option_data_m = list()

        self.subscribed_item = set()
        self.api = None
        
        self.InitUI()

        self.timer = wx.Timer(self)
        self.Bind(wx.EVT_TIMER, self.onTimeToggle, self.timer)
        self.timer.Start(1000)

        #self.onUpdateOP()

        self.onUpdateVix()
        self.onUpdateCat()
        self.onUpdateStockFromRtd()
        
    def InitUI(self):
        # Font setting
        time_font = wx.Font(12, wx.SWISS, wx.NORMAL, wx.BOLD)
        other_font = wx.Font(10, wx.SWISS, wx.NORMAL, wx.BOLD)
        #setting StatusBar
        self.CreateStatusBar()
        self.SetStatusText(f"開始上班賺錢...")

        # menu start
        menubar = wx.MenuBar()
        sysMenu = wx.Menu()
        # wx.MenuItem(parent_menu, id, text, kind)
        app_quit = wx.MenuItem(sysMenu, wx.ID_CLOSE, '離開')
        sysMenu.Append(app_quit)

        # 永豐選單
        sjMenu = wx.Menu()
        sj_login = wx.MenuItem(sjMenu, 201, '登入')
        sj_subscribe_option = wx.MenuItem(sjMenu, 203, '訂閱選擇權')
        sj_unsubscribe_option = wx.MenuItem(sjMenu, 203, '取消訂閱選擇權')
        sj_logout = wx.MenuItem(sjMenu, 204, '登出')
        sj_subscribe_Future = wx.MenuItem(sjMenu, 205, '訂閱指數與期貨')
        sj_unsubscribe_Future = wx.MenuItem(sjMenu, 206, '取消訂閱指數與期貨')
        sjMenu.Append(sj_login)
        sjMenu.Append(sj_subscribe_Future)
        sjMenu.Append(sj_subscribe_option)
        sjMenu.Append(sj_unsubscribe_option)
        sjMenu.Append(sj_unsubscribe_Future)
        sjMenu.Append(sj_logout)

        menubar.Append(sysMenu, '系統')
        menubar.Append(sjMenu, '永豐API')
        self.SetMenuBar(menubar)

        self.Bind(wx.EVT_MENU, self.onMenuHandler)
        # menu end

        # splitter window
        mainSplitter = wx.SplitterWindow(self, -1)
        splitter = wx.SplitterWindow(mainSplitter, -1)
        notebook = wx.Notebook(mainSplitter)
        stockSplitter = wx.SplitterWindow(notebook, -1)
        chartSplitter = wx.SplitterWindow(notebook, -1)
        
        # stock screen
        panel_buy = wx.Panel(stockSplitter, -1, size=(550,350))
        self.sbox_stock_buy = wx.StaticBox(panel_buy, -1, '做多觀察')
        sbSizer_stock_buy = wx.StaticBoxSizer(self.sbox_stock_buy, wx.VERTICAL)

        # table for cat
        self.table_cat = wx.grid.Grid(panel_buy)
        self.table_cat.EnableEditing(False)      # 停用輸入功能
        self.table_cat.SetRowLabelSize(40)        # hide row labels
        self.table_cat.SetColLabelSize(0)        # hide row labels
        self.table_cat.CreateGrid(3,10)

        attr_cat = wx.grid.GridCellAttr()
        attr_cat.SetAlignment(wx.ALIGN_CENTER,wx.ALIGN_CENTER)
        attr_cat.SetTextColour("black")
        attr_cat.SetFont(wx.Font(font_size, wx.SWISS, wx.NORMAL, wx.BOLD))

        attr_cat_name = wx.grid.GridCellAttr()
        attr_cat_name.SetBackgroundColour(wx.Colour(255, 191, 0))
        attr_cat_name.SetFont(wx.Font(8, wx.SWISS, wx.NORMAL, wx.BOLD))

        self.table_cat.SetRowAttr(0, attr_cat_name)
        self.table_cat.SetCellValue(0, 9, "大盤量")
        self.table_cat.SetCellBackgroundColour(1, 9, "pink")
        self.table_cat.SetCellBackgroundColour(2, 9, "yellow")

        for i in range(3):
            self.table_cat.SetRowLabelValue(i, self.cat_rowLabels[i])

        for j in range(10):
            self.table_cat.SetColSize(j, 55)
            self.table_cat.SetColAttr(j, attr_cat.Clone())
        
        # stock table setting
        attr_stockCode = wx.grid.GridCellAttr()
        attr_stockCode.SetAlignment(wx.ALIGN_CENTER,wx.ALIGN_CENTER)
        #attr_stockCode.SetTextColour(wx.Colour(128, 159, 255))
        attr_stockCode.SetTextColour('yellow')
        attr_stockCode.SetBackgroundColour("black")
        attr_stockCode.SetFont(wx.Font(font_size, wx.SWISS, wx.NORMAL, wx.BOLD))

        attr_stockName = wx.grid.GridCellAttr()
        attr_stockName.SetAlignment(wx.ALIGN_LEFT,wx.ALIGN_CENTER)
        #attr_stockName.SetTextColour(wx.Colour(128, 159, 255))
        attr_stockName.SetTextColour('yellow')
        attr_stockName.SetBackgroundColour("black")
        attr_stockName.SetFont(wx.Font(name_font, wx.SWISS, wx.NORMAL, wx.BOLD))

        attr_stockData = wx.grid.GridCellAttr()
        attr_stockData.SetAlignment(wx.ALIGN_CENTER,wx.ALIGN_CENTER)
        attr_stockData.SetTextColour("white")
        attr_stockData.SetBackgroundColour("black")
        attr_stockData.SetFont(wx.Font(font_size, wx.SWISS, wx.NORMAL, wx.NORMAL))

        # table for buy
        self.table_buy = wx.grid.Grid(panel_buy)
        self.table_buy.EnableEditing(False)      # 停用輸入功能
        self.table_buy.SetRowLabelSize(0)        # hide row labels
        self.table_buy.CreateGrid(50,8)
        self.table_buy.SetGridLineColour("gray")

        for i in range(8):
            self.table_buy.SetColLabelValue(i, self.stock_colLabels[i])
            self.table_buy.SetColSize(i, 70)
            if i == 0:
                self.table_buy.SetColAttr(i, attr_stockCode.Clone())
            elif i == 1:
                self.table_buy.SetColAttr(i, attr_stockName.Clone())
            else:
                self.table_buy.SetColAttr(i, attr_stockData.Clone())

        self.table_buy.SetColSize(1, 80)

        sbSizer_stock_buy.Add(self.table_cat)
        sbSizer_stock_buy.Add(self.table_buy)
        panel_buy.SetSizer(sbSizer_stock_buy)

        # stock down
        panel_sell = wx.Panel(stockSplitter, -1)
        self.sbox_stock_sell = wx.StaticBox(panel_sell, -1, '做空觀察')
        sbSizer_stock_sell = wx.StaticBoxSizer(self.sbox_stock_sell, wx.HORIZONTAL)

        self.table_sell = wx.grid.Grid(panel_sell)
        self.table_sell.EnableEditing(False)      # 停用輸入功能
        self.table_sell.SetRowLabelSize(0)        # hide row labels
        self.table_sell.CreateGrid(50,8)
        self.table_sell.SetGridLineColour("gray")

        for i in range(8):
            self.table_sell.SetColLabelValue(i, self.stock_colLabels[i])
            self.table_sell.SetColSize(i, 70)
            if i == 0:
                self.table_sell.SetColAttr(i, attr_stockCode.Clone())
            elif i == 1:
                self.table_sell.SetColAttr(i, attr_stockName.Clone())
            else:
                self.table_sell.SetColAttr(i, attr_stockData.Clone())
        
        self.table_sell.SetColSize(1, 80)
        sbSizer_stock_sell.Add(self.table_sell)
        panel_sell.SetSizer(sbSizer_stock_sell)

        #設定 OP Table
        pnl_op = wx.Panel(splitter, -1, size=(550,500))
        self.sbox_op = wx.StaticBox(pnl_op, -1, '日盤')
        sbSizer_day = wx.StaticBoxSizer(self.sbox_op, wx.VERTICAL)
        #sbSizer_day = wx.BoxSizer(wx.VERTICAL)
        
        # Start Btn
        self.toggleBtn = wx.Button(pnl_op, wx.ID_ANY, "Start")
        self.toggleBtn.Bind(wx.EVT_BUTTON, self.onToggle)

        # 時間
        self.ntime = wx.StaticText(pnl_op, -1, "2022/08/31\n\n08:46:00", style = wx.ALIGN_CENTRE_HORIZONTAL, size=(80,40))
        self.ntime.SetBackgroundColour((255,255,0))
        self.ntime.SetFont(time_font)

        vbSizer_1 = wx.BoxSizer(wx.VERTICAL)
        vbSizer_1.Add(self.toggleBtn, 0, wx.ALL|wx.ALIGN_LEFT,2)
        vbSizer_1.Add(self.ntime, 0, wx.ALL|wx.ALIGN_LEFT,2)

        # 現貨
        self.twse = wx.StaticText(pnl_op, -1, "現      貨")
        self.twse_data = wx.StaticText(pnl_op, -1, "14949", style = wx.ALIGN_CENTRE_HORIZONTAL, size=(70,45))
        self.twse_data.SetBackgroundColour((255,255,0))

        vbSizer_2 = wx.BoxSizer(wx.VERTICAL)
        vbSizer_2.Add(self.twse, 0, wx.ALL|wx.ALIGN_LEFT,2)
        vbSizer_2.Add(self.twse_data, 0, wx.ALL|wx.ALIGN_LEFT,2)
        # 期貨
        self.fut = wx.StaticText(pnl_op, -1, "期      貨")
        self.fut_data = wx.StaticText(pnl_op, -1, "14789", style = wx.ALIGN_CENTRE_HORIZONTAL, size=(70,45))
        self.fut_data.SetBackgroundColour((255,255,0))
        
        vbSizer_3 = wx.BoxSizer(wx.VERTICAL)
        vbSizer_3.Add(self.fut, 0, wx.ALL|wx.ALIGN_LEFT,2)
        vbSizer_3.Add(self.fut_data, 0, wx.ALL|wx.ALIGN_LEFT,2)
        # 價差
        self.diff = wx.StaticText(pnl_op, -1, "價     差")
        self.diff_data = wx.StaticText(pnl_op, -1, "-140", style = wx.ALIGN_CENTRE_HORIZONTAL, size=(70,45))
        self.diff_data.SetBackgroundColour((255,255,0))

        vbSizer_4 = wx.BoxSizer(wx.VERTICAL)
        vbSizer_4.Add(self.diff, 0, wx.ALL|wx.ALIGN_LEFT,2)
        vbSizer_4.Add(self.diff_data, 0, wx.ALL|wx.ALIGN_LEFT,2)

        self.trend = wx.StaticText(pnl_op, -1, "盤勢力道")
        self.trend_data = wx.StaticText(pnl_op, -1, "--測", style = wx.ALIGN_CENTRE_HORIZONTAL, size=(70,45))
        self.trend_data.SetBackgroundColour((255,255,0))

        vbSizer_5 = wx.BoxSizer(wx.VERTICAL)
        vbSizer_5.Add(self.trend, 0, wx.ALL|wx.ALIGN_LEFT,2)
        vbSizer_5.Add(self.trend_data, 0, wx.ALL|wx.ALIGN_LEFT,2)

        # 改為未平倉
        oi_force = 0 / 1
        oi_text = format(oi_force, '.2f')
        #self.oi_data.SetLabel(oi_text)
        self.oi = wx.StaticText(pnl_op, -1, "未平倉")
        self.oi_data = wx.StaticText(pnl_op, -1, oi_text, style = wx.ALIGN_CENTRE_HORIZONTAL, size=(70,45))
        self.oi_data.SetBackgroundColour((255,255,0))

        vbSizer_6 = wx.BoxSizer(wx.VERTICAL)
        vbSizer_6.Add(self.oi, 0, wx.ALL|wx.ALIGN_LEFT,2)
        vbSizer_6.Add(self.oi_data, 0, wx.ALL|wx.ALIGN_LEFT,2)
        
        
        self.twse_data.SetFont(other_font)
        self.fut_data.SetFont(other_font)
        self.diff_data.SetFont(other_font)
        self.trend_data.SetFont(other_font)
        self.oi_data.SetFont(other_font)
        

        hBox_op = wx.BoxSizer(wx.HORIZONTAL)
        hBox_op.Add(vbSizer_1, 0, wx.ALL|wx.CENTRE,2)
        hBox_op.Add(vbSizer_2, 0, wx.ALL|wx.CENTRE,2)
        hBox_op.Add(vbSizer_3, 0, wx.ALL|wx.CENTRE,2)
        hBox_op.Add(vbSizer_4, 0, wx.ALL|wx.CENTRE,2)
        hBox_op.Add(vbSizer_5, 0, wx.ALL|wx.CENTRE,2)
        hBox_op.Add(vbSizer_6, 0, wx.ALL|wx.CENTRE,2)

        sbSizer_day.Add(hBox_op, 0, wx.ALL|wx.EXPAND, 2)

        #設定 第一個區塊: 登入、內部元件排列 水平方向
        loginSubName1 = wx.StaticBox(pnl_op, -1, '登入')
        loginSubSizer1 = wx.StaticBoxSizer(loginSubName1, wx.HORIZONTAL)
        loginSubBox1 = wx.BoxSizer(wx.HORIZONTAL)
        accountLabel = wx.StaticText(pnl_op, -1, "帳號")
        # wx.TextCtrl(parent, id, value, pos, size, style)
        self.accountText = wx.TextCtrl(pnl_op, -1, style = wx.ALIGN_LEFT)
        
        passwdLabel = wx.StaticText(pnl_op, -1, "API Key")
        self.passwdText = wx.TextCtrl(pnl_op, -1, style = wx.ALIGN_LEFT | wx.TE_PASSWORD)

        #loginButton = wx.Button(pnl_op, -1, '登入')
        loginSimButton = wx.Button(pnl_op, -1, '模擬登入')
        loginSimButton.Bind(wx.EVT_BUTTON, self.onLoginSj)
        #loginSubBox2.Add(loginButton, 0, wx.ALL|wx.LEFT, 10)
            
        loginSubBox1.Add(accountLabel, 0, wx.ALL|wx.CENTER, 5)
        loginSubBox1.Add(self.accountText, 0, wx.ALL|wx.CENTER, 5)
        loginSubBox1.Add(passwdLabel, 0, wx.ALL|wx.CENTER, 5)
        loginSubBox1.Add(self.passwdText, 0, wx.ALL|wx.CENTER, 5)
        #loginSubBox1.Add(loginButton, 0, wx.ALL|wx.CENTER, 5)
        loginSubBox1.Add(loginSimButton, 0, wx.ALL|wx.CENTER, 5)
        loginSubSizer1.Add(loginSubBox1, 0, wx.ALL|wx.CENTER, 10)

        sbSizer_day.Add(loginSubSizer1, 0, wx.ALL|wx.EXPAND, 5)

        loginSubName3 = wx.StaticBox(pnl_op, -1, "狀態")
        loginSubSizer3 = wx.StaticBoxSizer(loginSubName3, wx.VERTICAL)

        loginSubBox3 = wx.BoxSizer(wx.VERTICAL)
        self.loginSubResponse = wx.TextCtrl(pnl_op, -1, size=(500,100), style = wx.ALIGN_LEFT|wx.TE_MULTILINE|wx.TE_READONLY)
        
        loginSubBox3.Add(self.loginSubResponse, 0, wx.ALL | wx.CENTER, 10)
        loginSubSizer3.Add(loginSubBox3, 0, wx.ALL|wx.CENTER, 10)

        sbSizer_day.Add(loginSubSizer3, 0, wx.ALL|wx.EXPAND, 5)
        
        pnl_op.SetSizer(sbSizer_day)

        # chart init
        pnl_chart = wx.Panel(splitter, -1)
        chartBox = wx.BoxSizer(wx.HORIZONTAL)

        x_data = [1, 2, 3, 4, 5, 6, 7, 8, 9]
        y_data = [2, 4, 6, 4, 2, 5, 6, 7, 1]
        xy_data = list(zip(x_data, y_data))
        
        line = wxplot.PolySpline(
            xy_data,
            colour=wx.Colour(0, 0, 255),
            width=3,
        )

        # create chart object
        chart = wxplot.PlotGraphics([line])
        # create canvas
        self.canvas = wxplot.PlotCanvas(pnl_chart)
        self.canvas.Draw(chart)

        chartBox.Add(self.canvas, 1, wx.EXPAND)
        pnl_chart.SetSizerAndFit(chartBox)

        # chart: 委買賣口差
        panel_lotChart = wx.Panel(chartSplitter, -1)
        chartBox2 = wx.BoxSizer(wx.HORIZONTAL)
        lot_chart = wxplot.PlotGraphics([line])
        self.lotCanvas = wxplot.PlotCanvas(panel_lotChart)
        self.lotCanvas.Draw(lot_chart)
        chartBox2.Add(self.lotCanvas, 1, wx.EXPAND)
        panel_lotChart.SetSizerAndFit(chartBox2)

        # chart: 委買賣筆差
        panel_orderChart = wx.Panel(chartSplitter, -1)
        chartBox3 = wx.BoxSizer(wx.HORIZONTAL)
        
        order_chart = wxplot.PlotGraphics([line])
        self.orderCanvas = wxplot.PlotCanvas(panel_orderChart)
        self.orderCanvas.Draw(lot_chart)
        chartBox3.Add(self.orderCanvas, 1, wx.EXPAND)
        panel_orderChart.SetSizerAndFit(chartBox3)

        #設定 Box 內的Control 元件位置方向
        #mainBox = wx.BoxSizer(wx.HORIZONTAL)
        #mainBox.Add(sbSizer_day,0, wx.ALL|wx.ALIGN_CENTER, 2)
        #mainBox.Add(self.canvas,1, wx.ALL|wx.EXPAND, 2)
        
        splitter.SplitHorizontally(pnl_op, pnl_chart)
        stockSplitter.SplitHorizontally(panel_buy, panel_sell)
        chartSplitter.SplitHorizontally(panel_lotChart, panel_orderChart, sashPosition=325)

        notebook.AddPage(stockSplitter, "Stock")
        notebook.AddPage(chartSplitter, "Future Chart")

        mainSplitter.SplitVertically(splitter, notebook, sashPosition=500)
        #self.SetSizer(mainBox)
        self.Centre()
        #self.Fit()

        # Bind event 兩種寫法
        #loginButton.Bind(wx.EVT_BUTTON, self.onLoginBtn)
        #self.Bind(wx.EVT_BUTTON, self.onCancelBtn, cancelButton)
        #self.SetProperties()

    def SetProperties(self):
        frameIcon = wx.Icon("./icon/FCat2.ico")
        self.SetIcon(frameIcon)

    def onMenuHandler(self, event):
        evt_id = event.GetId()
        if evt_id == wx.ID_CLOSE:
            self.Close(True)
        elif evt_id == 101:
            self.GetStocks()
        elif evt_id == 102:
            self.onUpdateFutureOrder()
        elif evt_id == 201:
            self.onLoginSj()
        elif evt_id == 202:
            self.onSubscribeOption()
        elif evt_id == 203:
            self.onUnsubscribeOption()
        elif evt_id == 204:
            self.onLogoutSj()
        elif evt_id == 205:
            self.onSubscribeFuture()
        elif evt_id == 206:
            pass

    def onCloseWindow(self, event):
        self.timer.Stop()
        self.Destroy()

    def onToggle(self, event):
        btnLabel = self.toggleBtn.GetLabel()
        if btnLabel == "Start":
            #self.timer.Start(1000)
            self.toggleBtn.SetLabel("Stop")
        else:
            #self.timer.Stop()
            self.toggleBtn.SetLabel("Start")

    def onTimeToggle(self, event):
        now = datetime.now()
        
        self.time_center = now.strftime("%Y/%m/%d %H:%M:%S")
        today = self.time_center[:10]
        t_now = self.time_center[11:]
        
        time_now = today + '\n' + t_now
        self.ntime.SetLabel(time_now)
        
    
    def onUpdateStock(self):
        # 過濾出有用到的欄位
        col_code = ['c','n','z','tv','v','o','h','l','y']
        col_name = ['股票代號','公司簡稱','當盤成交價','當盤成交量','累積成交量','開盤價','最高價','最低價','昨收價']
        # 組成stock_list 最多 150 一次
        tse_list = '|'.join('tse_{}.tw'.format(tse_stk) for tse_stk in self.tse_stocks[:150])
        otc_list = '|'.join('otc_{}.tw'.format(otc_stk) for otc_stk in self.otc_stocks)
        #print(tse_list)
        #　query data
        query_tse = "https://mis.twse.com.tw/stock/api/getStockInfo.jsp?ex_ch="+ tse_list + "&json=1&delay=0"
        data_tse = json.loads(urlopen(query_tse).read())
        #print(data_tse)
        df_tse = pd.DataFrame(data_tse['msgArray'], columns=col_code)
        df_tse.columns = col_name
        


    def onUpdateVix(self):
        url = 'https://mis.taifex.com.tw/futures/api/getChartDataTick'
        payload ={
            "SymbolID": "TAIWANVIX"
        }
        res = requests.post(url, json=payload)
        data = res.json()
        vix = data['RtData']['Ticks']   #list
        
        #x_data = [row[0] for row in vix]
        #x_data = [i for i in range(1145)]
        x_data = self.chart_xData
        y_data = [row[4] for row in vix]
        
        xy_data = list(zip(x_data, y_data))
        
        line = wxplot.PolySpline(
            xy_data,
            colour=wx.Colour(0, 0, 255),   # Color: blue
            width=2,
        )
        """
        line2 = wxplot.PolySpline(
            self.chart_xyy,
            colour=wx.Colour(255, 0, 0),   # Color: black
            width=1,
            #style=wx.PENSTYLE_DOT
        )
        """
        chart = wxplot.PlotGraphics([line])
        #chart = wxplot.PlotGraphics([line, line2])
        # draw chart
        self.canvas.Draw(chart, xAxis=(0,1145))

        # save data after close
        t_now = self.time_center[11:]
        if t_now >= "14:00:00":
            today = self.time_center[:10].replace('/','')
            vix_file = './vix/' + today + '_vix.txt'
            if not os.path.exists(vix_file):
                t_data = [row[0] for row in vix]
                with open(vix_file, 'w') as f:
                    for i in range(len(t_data)):
                        f.write(f"{t_data[i]},{y_data[i]}\n")

    
        
    def onUpdateCat(self):
        today = self.time_center[:10].replace('/','')
        tmpFile = 'Cat.txt'
        catfile = './stock/' + today + '_Cat.txt'
        shutil.copyfile(tmpFile, catfile)

        if os.path.exists(catfile):
            with open(catfile, 'r', encoding='Big5') as f:
                lines = f.readlines()
                
            l0 = lines[0].split(',')
            predict_vol = round(float(l0[1]))
            ytd_vol = round(float(l0[2]))
            self.table_cat.SetCellValue(1, 9, "%s" % predict_vol)
            self.table_cat.SetCellValue(2, 9, "%s" % ytd_vol)
            if predict_vol >= ytd_vol:
                self.table_cat.SetCellTextColour(1, 9, "red")
            else:
                self.table_cat.SetCellTextColour(1, 9, wx.Colour(0,204,0))
            for i in range(1,4):
                cat_data = lines[i].split(',')
                for j in range(9):
                    if i == 1:
                        data = cat_data[j].replace('"','')
                    elif i == 2:
                        d = float(cat_data[j])
                        data = format(float(cat_data[j]), '.2f')
                        if d >= 1:
                            self.table_cat.SetCellBackgroundColour(i-1, j, "yellow")
                        else:
                            self.table_cat.SetCellBackgroundColour(i-1, j, "white")
                    elif i == 3:
                        d = float(cat_data[j])
                        data = format(float(cat_data[j]), '.2f')
                        if d >= 0.5:
                            self.table_cat.SetCellBackgroundColour(i-1, j, "red")
                        elif d <= -0.5:
                            self.table_cat.SetCellBackgroundColour(i-1, j, wx.Colour(0,204,0))
                        else:
                            self.table_cat.SetCellBackgroundColour(i-1, j, "white")
                    else:
                        data = cat_data[j]
                    self.table_cat.SetCellValue(i-1, j, "%s" % data)


    def onUpdateStockFromRtd(self):
        today = self.time_center[:10].replace('/','')
        tmpFile1 = 'buy_list.txt'
        tmpFile2 = 'sell_list.txt'
        
        upFile = './stock/' + today + '_buy_list.txt'
        dnFile = './stock/' + today + '_sell_list.txt'
        shutil.copyfile(tmpFile1, upFile)
        shutil.copyfile(tmpFile2, dnFile)

        if os.path.exists(upFile):
            df_buy = pd.read_csv(upFile, sep=",", header=None, names=self.stock_colLabels, encoding='Big5')
            df_buy = df_buy.sort_values(by='增量', ascending=False)
            self.table_buy.ClearGrid()
            for r in range(len(df_buy)):
                for c in range(len(df_buy.columns)):
                    if c == 5:
                        data = df_buy.iloc[r,c].astype('float').round(2)
                    elif c == 3:
                        data = df_buy.iloc[r,c].astype('float').round(2)
                        if data >= 3 and data <5:
                            self.table_buy.SetCellBackgroundColour(r, c, wx.Colour(255, 153, 128))
                            self.table_buy.SetCellFont(r, c, wx.Font(font_size, wx.SWISS, wx.NORMAL, wx.BOLD))
                        elif data >= 5:
                            self.table_buy.SetCellBackgroundColour(r, c, wx.Colour(255, 51, 0))
                            self.table_buy.SetCellFont(r, c, wx.Font(font_size, wx.SWISS, wx.NORMAL, wx.BOLD))
                        else:
                            self.table_buy.SetCellBackgroundColour(r, c, 'black')
                            self.table_buy.SetCellFont(r, c, wx.Font(font_size, wx.SWISS, wx.NORMAL, wx.NORMAL))
                    elif c == 2:
                        data = df_buy.iloc[r,c].astype('int')
                        if data >=300 and data < 600:
                            self.table_buy.SetCellBackgroundColour(r, c, 'pink')
                            self.table_buy.SetCellFont(r, c, wx.Font(font_size, wx.SWISS, wx.NORMAL, wx.BOLD))
                        elif data >= 600:
                            self.table_buy.SetCellBackgroundColour(r, c, 'red')
                            self.table_buy.SetCellFont(r, c, wx.Font(font_size, wx.SWISS, wx.NORMAL, wx.BOLD))
                        else:
                            self.table_buy.SetCellBackgroundColour(r, c, 'black')
                            self.table_buy.SetCellFont(r, c, wx.Font(font_size, wx.SWISS, wx.NORMAL, wx.NORMAL))
                    else:
                        data = df_buy.iloc[r,c]
                    self.table_buy.SetCellValue(r, c, "%s" % data)

        if os.path.exists(dnFile):
            df_sell = pd.read_csv(dnFile, sep=",", header=None, names=self.stock_colLabels, encoding='Big5')
            df_sell = df_sell.sort_values(by='增量', ascending=False)
            self.table_sell.ClearGrid()
            for r in range(len(df_sell)):
                for c in range(len(df_sell.columns)):
                    if c == 5:
                        data = df_sell.iloc[r,c].astype('float').round(2)
                    elif c == 3:
                        data = df_sell.iloc[r,c].astype('float').round(2)
                        if data <= -3 and data >-5:
                            self.table_sell.SetCellBackgroundColour(r, c, wx.Colour(51, 204, 204))
                            self.table_sell.SetCellFont(r, c, wx.Font(font_size, wx.SWISS, wx.NORMAL, wx.BOLD))
                        elif data <= -5:
                            self.table_sell.SetCellBackgroundColour(r, c, wx.Colour(0, 153, 0))
                            self.table_sell.SetCellFont(r, c, wx.Font(font_size, wx.SWISS, wx.NORMAL, wx.BOLD))
                        else:
                            self.table_sell.SetCellBackgroundColour(r, c, 'black')
                            self.table_sell.SetCellFont(r, c, wx.Font(font_size, wx.SWISS, wx.NORMAL, wx.NORMAL))
                    elif c == 2:
                        data = df_sell.iloc[r,c].astype('int')
                        if data >=300 and data < 600:
                            self.table_sell.SetCellBackgroundColour(r, c, 'pink')
                            self.table_sell.SetCellFont(r, c, wx.Font(font_size, wx.SWISS, wx.NORMAL, wx.BOLD))
                        elif data >= 600:
                            self.table_sell.SetCellBackgroundColour(r, c, 'red')
                            self.table_sell.SetCellFont(r, c, wx.Font(font_size, wx.SWISS, wx.NORMAL, wx.BOLD))
                        else:
                            self.table_sell.SetCellBackgroundColour(r, c, 'black')
                            self.table_sell.SetCellFont(r, c, wx.Font(font_size, wx.SWISS, wx.NORMAL, wx.NORMAL))
                    else:
                        data = df_sell.iloc[r,c]
                    self.table_sell.SetCellValue(r, c, "%s" % data)


    def onUpdateFutureOrder(self):
        today = self.time_center[:10].replace('/','')
        tmpFile1 = 'fut_lot.txt'
        tmpFile2 = 'fut_order.txt'
        lotFile = './future/' + today + '_lot.txt'
        orderFile = './future/' + today + '_order.txt'

        if os.path.exists(tmpFile1):
            shutil.copyfile(tmpFile1, lotFile)
            with open(lotFile, 'r') as f:
                lines = f.readlines()
            y_data = lines[0].split(',')
            y_data.pop()
            xy_data = list(zip(self.chart2_xData, y_data))
    
            line = wxplot.PolySpline(
                xy_data,
                colour=wx.Colour(0, 0, 255),   # Color: blue
                width=2,
            )

            chart = wxplot.PlotGraphics([line, self.baseline])
            self.lotCanvas.Draw(chart, xAxis=(0,3600))

        if os.path.exists(tmpFile2):
            shutil.copyfile(tmpFile2, orderFile)
            with open(orderFile, 'r') as f:
                lines = f.readlines()
            y_data = lines[0].split(',')
            y_data.pop()
            xy_data = list(zip(self.chart2_xData, y_data))
        
            line = wxplot.PolySpline(
                xy_data,
                colour=wx.Colour(0, 0, 255),   # Color: blue
                width=2,
            )

            chart = wxplot.PlotGraphics([line, self.baseline])
            self.orderCanvas.Draw(chart, xAxis=(0,3600))

    def onLoginSj(self, event):
        self.loginSubResponse.WriteText(f"登入帳號: {self.accountText.GetValue()}\r")
        self.loginSubResponse.WriteText(f"登入密碼: {self.passwdText.GetValue()}\r")
        '''if self.api == None:
            self.api = sj.Shioaji(simulation=True)

        try:
            #self.api.login(person_id=self.pid, passwd=self.pwid, contracts_cb=self.getContracts)
            self.api.login(api_key=self.accountText.GetValue(), secret_key=self.passwdText.GetValue(), contracts_cb=self.getContracts)
        except Exception as err:
            self.SetStatusText(f"登入永豐發生異常 {err}")

        self.SetStatusText(f"已登入永豐Api")
        self.api.quote.set_on_tick_fop_v1_callback(self.onGetFopTick)'''
            
    def onLogoutSj(self):
        if self.api != None:
            if len(self.subscribed_item) != 0:
                self.onUnsubscribeOption()

            self.api.logout()
            self.api = None
            self.SetStatusText(f"已登出永豐Api")

    def getContracts(self, security_type):
        if security_type == sj.constant.SecurityType.Index:
            self.loginSubResponse.WriteText(f"取得指數 Contracts\r")
            self.SetStatusText(f"取得指數 Contracts")
        if security_type == sj.constant.SecurityType.Stock:
            self.loginSubResponse.WriteText(f"取得股票 Contracts\r")
            self.SetStatusText(f"取得股票 Contracts")
        if security_type == sj.constant.SecurityType.Future:
            self.loginSubResponse.WriteText(f"取得期貨 Contracts\r")
            self.SetStatusText(f"取得期貨 Contracts")
        if security_type == sj.constant.SecurityType.Option:
            self.loginSubResponse.WriteText(f"取得選擇權 Contracts\r")
            self.SetStatusText(f"取得選擇權 Contracts")

    def onSubscribeFuture(self):
        symbol_list = ['idx_TSE_TSE001', 'fut_TXF_TXFR1']
        if self.api != None:
            for symbol in symbol_list:
                contract = symbol.split('_')
                ct_type = contract[0]
                ct_category = contract[1]
                ct_code = contract[2]
                if ct_type == 'idx':
                    snapshots = self.api.snapshots([self.api.Contracts.Indexs[ct_category][ct_code]])
                    data = {'symbol': self.api.Contracts.Indexs[ct_category][ct_code].symbol,
                            'code': self.api.Contracts.Indexs[ct_category][ct_code].code,
                            'name': self.api.Contracts.Indexs[ct_category][ct_code].name,
                            'Date': pd.to_datetime(snapshots[0]['ts']),
                            'Time': pd.to_datetime(snapshots[0]['ts']),
                            'ref_price': snapshots[0]['open'] + snapshots[0]['change_price'],
                            'open': snapshots[0]['open'],
                            'high': snapshots[0]['high'],
                            'low': snapshots[0]['low'],
                            'close': snapshots[0]['close'],
                            'range': format(snapshots[0]['high'] - snapshots[0]['low'], '.2%'),
                            'change_rate': snapshots[0]['change_rate'],
                            'change_price': snapshots[0]['change_price'],
                            'total_amount': snapshots[0]['total_amount'],
                            }
                    self.idxAndFut['TSE'] = data
                    tse_text = str(data['close']) + '\n' + str(data['change_rate']) + '\n' + str(data['range'])
                    self.twse_data.SetLabel(tse_text)
                    if data['close'] >= data['ref_price']:
                        self.twse_data.SetForegroundColour((255,0,0))
                    else:
                        self.twse_data.SetForegroundColour((0,204,0))
                elif ct_type == 'fut':
                    snapshots = self.api.snapshots([self.api.Contracts.Futures[ct_category][ct_code]])
                    data = {'symbol': self.api.Contracts.Futures[ct_category][ct_code].symbol,
                            'code': self.api.Contracts.Futures[ct_category][ct_code].code,
                            'name': self.api.Contracts.Futures[ct_category][ct_code].name,
                            'Date': self.api.Contracts.Futures[ct_category][ct_code].update_date,
                            'Time': pd.to_datetime(snapshots[0]['ts']),
                            'ref_price': self.api.Contracts.Futures[ct_category][ct_code].reference,
                            'open': snapshots[0]['open'],
                            'high': snapshots[0]['high'],
                            'low': snapshots[0]['low'],
                            'close': snapshots[0]['close'],
                            'range': format(snapshots[0]['high'] - snapshots[0]['low'], '.2%'),
                            'change_rate': snapshots[0]['change_rate'],
                            'avg_price': snapshots[0]['average_price'],
                            'total_volume': snapshots[0]['total_volume'],
                            'bid_total_volume': 0,
                            'ask_total_volume': 0,
                            'bid_ask_diff': 0
                            }
                    self.idxAndFut['TXFR1'] = data
                    fut_text = str(data['close'].split('.')[0]) + '\n' + str(data['change_rate']) + '\n' + str(data['range'])
                    self.fut_data.SetLabel(fut_text)
                    if data['close'] >= data['ref_price']:
                        self.fut_data.SetForegroundColour((255,0,0))
                    elif tx['close'] < tx['ref_price']:
                        self.fut_data.SetForegroundColour((0,204,0))

            spread = round(self.idxAndFut['TXFR1']['close'] - self.idxAndFut['TSE']['close'],0)
            self.diff_data.SetLabel(f'{spread}')


    def onGetFopTick(self, exchange:Exchange, tick:TickFOPv1):
        if tick.code in self.option_code_w:
            data_pos = self.option_code_w.index(tick.code)
            data = {'symbol': self.option_data_w[data_pos]['symbol'],
                    'code': self.option_data_w[data_pos]['code'],
                    'name': self.option_data_w[data_pos]['name'],
                    'strike_price': self.option_data_w[data_pos]['strike_price'],
                    'CP': self.option_data_w[data_pos]['CP'],
                    'Date': self.option_data_w[data_pos]['Date'],
                    'Time': tick.datetime.strftime('%Y/%m/%d %H:%M:%S'),
                    'ref_price': self.option_data_w[data_pos]['ref_price'],
                    'close': float(tick.close),
                    'avg_price': float(tick.avg_price),
                    'total_volume': tick.total_volume,
                    'bid_total_volume': tick.bid_side_total_vol,
                    'ask_total_volume': tick.ask_side_total_vol,
                    'bid_ask_diff': tick.bid_side_total_vol - tick.ask_side_total_vol
                    }
            self.option_data_w[data_pos] = data
            self.SetStatusText(f"更新資料:{data['Time']}")


if __name__ == '__main__':
    # When this module is run (not imported) then create the app, the
    # frame, show it, and start the event loop.
    app = wx.App()
    frm = MainFrame(None, title='招財貓賺錢系統')
    #panel = wx.Panel(frm)
    frm.Show()
    app.MainLoop()
    