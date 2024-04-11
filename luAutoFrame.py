
import sys,requests,time,threading,os,json
from PyQt5.QtWidgets import QMainWindow,QApplication,QWidget,QMessageBox,QTreeWidget,QTreeWidgetItem
from PyQt5.QtCore import Qt,pyqtSignal
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service

from Ui_luAutoFrame import Ui_Form
 
# 定义全局常量
ISRUN = 0       # ISRUN == 1是运行，ISRUN == 0是停止
DRIVERDIR =""       # 保存Chrome驱动的路径

class luAutoFrame( QMainWindow, Ui_Form): 
    #定义一个信号，用于子线程给主线程发信号
    signalThread = pyqtSignal(str, str)     #两个str参数，第一个接收信号类型，是QMessageBox提示,还teDisplay显示

    def __init__(self,parent =None):
        super( luAutoFrame,self).__init__(parent)
        self.setupUi(self)

        #打开配置文件，初始化界面数据
        if os.path.exists( "./lu.ini"):
            try:
                iniFileDir = os.getcwd() + "\\"+ "lu.ini"
                with open( iniFileDir, 'r', encoding="utf-8") as iniFile:
                    iniDict = json.loads( iniFile.read())
                if iniDict:
                    self.leURL.setText( iniDict['URL'])
                    self.sbX.setValue( iniDict['sbX'])
                    self.sbY.setValue( iniDict['sbY'])
                    self.sbWidth.setValue( iniDict['sbWidth'])
                    self.sbHeight.setValue( iniDict['sbHeight'])
            except:
                QMessageBox.about( self, "提示", "打开初始化文件lu.ini异常")

        #初始化twBrowser表头
        self.twBrowser.headerItem().setText( 0, "分组:(序号--名称--ID)")

        #绑定槽函数
        self.btnConnect.clicked.connect( self.mfConnect)    # 点击连接/刷新按键
        self.twBrowser.itemClicked.connect( self.mftwClicked)       # twBrowser中父节点选中时，自动选中所有子节点
        self.btnClear.clicked.connect( self.mfteClear)      # 清空teDisplay中的内容
        self.btnStop.clicked.connect( self.mfStop)      # 点击停止按钮，在执行完当前任务后停止
        self.btnStart.clicked.connect( self.mfStart)        #开始自动运行
        self.signalThread.connect( self.mfSignal)       # 处理子线程给主线程发的信号


    #通过twBrowser中的第一个AdsPower窗口账号，获取Chrome驱动文件的位置，保存在全局DRIVERDIR中
    def mfFindDRIVERDIR( self):
        global DRIVERDIR
        if self.twBrowser.topLevelItemCount() != 0:
            if self.twBrowser.topLevelItem(0).childCount() != 0:
                browserSN, browserName, browserID = self.twBrowser.topLevelItem(0).child(0).text(0).split("--")
                #通过ID查询账号信息
                queryURL = self.leURL.text() + "/api/v1/browser/start"
                #打开参数需要改一下，默认打开的窗口会有很多标签，改成只打开一个标签，提升速度
                #"open_tabs"  是否打开平台和历史页面，0:打开(默认)，1:不打开
                #"ip_tab"  是否打开ip检测页，0:不打开，1:打开(默认), AdsPower需升级到V2.5.7.9及以上版本
                queryParams = { "user_id" : browserID, "open_tabs" : 1}
                try:
                    queryJson = requests.get( queryURL, queryParams).json()
                    if queryJson['code'] == 0:
                        DRIVERDIR = queryJson['data']['webdriver']
                    else:
                        QMessageBox.about( self, "提示", "根据ID查询Chrome驱动位置失败，request 'code' != 0")
                except:
                    QMessageBox.about( self, "提示", "根据ID查询Chrome驱动位置出现异常")
                
                #关闭窗口
                try:
                    closeURL = self.leURL.text() + "/api/v1/browser/stop"
                    closeParams = {'user_id' : browserID}
                    time.sleep(1.2)
                    closeJson = requests.get( closeURL, params=closeParams).json()
                    if closeJson['code'] == 0:
                        closeSuccess = "关闭账号成功    序号: " + browserSN + "  名称: " + browserName + "  ID: " + browserID
                        self.signalThread.emit( "Display", closeSuccess)
                    else:
                        closeError = "关闭账号失败    序号: " + browserSN + "  名称: " + browserName + "  ID: " + browserID
                        self.signalThread.emit( "Display", closeError)
                except:
                    self.signalThread.emit( "QMessageBox", "关闭账号时request出现异常")
            else:
                QMessageBox.about( self, "提示", "没有可用的账号，无法获取Chrome驱动的位置")
        else:
                QMessageBox.about( self, "提示", "没有可用的分组，无法获取Chrome驱动的位置")


    # 从AdsPower获取分组和浏览器列表，刷新twBrowser控件
    def mfConnect( self): 
        self.twBrowser.clear()
        URL = self.leURL.text()
        groupsDict = {}     #存储twBrowser中根节点实例的字典， 分组名称:控件实例
        
        # 查询所有分组
        try:
            groupsURL = URL + "/api/v1/group/list"
            groupsParams = {'page_size':100}    #定义查询参数，使查询1页显示100条信息，也可以用page定义查询第几页
            groups = requests.get( groupsURL, params= groupsParams).json()
            if groups['code'] == 0:
                # 向twBrowser控件添加分组做为根节点
                for i in groups['data']['list']:    # i['group_name']是分组名称， i['group_id']是分组ID
                    tempTreeRoot = QTreeWidgetItem( self.twBrowser)
                    tempTreeRoot.setText( 0, i['group_name'] + "--" + i['group_id'])
                    tempTreeRoot.setCheckState( 0, 0)
                    groupsDict[i['group_name']] = tempTreeRoot

            else:
                QMessageBox.about( self, "提示", "获取分组失败")
            
        except:
            QMessageBox.about( self, "提示", """获取分组信息时失败，由于目标计算机积极拒绝，无法连接。
            请重启AdsPower并确定接口地址""")
            return


        #查询所有账号，并添加到twBrowser
        try:
            browsersURL = URL + "/api/v1/user/list"
            browsersParams = { 'page_size':100}     #查询所有账号的第1页的前100个账号
            time.sleep(1.2)     # AdsPower接口要求每秒只能访问1次
            browsers = requests.get( browsersURL, params=browsersParams).json()
            
            if browsers['code'] == 0:
                for i in browsers['data']['list']:
                    tempTreeChild = QTreeWidgetItem( groupsDict[ i['group_name']])
                    tempTreeChild.setText( 0, str(i['serial_number']) + "--" + i['name'] + "--" + i['user_id'])
                    tempTreeChild.setCheckState( 0, 0)
                    #网上教程需要 groupsDict[ i['group_name']].addChild( tempTreeChild), 没有这一句也可以
                
                self.twBrowser.expandAll()      # 展开所有的树形节点

            else:
                QMessageBox.about( self, "提示", "获取所有账号信息失败")

        except:
            QMessageBox.about( self, "提示", """获取所有账号时失败，由于目标计算机积极拒绝，无法连接。
            请重启AdsPower并确定接口地址""")
            return
        
        #初始化twBrowser之后，调用函数获得DRIVERDIR
        global DRIVERDIR
        self.mfFindDRIVERDIR()

        
    # twBrowser中父节点选中时，自动选中所有子节点
    def mftwClicked( self, item, column):
        count = item.childCount()       # 获得子节点数量
        state = item.checkState( 0)     # 获得子节点状态
        if count != 0:      #点击的节点不是子节点时，改变其子节点的状态
            for i in range(count):
                item.child(i).setCheckState( 0, state)

    # 清空teDisplay中的内容
    def mfteClear( self):
        self.teDisplay.clear()

    # 处理子线程给主线程发的信号, 信号signalType是字符串'QMessageBox' 'Display'
    def mfSignal( self, signalType, content):
        if signalType == 'QMessageBox':
            QMessageBox.about( self, "提示", content)
        elif signalType == 'Display':
            self.teDisplay.append( content)

    # 点击停止按钮，在执行完当前任务后停止
    def mfStop( self):
        global ISRUN
        ISRUN = 0

    # 开始自动运行
    def mfStart( self):
        global ISRUN
        ISRUN = 1
        # 创建一个新线程来运行自动化操作
        inRunThreading = threading.Thread( target= self.mfRun)
        inRunThreading.start()
        # inRunThreading.join()     join()是将指定线程加入当前线程，将两个交替执行的线程转换成顺序执行。
        #                           把inRunThreading加入到主线程，会引起主线程阻塞，这里不能.join()


    # 核心代码， 执行自动化**********************************
    def mfRun(self):
        global DRIVERDIR
        global ISRUN
        URL = self.leURL.text()     #AdsPower接口URL
        browserList = []        #存储被选择的账号，格式为  序号--名称--ID
        # browsersList = self.twBrowser.selectedItems()  这行是错的，selectedItems()是鼠标选中的行，不是复选框打对号的行
        # 遍历twBrowser，获得复选框打对号的行 填入browserList[]
        rootCount = self.twBrowser.topLevelItemCount()      # 获得根节点数量
        for i in range( 0, rootCount):      #遍历根节点
            rootItem = self.twBrowser.topLevelItem( i)
            childCount = rootItem.childCount()      #获得子节点数量
            for j in range( 0, childCount):         #遍历子节点
                if rootItem.child( j).checkState( 0) == Qt.Checked:  #检查子节点选中状态，把被选中的添加到browserList[]
                    browserList.append( rootItem.child( j).text( 0))

        if len( browserList) == 0:      #如果没有选择账号，则给主线程发送信号，在QMessageBox中提示
            self.signalThread.emit( 'QMessageBox', "请选择要打开的账号")
            
    
        #核心代码， 打开浏览器并自动化操作******************************
        for i in browserList:
            if ISRUN == 0:
                break       #这里用break还是return ？

            browserSN, browserName, browserID = i.split('--')
            debuggerAddress = ""        # 接口地址

            #启动浏览器
            #先检查浏览器的启动状态，如果已启动，则直接获得接口。如果未启动，则通过API获得接口
            stateURL = URL + "/api/v1/browser/active"
            stateParams = { 'user_id' : browserID}
            try:
                time.sleep(1.2)
                stateJson = requests.get( stateURL, params=stateParams).json()
                if stateJson['code'] != 0:      # 0为成功， -1为失败
                    self.signalThread.emit( "QMessageBox", "检查账号状态时request失败，返回错误代码 code == -1或其它")
                    break
                else:
                    #如果账号已打开则直接获得接口地址，如果未打开则调用request打开账号然后获得接口地址
                    if stateJson['data']['status'] == "Active":     
                        debuggerAddress = stateJson["data"]["ws"]["selenium"]
                        
                    elif stateJson['data']['status'] == "Inactive":
                        openURL = URL +  "/api/v1/browser/start"
                        openParams = {"user_id" : browserID, "open_tabs" : 1}
                        time.sleep(1.2)
                        openJson = requests.get( openURL, params=openParams).json()
                        if openJson['code'] != 0:       # 0为成功， -1为失败
                            self.signalThread.emit( "QMessageBox", "启动账号时request失败，返回错误代码 code == -1或其它")
                            break
                        else:
                            debuggerAddress = openJson["data"]["ws"]["selenium"]

            except:
                exceptStateStr = "检查账号状态时出现异常  " + browserName + "--" + browserID
                self.signalThread.emit( "QMessageBox", exceptStateStr)
                break

            # 初始化selenium调用chrome
            chrome_options = Options()
            chrome_options.add_experimental_option("debuggerAddress", debuggerAddress)
            # browserDriver as bd, 下面这一行注意，由于selenium更新，第一个参数原来是直接写chrome驱动的路径
            #现在要求传递一个Service对象，改为下面这个写法
            bd = webdriver.Chrome(service = Service(DRIVERDIR), options=chrome_options)  
            bd.set_window_rect( x=self.sbX.value(), y=self.sbY.value(),         #设置窗口位置，尺寸
                width=self.sbWidth.value(), height=self.sbHeight.value())

            #######################################################
            #从这里开始是操作浏览器的脚本，根据不同的网站，写不同的操作#
            #######################################################
            bd.get("https://www.baidu.com")



            #######################################################
            # 关闭chrome窗口。并调用接口，关闭账号
            time.sleep(1.2)
            bd.quit()
            try:
                closeURL = URL + "/api/v1/browser/stop"
                closeParams = {'user_id' : browserID}
                time.sleep(1.2)
                closeJson = requests.get( closeURL, params=closeParams).json()
                if closeJson['code'] == 0:
                    closeSuccess = "关闭账号成功    序号: " + browserSN + "  名称: " + browserName + "  ID: " + browserID
                    self.signalThread.emit( "Display", closeSuccess)
                else:
                    closeError = "关闭账号失败    序号: " + browserSN + "  名称: " + browserName + "  ID: " + browserID
                    self.signalThread.emit( "Display", closeError)

            except:
                self.signalThread.emit( "QMessageBox", "关闭账号时request出现异常")



#主程序入口
if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWin = luAutoFrame()
    myWin.show()

    appExit = app.exec_()
    #退出程序之前，保存界面上的设置
    tempDict = { 'URL':myWin.leURL.text(), 'sbX':myWin.sbX.value(), 'sbY':myWin.sbY.value(),
                    'sbWidth':myWin.sbWidth.value(), 'sbHeight':myWin.sbHeight.value()}
    saveIniJson = json.dumps( tempDict, indent=4)
    try:
        saveIniFile = open( "./lu.ini", "w",  encoding="utf-8")
        saveIniFile.write( saveIniJson)
        saveIniFile.close()
    except:
        QMessageBox.about( myWin, "提示", "保存配置文件lu.ini失败")

    sys.exit( appExit)
# sys.exit(app.exec_())  