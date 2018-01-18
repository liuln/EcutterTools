#-*- coding: UTF-8 -*-
import sys
import os
import logging
import wmi
import time
import threading
from winreg import *
from xml.dom.minidom import parse
from mainwindow import *
from selenium import webdriver

logger = logging.getLogger('mylogger')
logger.setLevel(logging.DEBUG)
fh = logging.FileHandler('EcutterTools.log')
fh.setLevel(logging.DEBUG)
formatter = logging.Formatter(
    '[%(asctime)s][%(thread)d][%(filename)s][line: %(lineno)d][%(levelname)s] ## %(message)s')
fh.setFormatter(formatter)
logger.addHandler(fh)

class MyApp(QtWidgets.QMainWindow, Ui_MainWindow):
    global s
    s = wmi.WMI()
    global div_gb_factor, strCpuType
    global flag
    flag =0
    # The registry (key, value) which stored in the xml file
    global regList
    regList = {}
    global palette
    palette = QtGui.QPalette()

    def getCPUSpeed(self, cpu):
        logger.info("获取CPU主频信息！")
        cpuType = str(cpu.Name)
        cpuType = cpuType.split('@', 1)
        result = float(cpuType[1].split('GHz', 1)[0])
        return result

    def getCPUType(self, cpu):
        logger.info("获取CPU型号信息！")
        cType = str(cpu.Name)
        cType = cType.split('@', 1)
        return cType

    def getCPUNum(self, cpu, cpuNum):
        logger.info("获取CPU的核数信息！")
        coreNum = str(cpu.NumberOfCores * cpuNum)
        return coreNum

    def setErrorText(self,controlName,text, errMsg=""):
        palette.setColor(QtGui.QPalette.WindowText, QtCore.Qt.red)
        controlName.setPalette(palette)
        controlName.setText(text + errMsg)

    def setText(self, controlName, text):
        controlName.setText(text)

    def getCpuInfo(self, cpus):
        logger.info("获取CPU硬件信息！")
        self.strCpuType = "初始状态"
        for cpu in cpus:
            speed = self.getCPUSpeed(cpu)
            cpuType = self.getCPUType(cpu)
            logger.info("判断CPU主频是否满足需求！")
            text = cpuType[1].strip()
            errMsg = u", 请更换主频更高的CPU!"
            if speed < 2.0:
                logger.info("CPU主频小于2.0GHz！")
                self.isNotUsable()
                self.setErrorText(self.label_hardware_CPUspeed_result,text, errMsg)
            else:
                logger.info("CPU主频大于等于2.0GHZ!")
                self.setText(self.label_hardware_CPUspeed_result, text)
            logger.info("处理多个CPU的情况！")
            if self.strCpuType == "初始状态":
                logger.info("设置CPU型号为: %s", cpuType[0])
                text = cpuType[0]
                self.setText(self.label_hardware_CPUmodel_result,text)
                self.strCpuType = cpuType
            elif self.strCpuType != cpuType:
                logger.info("存在CPU型号不一致，设置CPU的类型为其他")
                text = u"CPU型号不一致！"
                self.setErrorText(self.label_hardware_CPUmodel_result,text)
                self.strCpuType = "other"
            else:
                logger.info("存在多个相同型号的CPU")
                self.strCpuType = cpuType
        logger.info("获取CPU的核数！")
        coreNum = self.getCPUNum(cpu,len(cpus))
        if int(coreNum) < 4:
            logger.info("CPU的核数小于4，要设置红色字提示当前计算机的硬件环境信息不可以使用非编软件！")
            self.isNotUsable()
            self.setErrorText(self.label_hardware_CPUnum_result,coreNum,"核数小于4")
        logger.info("CPU核数满足使用非编软件的最低硬件环境需求！")
        self.setText(self.label_hardware_CPUnum_result,coreNum)

    def getMemoryInfo(self):
        logger.info("获取内存信息！")
        div_gb_factor = (1024.0 ** 3)
        for mem in s.Win32_ComputerSystem():
            memCap = round(int(mem.TotalPhysicalMemory)/div_gb_factor)
            logger.info("判断内存是否满足要求")
            text = str(memCap) + 'GB'
            if memCap < 8:
                logger.info("内存小于8G")
                self.isNotUsable()
                self.setErrorText(self.label_hardware_memory_result,text,"内存小于8G")
            logger.info("内存大于等于8G")
            self.setText(self.label_hardware_memory_result,text)

    def getGPUInfo(self):
        logger.info("获取显卡信息！")
        div_gb_factor = (1024 ** 3)
        for gpu in s.Win32_VideoController():
            gpuMemCap = gpu.CurrentNumberOfColors
            if gpuMemCap is None:
                logger.info("目前获取到的显卡的显存为零")
                continue
            else:
                gpuMemCap = int(gpuMemCap) / div_gb_factor
                text = str(gpuMemCap) + 'GB'
                if int(gpuMemCap) < 1:
                    logger.info("显存小于1GB")
                    self.isNotUsable()
                    self.setErrorText(self.label_hardware_GPUmemory_result,text,u"显存小于1G")
                else:
                    logger.info("显存大于等于1GB")
                    self.setText(self.label_hardware_GPUmemory_result, text)
                gpuInfo = gpu.Name
                logger.info("设置显卡型号信息")
                self.setText(self.label_hardware_GPUmodel_result,gpuInfo)
                #self.label_hardware_GPUmodel_result.setText(gpuInfo)
                return
        self.isNotUsable()
        # self.updateUI()
        self.setErrorText(self.label_hardware_GPUmodel_result,"请安装独立显卡！")
        #palette.setColor(QtGui.QPalette.WindowText, QtCore.Qt.red)
        #self.label_hardware_GPUmodel_result.setPalette(palette)
        #self.label_hardware_GPUmodel_result.setText("请安装独立显卡！")

    def downloadFile(self,path):
        browser = webdriver.Ie()
        browser.get(path)
        time.sleep(3)
        #browser.find_element_by_xpath("/html/body/div[2]/div[1]/div[1]/div[1]/div/div[2]/div/div/div[2]/a[2]/span/span").click()
        #手动下载

    def installFile(self, path):
        self.downloadFile(path)
        #手动安装

    def isFileInstalled(self, path, name, obj, insObj,key_name):
        index = path.find(":")
        isInstall = None
        if index != -1:
            isInstall = self.isFileExisted(path,name,obj,insObj)
        else:
            isInstall = self.isSoftwareInstalled(path,name,obj,insObj,key_name)
        return isInstall

    def isFileExisted(self, fPath, fileName, obj, insObj):
        logger.info("检查文件是否存在")
        isExisted = os.path.exists(fPath+"\\"+fileName)
        if not isExisted:
            logger.info("文件不存在")
            self.setErrorText(obj, r"未安装")
            insObj.setEnabled(True)
            self.isNotUsable()
            return False
        else:
            logger.info("文件存在")
            insObj.setEnabled(False)
            self.setText(obj,r"已安装")
            return True

    def isSoftwareInstalled(self, rPath, softName, obj, insObj,key_name):
        logger.info("检查软件是否安装")
        value = self.QueryReg(rPath, softName,key_name)
        if value == True:
            logger.info("软件已经安装")
            insObj.setEnabled(False)
            obj.setText(r"已安装")
        else:
            logger.info("软件未安装")
            self.isNotUsable()
            self.setErrorText(obj, r"未安装")

    def isNotUsable(self):
        logger.info("进入软件不可使用处理函数")
        strUsable = self.label_usable_result.text()
        if strUsable == r"可以使用":
            logger.info("可以使用-->无法使用")
            self.setErrorText(self.label_usable_result,r"无法使用,请关注标红部分！")
            self.pushButton.setEnabled(False)

    def isOptimize(self):
        logger.info("进入软件优化处理函数")
        strOptimize = self.label_optimize_result.text()
        if strOptimize == r"无需优化":
            logger.info("无需优化-->需要")
            self.setErrorText(self.label_optimize_result,r"需要优化")

    def updateUI(self):
        logger.info("更新界面信息")
        logger.info("设置非编是否可以使用")
        strUsable = self.label_usable_result.text()
        logger.info("设置非编是否需要优化")
        strOptimize = self.label_optimize_result.text()
        logger.info("设置一键优化按钮状态")
        if (strUsable == r"可以使用") and (strOptimize == r"无需优化"):
            self.pushButton.setEnabled(False)
        if strUsable.find(r"无法使用") == 0 or strOptimize == r"需要优化":
            self.pushButton.setEnabled(True)
        if strOptimize == r"优化完毕！":
            self.pushButton.setEnabled(False)

    def checkRegistryInfo(self, rPath , cpuTypeList):
        logger.info("检测注册表信息")
        #self.cpuTypeList = cpuTypeList
        regList = {}
        for cpuType in cpuTypeList:
            #Judge the cputype of the target machine is in the cpuTypeList, if yes then save in the global variable
            if self.strCpuType[0].find(cpuType) != -1:
                logger.info("模板文件中指定CPU类型的预制注册表值读取")
                regList = self.getListData(cpuType)
                self.regList = regList
                break
            else:
                continue
        for rKey in regList:
            regData = self.readFromReg(rPath, rKey)
            if regData == None or regData is None:
                logger.info("本机注册表值未读取到数据")
                self.isOptimize()
                #self.updateUI()
            elif str(regData) != str(regList[rKey]):
                logger.info("本机注册表值读取的数据和预制模板中读取的数据不一致")
                self.isOptimize()

    def getListData(self,cpuType):
        logger.info("从模板文件中根据输入的CPU类型来获取相关的注册表数据")
        listData = {}
        #root = dom.documentElement
        cType = dom.getElementsByTagName('CPUType')
        for type in cType:
            if cpuType == type.getAttribute("ctype"):
                # get all the child nodes
                for node in type.childNodes:
                    if node.nodeType == node.ELEMENT_NODE:
                        listData[node.nodeName] = node.firstChild.data
                return listData
            logger.info("返回注册表值列表")

    def QueryReg(self, rPath,value_name,key_name):
        logger.info("查询注册表中某个项的数值")
        keylist = []
        try:
            key = OpenKey(HKEY_LOCAL_MACHINE, rPath, 0, KEY_READ)
            subKey = QueryInfoKey(key)[0]
        except:
            logger.info("要查询的注册表路径不存在")
            return
        logger.info("获取注册表指定路径下的所有项")
        for i in range(int(subKey)):
            keylist.append(EnumKey(key, i))
        CloseKey(key)
        logger.info("获取指定路径下所有子项的路径下的" + key_name + "项的值")
        for i in keylist:
            tPath = rPath + "\\" + i
            rValue = self.readFromReg(tPath, key_name)
            if value_name == rValue:
                logger.info("注册表中获取到的值跟指定的值一致")
                return True
        logger.info("未查到指定路径下的跟传递过来的参数一样的值")
        return None

    def readFromReg(self,rPath, value_name):
        logger.info("读注册表中某个项的数值")
        try:
            key = OpenKey(HKEY_LOCAL_MACHINE, rPath, 0, KEY_READ)
            value,type = QueryValueEx(key, value_name)
            return value
        except:
            logger.info("注册表的路径或者项不存在")
            return

    def updateReg(self, value_name, value, regPath):
        logger.info("更新指定路径下某个项中的值")
        # Read from registry to see if the value_name is existed
        rValue = self.readFromReg(regPath, value_name)
        try:
            newKey = CreateKeyEx(HKEY_LOCAL_MACHINE, regPath, 0, KEY_WRITE)
        except:
            logger.info("指定路径不存在！")
            return
        value = int(value)
        logger.info("设置注册表某项的值")
        if rValue == None or rValue is None:
            SetValueEx(newKey, value_name, 0, REG_DWORD, value)
        if rValue != value:
            SetValueEx(newKey, value_name, 0, REG_DWORD, value)

    def optimizeOneKey(self, rPath):
        logger.info("一键优化按钮处理函数")
        for rKey in self.regList:
            self.updateReg(rKey, self.regList[rKey],rPath)
        self.setErrorText(self.label_optimize_result, r"优化完毕！")
        self.updateUI()

    def __init__(self):
        logger.info("QMainWindow初始化")
        QtWidgets.QMainWindow.__init__(self)
        logger.info("Ui_MainWindow初始化")
        Ui_MainWindow.__init__(self)
        logger.info("MyApp setupUi...")
        self.setupUi(self)
        logger.info("Load UI...")
        #self.loadUI()

    def loadHardUI(self):
        logger.info("获取CPU信息...")
        self.getCpuInfo(s.Win32_Processor())
        logger.info("获取内存信息...")
        self.getMemoryInfo()
        logger.info("获取GPU信息...")
        self.getGPUInfo()

    def loadSoftUI(self):
        key_name ="DisplayName"
        logger.info("判断DirectX 10是否安装")
        isInstalled = self.isFileInstalled("C:\\WINDOWS\\system32", "D3DX11_41.dll", self.label_software_Directx10_result, self.pushButton_Directx10,key_name)
        if isInstalled != True:
            #self.pushButton_Directx10.clicked.connect(lambda:self.installFile("https://pan.baidu.com/s/1qYmh2he"))
            path = "https://pan.baidu.com/s/1qYmh2he"
            self.pushButton_Directx10.clicked.connect(lambda: threading.Thread(target=self.installFile, args=(path,)).start())
        logger.info("判断DirectX 11是否安装")
        isInstalled = self.isFileInstalled("C:\\WINDOWS\\system32", "D3DX11_43.dll", self.label_software_Directx11_result, self.pushButton_Directx11,key_name)
        if isInstalled != True:
            self.pushButton_Directx11.clicked.connect(lambda:self.installFile("https://pan.baidu.com/s/1qYyYGqg"))
        logger.info("判断VC2005 x64是否安装")
        isInstalled = self.isFileInstalled("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall","Microsoft Visual C++ 2005 Redistributable (x64)", self.label_software_vc2005x64_result,self.pushButton_vc2005x64,key_name)
        if isInstalled != True:
            self.pushButton_vc2005x64.clicked.connect(lambda:self.installFile("https://pan.baidu.com/s/1csFDO6"))
        logger.info("判断Setup1是否安装")
        isInstalled = self.isFileInstalled("SOFTWARE\\Wow6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall","Setup1", self.label_software_setup1_result, self.pushButton_setup1,key_name)
        if isInstalled != True:
            self.pushButton_setup1.clicked.connect(lambda:self.installFile("https://pan.baidu.com/s/1hsJ7Kxu"))
        key_name = ""
        logger.info("判断QuickTime是否安装")
        isInstalled = self.isFileInstalled("SOFTWARE\\Clients\\Media","QuickTime",self.label_software_Quciktime_result,self.pushButton_Quciktime, key_name)
        if isInstalled != True:
            self.pushButton_Quciktime.clicked.connect(lambda:self.installFile("https://pan.baidu.com/s/1kVcjuxp"))

    def loadRegUI(self):
        logger.info("检查注册表信息")
        cpuTypeList = ['1620', '2609', '2620', '2630', '2650', 'other']
        self.checkRegistryInfo("SOFTWARE\\Dayang\\SoftCodec",cpuTypeList)
        logger.info("更新UI")
        self.updateUI()

    def loadOptkeyUI(self):
        logger.info("点击一键优化按钮")
        self.pushButton.clicked.connect(lambda:self.optimizeOneKey("SOFTWARE\\Dayang\\SoftCodec"))

if __name__ == "__main__":
    logger.info("Enter into main func...")
    regTempPath = os.path.abspath(os.path.dirname(sys.argv[0]))
    dom = parse(regTempPath + '\\registryTemplate.xml')
    logger.info("QApplication...")
    app = QtWidgets.QApplication(sys.argv)
    logger.info("MyApp...")
    window = MyApp()
    logger.info("MyApp show...")
    window.loadHardUI()
    window.loadSoftUI()
    window.loadRegUI()
    window.loadOptkeyUI()
    window.show()
    logger.info("QApplication exit...")
    sys.exit(app.exec_())
