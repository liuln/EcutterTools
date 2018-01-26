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
import win32process
import win32event
import win32api

logger = logging.getLogger('mylogger')
logger.setLevel(logging.DEBUG)
fh = logging.FileHandler('Log.log','w')
fh.setLevel(logging.DEBUG)
formatter = logging.Formatter(
    '[%(asctime)s][%(thread)d][%(filename)s][line: %(lineno)d][%(levelname)s] ## %(message)s')
fh.setFormatter(formatter)
logger.addHandler(fh)

global GUSEABLE, GOPTIMIZE
GOPTIMIZE = True
# The registry (key, value) which stored in the xml file
global regList
regList = {}
global strCpuType
strCpuType = "初始状态"
global singleUse
singleUse = {}

class MyApp(QtWidgets.QMainWindow, Ui_MainWindow):
    palette = QtGui.QPalette()
    global div_gb_factor, strCpuType
    div_gb_factor = (1024 ** 3)
    s = wmi.WMI()
    global regList, strCpuType

    def __init__(self):
        logger.info("QMainWindow初始化")
        QtWidgets.QMainWindow.__init__(self)
        logger.info("Ui_MainWindow初始化")
        Ui_MainWindow.__init__(self)
        logger.info("MyApp setupUi...")
        self.setupUi(self)
        logger.info("Load UI...")

    def setErrorText(self, controlName, text, errMsg=""):
        self.palette.setColor(QtGui.QPalette.WindowText, QtCore.Qt.red)
        controlName.setPalette(self.palette)
        controlName.setText(text + errMsg)

    def setText(self, controlName, text):
        self.palette.setColor(QtGui.QPalette.WindowText, QtCore.Qt.black)
        controlName.setPalette(self.palette)
        controlName.setText(text)

    def getListData(self, cpuType):
        logger.info("从模板文件中获取 %s 的注册表数据", cpuType)
        listData = {}
        # root = dom.documentElement
        cType = dom.getElementsByTagName('CPUType')
        for type in cType:
            if cpuType == type.getAttribute("ctype"):
                # get all the child nodes
                for node in type.childNodes:
                    if node.nodeType == node.ELEMENT_NODE:
                        listData[node.nodeName] = node.firstChild.data
                return listData
            logger.info("返回注册表值列表")

    def setSingleUse(self, obj, isUsable):
        global singleUse
        logger.info("保存%s的检测结果",obj)
        singleUse[obj] = isUsable

    def getSingleUse(self,obj):
        global singleUse
        logger.info("获取%s的检测结果",obj)
        return singleUse[obj]

    def setOptimize(self, isOptimize):
        logger.info("需要优化标签的设置")
        #global GOPTIMIZE
        #GOPTIMIZE = (GOPTIMIZE and isOptimize)
        strOptimize = self.label_optimize_result.text()
        if (isOptimize == False) and (strOptimize == r"无需优化"):
            logger.info("无需优化-->需要")
            self.setErrorText(self.label_optimize_result, r"需要优化")

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

    def getCpuInfo(self, cpus):
        logger.info("获取CPU硬件信息！")
        global strCpuType
        for cpu in cpus:
            speed = self.getCPUSpeed(cpu)
            cpuType = self.getCPUType(cpu)
            logger.info("判断CPU主频是否满足需求！")
            text = cpuType[1].strip()
            errMsg = u", 请更换主频更高的CPU!"
            if speed < 2.0:
                logger.info("CPU主频小于2.0GHz！")
                # self.isNotUsable()
                self.setSingleUse("Speed", False)
                self.setErrorText(self.label_hardware_CPUspeed_result, text, errMsg)
            else:
                logger.info("CPU主频大于等于2.0GHZ!")
                self.setSingleUse("Speed", True)
                self.setText(self.label_hardware_CPUspeed_result, text)
            logger.info("处理多个CPU的情况！")
            if strCpuType == "初始状态":
                logger.info("设置CPU型号为: %s", cpuType[0])
                text = cpuType[0]
                self.setSingleUse("CPU型号", True)
                self.setText(self.label_hardware_CPUmodel_result, text)
                strCpuType = cpuType
            elif strCpuType != cpuType:
                logger.info("存在CPU型号不一致，设置CPU的类型为其他")
                text = u"CPU型号不一致！"
                self.setSingleUse("CPU型号", False)
                self.setErrorText(self.label_hardware_CPUmodel_result, text)
                strCpuType = "other"
            else:
                logger.info("存在多个相同型号的CPU")
                self.setSingleUse("CPU型号", True)
                strCpuType = cpuType
        logger.info("获取CPU的核数！")
        coreNum = self.getCPUNum(cpu, len(cpus))
        if int(coreNum) < 4:
            logger.info("CPU的核数小于4，要设置红色字提示当前计算机的硬件环境信息不可以使用非编软件！")
            self.setSingleUse("Core", False)
            self.setErrorText(self.label_hardware_CPUnum_result, coreNum, "核数小于4")
        else:
            logger.info("CPU核数满足使用非编软件的最低硬件环境需求！")
            self.setSingleUse("Core",True)
            self.setText(self.label_hardware_CPUnum_result, coreNum)

    def getMemoryInfo(self):
        logger.info("获取内存信息！")
        div_gb_factor = (1024.0 ** 3)
        for mem in self.s.Win32_ComputerSystem():
            memCap = round(int(mem.TotalPhysicalMemory) / div_gb_factor)
            logger.info("判断内存是否满足要求")
            text = str(memCap) + 'GB'
            if memCap < 8:
                logger.info("内存小于8G")
                # self.isNotUsable()
                self.setSingleUse("Memory",False)
                self.setErrorText(self.label_hardware_memory_result, text, "内存小于8G")
            else:
                logger.info("内存大于等于8G")
                self.setSingleUse("Memory",True)
                self.setText(self.label_hardware_memory_result, text)

    def getGPUInfo(self):
        logger.info("获取显卡信息！")
        for gpu in self.s.Win32_VideoController():
            gpuMemCap = gpu.CurrentNumberOfColors
            if gpuMemCap is None:
                logger.info("目前获取到的显卡的显存为零")
                self.setSingleUse("GPUType",False)
                self.setErrorText(self.label_hardware_GPUmodel_result, "请安装独立显卡！")
                continue
            else:
                gpuMemCap = int(gpuMemCap) / div_gb_factor
                text = str(gpuMemCap) + 'GB'
                if int(gpuMemCap) < 1:
                    logger.info("显存小于1GB")
                    # self.isNotUsable()
                    self.setSingleUse("GPU",False)
                    self.setErrorText(self.label_hardware_GPUmemory_result, text, u"显存小于1G")
                else:
                    logger.info("显存大于等于1GB")
                    self.setSingleUse("GPU",True)
                    self.setText(self.label_hardware_GPUmemory_result, text)
                gpuInfo = gpu.Name
                logger.info("设置显卡型号信息")
                if (gpuInfo is None):
                    logger.info("显卡型号信息获取不正常")
                    self.setSingleUse("GPUType",False)
                    self.setErrorText(self.label_hardware_GPUmodel_result, u"未获取显卡型号信息")
                else:
                    logger.info("显卡型号信息获取正常")
                    self.setSingleUse("GPUType",True)
                    self.setText(self.label_hardware_GPUmodel_result, gpuInfo)

    def loadHardUI(self):
        logger.info("获取CPU信息...")
        self.getCpuInfo(self.s.Win32_Processor())
        logger.info("获取内存信息...")
        self.getMemoryInfo()
        logger.info("获取GPU信息...")
        self.getGPUInfo()

    def checkRegistryInfo(self, rPath, cpuTypeList):
        global regList
        for cpuType in cpuTypeList:
            # Judge the cputype of the target machine is in the cpuTypeList, if yes then save in the global variable
            if strCpuType[0].find(cpuType) != -1:
                logger.info("模板文件中CPU类型为 %s 的预制注册表值读取", cpuType)
                regList = self.getListData(cpuType)
                break
            else:
                continue
        for rKey in regList:
            regData = self.readFromReg(rPath, rKey)
            if regData == None or regData is None:
                logger.info("注册表路径 %s值未读取到 %s 的数据", rPath, rKey)
                self.setOptimize(False)
            elif str(regData) != str(regList[rKey]):
                logger.info("注册表值读取的数据为：%s，预制模板中读取的数据为：%s，两者不一致",str(regData),str(regList[rKey]))
                self.setOptimize(False)
            else:
                logger.info("注册表值读取的数据为：%s，预制模板中读取的数据为：%s，两者一致", str(regData), str(regList[rKey]))
                self.setOptimize(True)

    def QueryReg(self, rPath, value_name, key_name):
        logger.info("查询注册表中 %s的数值", value_name)
        keylist = []
        try:
            key = OpenKey(HKEY_LOCAL_MACHINE, rPath, 0, KEY_READ)
            subKey = QueryInfoKey(key)[0]
        except:
            logger.info("要查询的注册表路径 %s 不存在", rPath)
            return
        logger.info("获取注册表%s下的所有项", rPath)
        for i in range(int(subKey)):
            keylist.append(EnumKey(key, i))
        CloseKey(key)
        logger.info("获取指定路径下所有子项的路径下的" + key_name + "项的值")
        for i in keylist:
            tPath = rPath + "\\" + i
            # print("tPath is: " + tPath)
            rValue = self.readFromReg(tPath, key_name)
            rValue = str(rValue)
            if (rValue.find(value_name) != -1):
                logger.info("注册表中获取到的值: %s, 指定的值: %s，两者一致",rValue, value_name)
                return True
        logger.info("未查到指定路径下的 %s 的值", key_name)
        return None

    def readFromReg(self, rPath, value_name):
        try:
            key = OpenKey(HKEY_LOCAL_MACHINE, rPath, 0, KEY_READ)
            value, type = QueryValueEx(key, value_name)
            logger.info("读注册表中%s的数值", value)
            return value
        except:
            logger.info("注册表的路径或者项不存在")
            return None

    def updateReg(self, value_name, value, regPath):
        logger.info("更新%s路径下%s的值",regPath,value_name)
        # Read from registry to see if the value_name is existed
        rValue = self.readFromReg(regPath, value_name)
        try:
            newKey = CreateKeyEx(HKEY_LOCAL_MACHINE, regPath, 0, KEY_WRITE)
        except:
            logger.info("指定路径%s不存在！", regPath)
            return
        value = int(value)
        logger.info("设置注册表%s的值",value_name)
        if rValue == None or rValue is None:
            SetValueEx(newKey, value_name, 0, REG_DWORD, value)
        if rValue != value:
            SetValueEx(newKey, value_name, 0, REG_DWORD, value)

    def loadRegUI(self):
        global GOPTIMIZE
        logger.info("检查注册表信息")
        cpuTypeList = ['1620', '2609', '2620', '2630', '2650', 'other']
        self.checkRegistryInfo("SOFTWARE\\Dayang\\SoftCodec", cpuTypeList)
        self.loadOptkeyUI()

    def installFile(self, mPath, isMIS):
        logger.info("开始安装软件...")
        if mPath and os.path.exists(mPath):
            tPath = os.listdir(mPath)
            logger.info("判断给定路径下第一个元素是否为文件?")
            mpath = str(mPath) + "\\" + str(tPath[0])
            if os.path.isfile(mpath):
                try:
                    if isMIS:
                        handle = os.system(mpath)
                    else:
                        handle = win32process.CreateProcess(mpath, '', None, None, 0, win32process.CREATE_NO_WINDOW,
                                                            None, None, win32process.STARTUPINFO())
                except Exception as e:
                    print(e.message)
            else:
                logger.info("给定路径下第一个元素是文件夹")
        else:
            logger.info("%s,目录不存在！", mPath)
        # 手动安装完成后要刷新页面
        if (isMIS == False):
            win32event.WaitForSingleObject(handle[0], -1)
        self.loadSoftUI()
        self.updateUI()

    def isFileInstalled(self, path, name, obj, insObj, key_name):
        index = path.find(":")
        isInstall = None
        if index != -1:
            isInstall = self.isFileExisted(path, name, obj, insObj)
        else:
            isInstall = self.isSoftwareInstalled(path, name, obj, insObj, key_name)
        return isInstall

    def isFileExisted(self, fPath, fileName, obj, insObj):
        logger.info("检查文件是否存在")
        isExisted = os.path.exists(fPath + "\\" + fileName)
        if not isExisted:
            logger.info("文件不存在")
            self.setErrorText(obj, r"未安装")
            insObj.setEnabled(True)
            self.setSingleUse(fileName,False)
            return False
        else:
            logger.info("文件存在")
            insObj.setEnabled(False)
            self.setText(obj, r"已安装")
            self.setSingleUse(fileName,True)
            return True

    def isSoftwareInstalled(self, rPath, softName, obj, insObj, key_name):
        logger.info("检查软件是否安装")
        value = self.QueryReg(rPath, softName, key_name)
        if value == True:
            logger.info("软件已经安装")
            insObj.setEnabled(False)
            self.setText(obj, r"已安装")
            self.setSingleUse(softName,True)
            return True
        else:
            logger.info("软件未安装")
            # self.isNotUsable()
            self.setErrorText(obj, r"未安装")
            self.setSingleUse(softName,False)
            return False

    def loadSoftUI(self):
        key_name = "DisplayName"
        logger.info("判断DirectX 10是否安装")
        isInstalled = self.isFileInstalled("C:\\WINDOWS\\system32", "D3DX11_41.dll", self.label_software_Directx10_result, self.pushButton_Directx10, key_name)
        if (isInstalled != True):
            self.pushButton_Directx10.clicked.connect(lambda: self.installFile(".\\Tools\\DirectX10", False))

        logger.info("判断DirectX 11是否安装")
        isInstalled = self.isFileInstalled("C:\\WINDOWS\\system32", "D3DX11_43.dll", \
                                           self.label_software_Directx11_result, self.pushButton_Directx11, key_name)
        if (isInstalled != True):
            self.pushButton_Directx11.clicked.connect(lambda: self.installFile(".\\Tools\\DirectX11",False))

        logger.info("判断VC2010 x64是否安装")
        isInstalled = self.isFileInstalled("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall", \
                                           "Microsoft Visual C++ 2010  x64 Redistributable", \
                                           self.label_software_vc2010x64_result, self.pushButton_vc2010x64, key_name)
        if (isInstalled != True):
            self.pushButton_vc2010x64.clicked.connect(lambda: self.installFile(".\\Tools\\VC2010X64", False))

        logger.info("判断VC2010 x86是否安装")
        isInstalled = self.isFileInstalled("SOFTWARE\\Wow6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall", \
                                           "Microsoft Visual C++ 2010  x86 Redistributable", \
                                           self.label_software_vc2010x86_result, self.pushButton_vc2010x86, key_name)
        if (isInstalled != True):
            self.pushButton_vc2010x86.clicked.connect(lambda: self.installFile(".\\Tools\\VC2010X86", False))

        logger.info("判断Setup1是否安装")
        isInstalled = self.isFileInstalled("SOFTWARE\\Wow6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall", \
            "Setup1", self.label_software_setup1_result, self.pushButton_setup1, key_name)
        if (isInstalled != True):
            self.pushButton_setup1.clicked.connect(lambda: self.installFile(".\\Tools\\Setup1", True))

        logger.info("判断QuickTime是否安装")
        key_name = "ApplicationName"
        isInstalled = self.isFileInstalled("SOFTWARE\\Clients\\Media\\QuickTime", "QuickTime", \
                                           self.label_software_Quciktime_result, self.pushButton_Quciktime, key_name)
        if (isInstalled != True):
            self.pushButton_Quciktime.clicked.connect(lambda: self.installFile(".\\Tools\\QuickTime", False))

    def getUsable(self):
        global singleUse,GUSEABLE
        GUSEABLE = True
        for key in singleUse:
            temp = self.getSingleUse(key)
            logger.info("检查项 %s 是否可用：%s", key, str(temp))
            GUSEABLE = GUSEABLE and temp
            logger.info("显示结果区域中显示是否可用：%s", GUSEABLE)

    def updateUI(self):
        logger.info("更新显示结果区域的信息")
        global GUSEABLE
        self.getUsable()
        if GUSEABLE:
            self.setText(self.label_usable_result, r"可以使用")
        else:
            self.setErrorText(self.label_usable_result, r"无法使用，请关注标红部分！")
        strUsable = self.label_usable_result.text()
        strOptimize = self.label_optimize_result.text()
        logger.info("设置一键优化按钮状态")
        if (strUsable == r"可以使用") and (strOptimize == r"无需优化"):
            self.pushButton.setEnabled(False)
        if strUsable.find(r"无法使用") == 0 or strOptimize == r"需要优化":
            self.pushButton.setEnabled(True)
        if strOptimize == r"优化完毕！":
            self.pushButton.setEnabled(False)

    def optimizeOneKey(self, rPath):
        global regList
        logger.info("一键优化按钮处理函数")
        for rKey in regList:
            self.updateReg(rKey, regList[rKey],rPath)
        self.setText(self.label_optimize_result, r"优化完毕！")
        self.updateUI()

    def loadOptkeyUI(self):
        logger.info("点击一键优化按钮")
        self.pushButton.clicked.connect(lambda: self.optimizeOneKey("SOFTWARE\\Dayang\\SoftCodec"))

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
    window.updateUI()
    window.show()
    logger.info("QApplication exit...")
    sys.exit(app.exec_())
