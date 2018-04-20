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
from PyQt5 import QtGui
from selenium import webdriver
import win32process
import win32event
import win32api
import RB6Version
import ctypes

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

global boardInfo
boardInfo = {}

class MyApp(QtWidgets.QMainWindow):
    def __init__(self):
        logger.info("QMainWindow初始化")
        QtWidgets.QMainWindow.__init__(self)

class TabCompute(MyApp,Ui_MainWindow):
    palette = QtGui.QPalette()
    global regList
    def __init__(self):
        logger.info("进入TabCompute类的初始化函数")
        super(TabCompute,self).__init__()
        logger.info("Ui_MainWindow初始化")
        Ui_MainWindow.__init__(self)
        logger.info("setupUi...")
        self.setupUi(self)

    def setErrorText(self, controlName, text, errMsg=""):
        logger.info("设置标红的文本信息")
        self.palette.setColor(QtGui.QPalette.WindowText, QtCore.Qt.red)
        controlName.setPalette(self.palette)
        controlName.setText(text + errMsg)

    def setText(self, controlName, text):
        logger.info("设置文本信息")
        self.palette.setColor(QtGui.QPalette.WindowText, QtCore.Qt.black)
        controlName.setPalette(self.palette)
        controlName.setText(text)

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
            logger.error("无需优化-->需要")
            self.setErrorText(self.label_optimize_result, r"需要优化")

    def getUsable(self):
        global singleUse,GUSEABLE
        GUSEABLE = True
        for key in singleUse:
            temp = self.getSingleUse(key)
            logger.info("检查项 %s 是否可用：%s", key, str(temp))
            GUSEABLE = GUSEABLE and temp
            logger.info("显示结果区域中显示是否可用：%s", GUSEABLE)

    def QueryReg(self, rPath, value_name, key_name):
        logger.info("查询注册表中 %s的数值", value_name)
        keylist = []
        try:
            key = OpenKey(HKEY_LOCAL_MACHINE, rPath, 0, KEY_READ)
            subKey = QueryInfoKey(key)[0]
        except:
            logger.error("要查询的注册表路径 %s 不存在", rPath)
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
        logger.error("未查到指定路径下的 %s 的值", key_name)
        return None

    def readFromReg(self, rPath, value_name):
        try:
            key = OpenKey(HKEY_LOCAL_MACHINE, rPath, 0, KEY_READ)
            value, type = QueryValueEx(key, value_name)
            logger.info("读注册表中%s的数值", value)
            return value
        except:
            logger.error("注册表的路径或者项不存在")
            return None

class HardwareInfo(TabCompute):
    div_gb_factor = (1024.0 ** 3)
    s = wmi.WMI()
    cpus = s.Win32_Processor()
    imem = s.Win32_ComputerSystem()
    igpu = s.Win32_VideoController()
    def __init__(self):
        logger.info("进入HardwareInfo类的初始化函数")
        super().__init__()

    def getCPUSpeed(self, cpu):
        logger.info("获取CPU主频信息")
        cpuType = str(cpu.Name)
        cpuType = cpuType.split('@', 1)
        result = float(cpuType[1].split('GHz', 1)[0])
        return result

    def getCPUType(self, cpu):
        logger.info("获取CPU型号信息")
        cType = str(cpu.Name)
        cType = cType.split('@', 1)
        return cType

    def getCPUNum(self, cpu, cpuNum):
        logger.info("获取CPU的核数信息")
        coreNum = str(cpu.NumberOfCores * cpuNum)
        return coreNum

    def handleCpuInfo(self, cpus):
        global strCpuType
        for cpu in cpus:
            cpuType = self.getCPUType(cpu)
            logger.info("处理多个CPU的情况！")
            if strCpuType == "初始状态":
                logger.info("初始状态下，设置CPU型号为: %s", cpuType[0])
                text = cpuType[0]
                super().setSingleUse("CPU型号", True)
                #MyApp.setSingleUse("CPU型号", True)
                super().setText(self.label_hardware_CPUmodel_result, text)
                strCpuType = cpuType
            elif strCpuType != cpuType:
                logger.warning("存在CPU型号不一致的情况，此时设置CPU的类型为其他")
                text = u"CPU型号不一致！"
                super().setSingleUse("CPU型号", False)
                super().setErrorText(self.label_hardware_CPUmodel_result, text)
                strCpuType = "other"
            else:
                logger.warning("存在多个相同型号的CPU")
                super().setSingleUse("CPU型号", True)
                strCpuType = cpuType
            #备注：若存在多块CPU的情况，若型号一致和一块的情况类似，若存在多块类型不一致的以最后检测到的那块信息显示
            speed = self.getCPUSpeed(cpu)
            logger.info("判断CPU主频是否满足需求?")
            text = cpuType[1].strip()
            errMsg = u", 请更换主频更高的CPU!"
            if speed < 2.0:
                logger.error("CPU主频小于2.0GHz！")
                # self.isNotUsable()
                super().setSingleUse("Speed", False)
                super().setErrorText(self.label_hardware_CPUspeed_result, text, errMsg)
            else:
                logger.info("CPU主频大于等于2.0GHZ!")
                super().setSingleUse("Speed", True)
                super().setText(self.label_hardware_CPUspeed_result, text)
        coreNum = self.getCPUNum(cpu, len(cpus))
        logger.info("判断CPU核数是否满足需求?")
        if int(coreNum) < 4:
            logger.error("CPU的核数小于4，要设置红色字提示当前计算机的硬件环境信息不可以使用非编软件！")
            super().setSingleUse("Core", False)
            super().setErrorText(self.label_hardware_CPUnum_result, coreNum, "核数小于4")
        else:
            logger.info("CPU核数满足使用非编软件的最低硬件环境需求！")
            super().setSingleUse("Core",True)
            super().setText(self.label_hardware_CPUnum_result, coreNum)

    def handleMemoryInfo(self,imem):
        for mem in imem:
            logger.info("获取内存信息")
            memCap = round(int(mem.TotalPhysicalMemory) / HardwareInfo.div_gb_factor)
            logger.info("判断内存是否满足要求")
            text = str(memCap) + 'GB'
            if memCap < 8:
                logger.error("内存小于8G")
                super().setSingleUse("Memory",False)
                super().setErrorText(self.label_hardware_memory_result, text, "内存小于8G")
            else:
                logger.info("内存大于等于8G")
                super().setSingleUse("Memory",True)
                super().setText(self.label_hardware_memory_result, text)

    def handleGPUInfo(self,igpu):
        logger.info("获取显卡信息")
        for gpu in igpu:
            gpuMemCap = gpu.CurrentNumberOfColors
            logger.info("判断显卡是否是独立显卡")
            if gpuMemCap is None:
                logger.error("目前获取到的显卡的显存为零")
                super().setSingleUse("GPUType",False)
                super().setErrorText(self.label_hardware_GPUmodel_result, "请安装独立显卡！")
                continue
            else:
                logger.info("判断显卡显存是否满足要求")
                gpuMemCap = int(gpuMemCap) / HardwareInfo.div_gb_factor
                text = str(gpuMemCap) + 'GB'
                if int(gpuMemCap) < 1:
                    logger.error("显存小于1GB")
                    # self.isNotUsable()
                    super().setSingleUse("GPU",False)
                    super().setErrorText(self.label_hardware_GPUmemory_result, text, u"显存小于1G")
                else:
                    logger.info("显存大于等于1GB")
                    super().setSingleUse("GPU",True)
                    super().setText(self.label_hardware_GPUmemory_result, text)
                gpuInfo = gpu.Name
                logger.info("判断显卡型号信息是否满足要求")
                if (gpuInfo is None):
                    logger.error("显卡型号信息获取不正常")
                    super().setSingleUse("GPUType",False)
                    super().setErrorText(self.label_hardware_GPUmodel_result, u"未获取显卡型号信息")
                else:
                    logger.info("显卡型号信息获取正常")
                    super().setSingleUse("GPUType",True)
                    super().setText(self.label_hardware_GPUmodel_result, gpuInfo)

    def hfillup(self):
        logger.info("开始填充硬件信息")
        self.handleCpuInfo(HardwareInfo.cpus)
        self.handleMemoryInfo(HardwareInfo.imem)
        self.handleGPUInfo(HardwareInfo.igpu)

class SoftwareInfo(TabCompute):
    def __init__(self):
        logger.info("进入SoftwareInfo类的初始化函数")
        super().__init__()

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
            logger.error("%s,目录不存在！", mPath)
        # 手动安装完成后要刷新页面
        if (isMIS == False):
            win32event.WaitForSingleObject(handle[0], -1)
        self.fillup()
        super().updateUI()

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
            logger.error("文件不存在")
            super().setErrorText(obj, r"未安装")
            insObj.setEnabled(True)
            super().setSingleUse(fileName,False)
            return False
        else:
            logger.info("文件存在")
            insObj.setEnabled(False)
            super().setText(obj, r"已安装")
            super().setSingleUse(fileName,True)
            return True

    def isSoftwareInstalled(self, rPath, softName, obj, insObj, key_name):
        logger.info("检查软件是否安装")
        value = super().QueryReg(rPath, softName, key_name)
        if value == True:
            logger.info("软件已经安装")
            insObj.setEnabled(False)
            super().setText(obj, r"已安装")
            super().setSingleUse(softName,True)
            return True
        else:
            logger.error("软件未安装")
            # self.isNotUsable()
            super().setErrorText(obj, r"未安装")
            super().setSingleUse(softName,False)
            return False

    def isBoardCardInstalled(self,fPath,driverName,obj,boardObj):
        #logger.info("检查板卡驱动是否存在")
        isExisted = os.path.exists(fPath + "\\" + driverName)
        if not isExisted:
            logger.warning("文件不存在")
            super().setText(obj, r"未知")
            self.label_software_IOCard_result_1.setText(r"无驱动信息")
            boardObj.setEnabled(True)
            return False
        else:
            logger.info("文件存在")
            #备注：必须先获取板卡的型号才能知道驱动版本信息
            obj.setText(self.getRB6Type())
            self.label_software_IOCard_result_1.setText(self.getRB6Version())
            boardObj.setEnabled(False)
            return True

    def getRB6Version(self):
        logger.info("获取RB6板卡信息！")
        version = RB6Version.getVersion()
        if version is None:
            logger.error("请先获取驱动型号的信息！")
            return None
        return version

    def getRB6Type(self):
        logger.info("获取RB6板卡型号信息！")
        type = RB6Version.getBoardType()
        return type

    def sfillup(self):
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

        logger.info("判断板卡驱动是否安装")
        isInstalled = self.isBoardCardInstalled("C:\\Program Files\\Beijing Qualitune Technology Inc\\RedBridge6 Runtime\\driver\\amd64","RB6_SDK.dll",\
                                           self.label_hardware_IOCard_result_1,self.pushButton_IOCard_1)
        if (isInstalled != True):
            self.pushButton_IOCard_1.clicked.connect(lambda: self.installFile(".\\Tools\\RB6Driver", False))

class RegistryInfo(TabCompute):
    cpuTypeList = ['1620', '2609', '2620', '2630', '2650', 'other']
    def __init__(self):
        logger.info("进入RegistryInfo类的初始化函数")
        super().__init__()

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

    def checkRegistryInfo(self, rPath, cpuTypeList):
        global regList
        logger.info("检查本机的cpu型号是否是已知型号?")
        for cpuType in cpuTypeList:
            # Judge the cputype of the target machine is in the cpuTypeList, if yes then save in the global variable
            if strCpuType[0].find(cpuType) != -1:
                logger.info("模板文件中CPU类型为 %s 的预制注册表值读取", cpuType)
                regList = self.getListData(cpuType)
                break
            else:
                continue
        for rKey in regList:
            regData = super().readFromReg(rPath, rKey)
            if regData == None or regData is None:
                logger.error("注册表路径 %s值未读取到 %s 的数据", rPath, rKey)
                super().setOptimize(False)
            elif str(regData) != str(regList[rKey]):
                logger.error("注册表值读取的数据为：%s，预制模板中读取的数据为：%s，两者不一致",str(regData),str(regList[rKey]))
                super().setOptimize(False)
            else:
                logger.info("注册表值读取的数据为：%s，预制模板中读取的数据为：%s，两者一致", str(regData), str(regList[rKey]))
                super().setOptimize(True)

    def updateReg(self, value_name, value, regPath):
        logger.info("更新%s路径下%s的值",regPath,value_name)
        # Read from registry to see if the value_name is existed
        rValue = super().readFromReg(regPath, value_name)
        try:
            newKey = CreateKeyEx(HKEY_LOCAL_MACHINE, regPath, 0, KEY_WRITE)
        except:
            logger.error("指定路径%s不存在！", regPath)
            return
        value = int(value)
        logger.info("设置注册表%s的值",value_name)
        if rValue == None or rValue is None:
            SetValueEx(newKey, value_name, 0, REG_DWORD, value)
        if rValue != value:
            SetValueEx(newKey, value_name, 0, REG_DWORD, value)

    def loadOptkeyUI(self):
        logger.info("点击一键优化按钮")
        self.pushButton.clicked.connect(lambda: self.optimizeOneKey("SOFTWARE\\Dayang\\SoftCodec"))

    def rfillup(self):
        global GOPTIMIZE
        logger.info("检查本机的cpu型号是否注册表信息")
        self.checkRegistryInfo("SOFTWARE\\Dayang\\SoftCodec", RegistryInfo.cpuTypeList)
        self.loadOptkeyUI()

    def optimizeOneKey(self, rPath):
        global regList
        logger.info("一键优化按钮处理函数")
        for rKey in regList:
            self.updateReg(rKey, regList[rKey],rPath)
        super().setText(self.label_optimize_result, r"优化完毕！")
        super().updateUI()

class CommonInfo(HardwareInfo,SoftwareInfo,RegistryInfo):
    def __init__(self):
        logger.info("进入CommonInfo类的初始化函数")
        super().__init__()

    def updateUI(self):
        logger.info("更新显示结果区域的信息")
        global GUSEABLE
        self.getUsable()
        if GUSEABLE:
            super().setText(self.label_usable_result, r"可以使用")
        else:
            super().setErrorText(self.label_usable_result, r"无法使用，请关注标红部分！")
        strUsable = self.label_usable_result.text()
        strOptimize = self.label_optimize_result.text()
        logger.info("设置一键优化按钮状态")
        if (strUsable == r"可以使用") and (strOptimize == r"无需优化"):
            self.pushButton.setEnabled(False)
        if strUsable.find(r"无法使用") == 0 or strOptimize == r"需要优化":
            self.pushButton.setEnabled(True)
        if strOptimize == r"优化完毕！":
            self.pushButton.setEnabled(False)

class AnDParameter():
    # SeekEX模式,0：单场模式（默认）1：双场模式
    _bSeekExMode = 0
    # D3D Edit Advance,在FxD3DXCommon.ini中,0：不启动(默认),1：启动
    _bD3DEditAdvance = 0
    # 同时解码器数量,3：同时开启3个解码器（默认）
    _nMaxDecodeThreadNum = 3
    # 视频解码器数量,30：最多打开30个视频解码器（默认）
    _nMaxVideoCodecNum = 30
    # 音频解码器数量,30：最多打开30个音频解码器（默认）
    _nMaxAudioCodecNum = 30
    # 解码预读,_bDecodePreread=0：解码预读为0（其他情况都为1）,_bDecodePreread=1且_nDecodePreread=2：解码预读为2（默认）
    _bDecodePreread = 1
    _nDecodePreread = 2
    # 回显模式	,0：偶场先,1：奇场先（默认）,2：同时,3：帧,4：差值
    _nViewMode = 1
    # 视音频一体解码优化,0：不启动视音频一体优化,1：启动视音频一体文件优化（默认）
    _bUseAudioCacheForAAFile = 1
    _bUseAudioCacheForVAFile = 1
    # Adcance DVE Edit,0：不启动Advancedve模式（默认）,1：启动Advancedve模式
    _bIsFullSize = 0
    # 系统量化比特,0：8BIT量化(默认),1：10bit量化
    _bEdit16 = 0
    # YUV模式，_bYUVMode=0：不启动YUV模式，_bYUVMode=1且_xProtocol=0、_nYUVxModeHD=0：601模式，_bYUVMode=1且_xProtocol=1、_nYUVxModeHD=1：709模式（默认），_bYUVMode=1且_xProtocol=2、_nYUVxModeHD=2：2020模式
    _bYUVMode = 1
    _xProtocol = 1
    _nYUVxModeHD = 1
    # 下变换模式，0：最快模式，1：一般模式，2：优质模式（默认）
    _nScaleMode = 2
    # 下变换质量，0：速度最快，1：1级,2：2级,3：3级,4：4级（默认）,5：5级,6：6级,7：质量最好
    _nScaleQuality = 4
    # 板卡第二路输出
    # _bIsCardDualPlay=0：不输出
    # _bIsCardDualPlay=1且_bIsCardSameSize=1：与主输出一致（默认）
    # _bIsCardDualPlay=1且_bIsCardSameSize=0且_nDualPlayScaleMode=0：变形上/下变换
    # _bIsCardDualPlay=1且_bIsCardSameSize=0且_nDualPlayScaleMode=1：切边上/下变换
    # _bIsCardDualPlay=1且_bIsCardSameSize=0且_nDualPlayScaleMode=2：信箱上/下变换
    _bIsCardDualPlay = 1
    _bIsCardSameSize = 1
    _nDualPlayScaleMode = 1

class TabDayang(MyApp):
    editsettings = []
    d3dsettings = []
    def __init__(self, devInfo = TabCompute, config = AnDParameter):
        logger.info("进入TabDayang类的初始化函数")
        self.__devInfo = devInfo
        self._config = config
        super(TabDayang,self).__init__()

    def getDayangVersion(self):
        logger.info("获取非编软件的版本信息")
        logger.info("判断注册表中是否存在版本信息")
        iversion = self.__devInfo.readFromReg("SOFTWARE\\Dayang", "SysVersionNo")
        if iversion is None:
            self.__devInfo.setErrorText(self.__devInfo.label_SysVersionNo_result,u"版本信息读取失败")
        else:
            self.__devInfo.setText(self.__devInfo.label_SysVersionNo_result, str(iversion))

    def readFile(self, name):
        readconfig = []
        filePath = self.getDayangPath("SOFTWARE\\Dayang","Path",name)
        logger.info("读取配置文件中的内容")
        file = open(filePath, 'r')  # 返回一个文件对象
        readconfig = file.readlines()  # 调用文件的 readline()方法
        file.close()
        return readconfig

    def getDayangPath(self,regpath,keyname,name):
        logger.info("获取软件的配置文件路径")
        installPath = self.__devInfo.readFromReg(regpath,keyname)
        if installPath is not None:
            installPath += "\\bin\\"
            installPath += name
            logger.info("配置文件的路径为：%s", installPath)
            return installPath
        logger.error("没有找到软件的配置文件路径")

    def initAnDParameter(self):
        logger.info("初始化AnDParameter类")
        d3dfile = self.readFile("FxD3DXCommon.ini")
        editfile = self.readFile("FxEditCommon.ini")
        d3dsetting = "|".join(d3dfile)
        editsetting = "|".join(editfile)
        if self.editsettings is not None:
            del self.editsettings[:]
        self.editsettings = editfile
        # SeekExMode模式
        value = self.handleStr(editsetting, "|_bSeekExMode=")
        self._config._bSeekExMode = value
        #回显模式
        value = self.handleStr(editsetting, "|_nViewMode=")
        self._config._nViewMode = value
        #Adcance DVE Edit
        value = self.handleStr(editsetting,"|_bIsFullSize=")
        self._config._bIsFullSize = value
        #系统量化比特
        value = self.handleStr(editsetting,"|_bEdit16=")
        self._config._bEdit16 = value
        #下变换模式
        value = self.handleStr(editsetting,"|_nScaleMode=")
        self._config._nScaleMode = value
        #下变换质量
        value = self.handleStr(editsetting,"|_nScaleQuality=")
        self._config._nScaleQuality = value
        #视音频一体解码优化
        value = self.handleStr(editsetting,"|_bUseAudioCacheForAAFile=")
        self._config._bUseAudioCacheForAAFile = value
        value = self.handleStr(editsetting,"|_bUseAudioCacheForVAFile=")
        self._config._bUseAudioCacheForVAFile = value
        #YUV模式
        value = self.handleStr(editsetting,"|_bYUVMode=")
        self._config._bYUVMode = value
        value = self.handleStr(editsetting,"|_xProtocol=")
        self._config._xProtocol = value
        value = self.handleStr(editsetting,"|_nYUVxModeHD=")
        self._config._nYUVxModeHD = value
        #板卡第二路输出
        value = self.handleStr(editsetting,"|_bIsCardDualPlay=")
        self._config._bIsCardDualPlay = value
        value = self.handleStr(editsetting, "|_bIsCardSameSize=")
        self._config._bIsCardSameSize = value
        value = self.handleStr(editsetting, "|_nDualPlayScaleMode=")
        self._config._nDualPlayScaleMode = value
        # 同时解码器数量
        value = self.handleStr(editsetting,"|_nMaxDecodeThreadNum=")
        self._config._nMaxDecodeThreadNum = value
        # 视频解码器数量
        value = self.handleStr(editsetting,"|_nMaxVideoCodecNum=")
        self._config._nMaxVideoCodecNum = value
        # 音频解码器数量
        value = self.handleStr(editsetting,"|_nMaxAudioCodecNum=")
        self._config._nMaxAudioCodecNum = value
        # 解码预读
        value = self.handleStr(editsetting,"|_bDecodePreread=")
        self._config._bDecodePreread = value
        value = self.handleStr(editsetting,"|_nDecodePreread=")
        self._config._nDecodePreread = value
        # D3D Edit Advance
        if self.d3dsettings is not None:
            del self.d3dsettings[:]
        self.d3dsettings = d3dfile
        value = self.handleStr(d3dsetting,"|_bD3DEditAdvance=")
        self._config._bD3DEditAdvance = value

    def setParameter(self):
        logger.info("SeekEx模式的设置")
        value = self._config._bSeekExMode
        if value == 0 or value == 1:
            self.__devInfo.comboBox_bSeekExMode.setCurrentIndex(value)
        else:
            index = self.__devInfo.comboBox_bSeekExMode.count()
            self.__devInfo.comboBox_bSeekExMode.setCurrentIndex(index-1)
            logger.error("SeekEx模式的参数获取错误,获取到的值为：%d",value)
        self.__devInfo.comboBox_bSeekExMode.currentIndexChanged.connect(lambda: self.comboboxSlot("FxEditCommon.ini","_bSeekExMode=",\
                                                                                                  self.__devInfo.comboBox_bSeekExMode,\
                                                                                                  self._config._bSeekExMode,1))
        logger.info("回显模式的设置")
        value = self._config._nViewMode
        if value == 0 or value == 1:
            self.__devInfo.comboBox_nViewMode.setCurrentIndex(value)
        else:
            index = self.__devInfo.comboBox_nViewMode.count()
            self.__devInfo.comboBox_nViewMode.setCurrentIndex(index-1)
            logger.error("回显模式的参数获取错误,获取到的值为：%d",value)
        self.__devInfo.comboBox_nViewMode.currentIndexChanged.connect(lambda:self.comboboxSlot("FxEditCommon.ini","_nViewMode=",\
                                                                                               self.__devInfo.comboBox_nViewMode,\
                                                                                               self._config._nViewMode,1))
        logger.info("Adcance DVE Edit的设置")
        value = self._config._bIsFullSize
        if value == 0 or value == 1:
            self.__devInfo.comboBox_bIsFullSize.setCurrentIndex(value)
        else:
            index = self.__devInfo.comboBox_bIsFullSize.count()
            self.__devInfo.comboBox_bIsFullSize.setCurrentIndex(index-1)
            logger.error("Adcance DVE Edit的参数获取错误,获取到的值为：%d",value)
        self.__devInfo.comboBox_bIsFullSize.currentIndexChanged.connect(lambda: self.comboboxSlot("FxEditCommon.ini", "_bIsFullSize=", \
                                                                                                  self.__devInfo.comboBox__bIsFullSize, \
                                                                                                  self._config.__bIsFullSize, 1))
        logger.info("系统量化比特的设置")
        value = self._config._bEdit16
        if value == 0 or value == 1:
            self.__devInfo.comboBox_bEdit16.setCurrentIndex(value)
        else:
            index = self.__devInfo.comboBox_bEdit16.count()
            self.__devInfo.comboBox_bEdit16.setCurrentIndex(index-1)
            logger.error("系统量化比特的参数获取错误,获取到的值为：%d",value)
        self.__devInfo.comboBox_bEdit16.currentIndexChanged.connect(lambda: self.comboboxSlot("FxEditCommon.ini", "_bEdit16=", \
                                                                                              self.__devInfo.comboBox_bEdit16, \
                                                                                              self._config._bEdit16, 1))
        logger.info("下变换模式的设置")
        value = self._config._nScaleMode
        if value == 0 or value == 1:
            self.__devInfo.comboBox_nScaleMode.setCurrentIndex(value)
        else:
            index = self.__devInfo.comboBox_nScaleMode.count()
            self.__devInfo.comboBox_nScaleMode.setCurrentIndex(index-1)
            logger.error("下变换模式的参数获取错误,获取到的值为：%d",value)
        self.__devInfo.comboBox_nScaleMode.currentIndexChanged.connect(lambda: self.comboboxSlot("FxEditCommon.ini", "_nScaleMode=", \
                                                                                                 self.__devInfo.comboBox_nScaleMode, \
                                                                                                 self._config._nScaleMode, 1))
        logger.info("下变换质量的设置")
        value = self._config._nScaleQuality
        if value == 0 or value == 1:
            self.__devInfo.comboBox_nScaleQuality.setCurrentIndex(value)
        else:
            index = self.__devInfo.comboBox_nScaleQuality.count()
            self.__devInfo.comboBox_nScaleQuality.setCurrentIndex(index-1)
            logger.error("下变换质量的参数获取错误,获取到的值为：%d",value)
        self.__devInfo.comboBox_nScaleQuality.currentIndexChanged.connect(lambda: self.comboboxSlot("FxEditCommon.ini", "_nScaleQuality=", \
                                                                                                    self.__devInfo.comboBox_nScaleQuality, \
                                                                                                    self._config._nScaleQuality, 1))
        logger.info("YUV模式的设置")
        value = self._config._bYUVMode
        if value == 0:
            self.__devInfo.comboBox_bYUVMode.setCurrentIndex(value)
        elif value == 1:
            value2 = self._config._xProtocol
            value3 = self._config._nYUVxModeHD
            if value2 == 0 and value3 == 0:
                self.__devInfo.comboBox_bYUVMode.setCurrentIndex(value)
            if value2 == 1 and value3 == 1:
                self.__devInfo.comboBox_bYUVMode.setCurrentIndex(value+1)
            if value2 == 2 and value3 == 2:
                self.__devInfo.comboBox_bYUVMode.setCurrentIndex(value+2)
            else:
                index = self.__devInfo.comboBox_bYUVMode.count()
                self.__devInfo.comboBox_bYUVMode.setCurrentIndex(index - 1)
                logger.error("YUN模式的参数获取失败，获取到的值为：%d,%d", value2,value3)
        else:
            index = self.__devInfo.comboBox_bYUVMode.count()
            self.__devInfo.comboBox_bYUVMode.setCurrentIndex(index - 1)
            logger.error("YUN模式的参数获取失败，获取到的值为：%d",value)
        self.__devInfo.comboBox_bYUVMode.currentIndexChanged.connect(lambda: self.YUNModeSlot("FxEditCommon.ini", self.__devInfo.comboBox_bYUVMode, self._config))
        logger.info("板卡第二路输出的设置")
        value = self._config._bIsCardDualPlay
        if value == 0:
            self.__devInfo.comboBox_bIsCardDualPlay.setCurrentIndex(value)
        elif value == 1:
            value2 = self._config._bIsCardSameSize
            if value2 == 1:
                self.__devInfo.comboBox_bIsCardDualPlay.setCurrentIndex(value2)
            elif value2 ==0:
                value3 = self._config._nDualPlayScaleMode
                if value3 == 0 or value3 == 1 or value3 == 2:
                    self.__devInfo.comboBox_bIsCardDualPlay.setCurrentIndex(value3+2)
                else:
                    index = self.__devInfo.comboBox_bIsCardDualPlay.count()
                    self.__devInfo.comboBox_bIsCardDualPlay.setCurrentIndex(index - 1)
                    logger.error("板卡第二路输出的参数获取失败，获取到的值为：%d", value3)
            else:
                index = self.__devInfo.comboBox_bIsCardDualPlay.count()
                self.__devInfo.comboBox_bIsCardDualPlay.setCurrentIndex(index - 1)
                logger.error("板卡第二路输出的参数获取失败，获取到的值为：%d", value2)
        else:
            index = self.__devInfo.comboBox_bIsCardDualPlay.count()
            self.__devInfo.comboBox_bIsCardDualPlay.setCurrentIndex(index - 1)
            logger.error("板卡第二路输出的参数获取失败，获取到的值为：%d", value)
        self.__devInfo.comboBox_bIsCardDualPlay.currentIndexChanged.connect(lambda: self.CardDualPlaySlot("FxEditCommon.ini", self.__devInfo.comboBox_bIsCardDualPlay, self._config))
        logger.info("视音频一体解码优化的设置")
        value = self._config._bUseAudioCacheForAAFile
        value1 = self._config._bUseAudioCacheForVAFile
        if value == 0 and value1 == 0:
            self.__devInfo.comboBox_bUseAudioCache.setCurrentIndex(value)
        elif value == 1 and value1 == 1:
            self.__devInfo.comboBox_bUseAudioCache.setCurrentIndex(value)
        else:
            index = self.__devInfo.comboBox_bUseAudioCache.count()
            self.__devInfo.comboBox_bUseAudioCache.setCurrentIndex(index - 1)
            logger.error("视音频一体解码优化的参数获取失败，获取到的值为：%d,%d", value,value1)
        self.__devInfo.comboBox_bUseAudioCache.currentIndexChanged.connect(lambda: self.UserAudioCacheSlot("FxEditCommon.ini", self.__devInfo.comboBox_bUseAudioCache, self._config))
        logger.info("同时解码器数量的设置")
        self.__devInfo.spinBox_nMaxDecodeThreadNum.setValue(self._config._nMaxDecodeThreadNum)
        self.__devInfo.spinBox_nMaxDecodeThreadNum.valueChanged.connect(
            lambda:self.setDecodeThreadValue("FxEditCommon.ini", "_nMaxDecodeThreadNum=",
                                             self.__devInfo.spinBox_nMaxDecodeThreadNum, 1))
        logger.info("视频解码器数量的设置")
        self.__devInfo.spinBox_nMaxVideoCodecNum.setValue(self._config._nMaxVideoCodecNum)
        self.__devInfo.spinBox_nMaxVideoCodecNum.valueChanged.connect(
            lambda:self.setVideoCodecValue("FxEditCommon.ini", "_nMaxVideoCodecNum=",
                                           self.__devInfo.spinBox_nMaxVideoCodecNum, 1))
        logger.info("音频解码器数量的设置")
        self.__devInfo.spinBox_nMaxAudioCodecNum.setValue(self._config._nMaxAudioCodecNum)
        self.__devInfo.spinBox_nMaxAudioCodecNum.valueChanged.connect(
            lambda: self.setAudioCodecValue("FxEditCommon.ini", "_nMaxAudioCodecNum=",
                                            self.__devInfo.spinBox_nMaxAudioCodecNum, 1))

        logger.info("解码预读的设置")
        value = self._config._bDecodePreread
        if value == 0:
            self.__devInfo.spinBox_nDecodePreread.setValue(value)
        elif value == 1:
            value1 = self._config._nDecodePreread
            if value1 < 1:
                logger.error("解码预读的参数获取失败，获取到的值为：%d",value1)
                self.__devInfo.spinBox_nDecodePreread.setValue(1)
            elif value1 > 200:
                logger.error("解码预读的参数获取失败，获取到的值为：%d", value1)
                self.__devInfo.spinBox_nDecodePreread.setValue(200)
            else:
                self.__devInfo.spinBox_nDecodePreread.setValue(value1)
        elif value < 0:
            logger.error("解码预读的参数获取失败，获取到的值为：%d",value)
            self.__devInfo.spinBox_nDecodePreread.setValue(0)
        else:
            logger.error("解码预读的参数获取失败，获取到的值为：%d", value)
            self.__devInfo.spinBox_nDecodePreread.setValue(1)
        self.__devInfo.spinBox_nDecodePreread.valueChanged.connect(
            lambda: self.setDecodePrereadValue("FxEditCommon.ini", self.__devInfo.spinBox_nDecodePreread, 1))
        logger.info("D3D Edit Advance的设置")
        value = self._config._bD3DEditAdvance
        if value == 0:
            self.__devInfo.comboBox_bD3DEditAdvance.setCurrentIndex(0)
        elif value == 1:
            self.__devInfo.comboBox_bD3DEditAdvance.setCurrentIndex(1)
        else:
            index = self.__devInfo.comboBox_bD3DEditAdvance.count()
            self.__devInfo.comboBox_bD3DEditAdvance.setCurrentIndex(index - 1)
            logger.error("D3D Edit Advance的参数获取错误，获取到的值为：%d", value)
        self.__devInfo.comboBox_bD3DEditAdvance.currentIndexChanged.connect(lambda: self.comboboxSlot("FxD3DXCommon.ini", "_bD3DEditAdvance=",\
                                                                                                      self.__devInfo.comboBox_bD3DEditAdvance,\
                                                                                                      self._config._bD3DEditAdvance,0))

    def handleStr(self,strName,substr):
        logger.info("开始查找配置文件中存在%s的设置",substr)
        if strName.find(substr) != -1:
            logger.info("查找成功！")
            temp = strName.partition(substr)
            istr = temp[2]
            index = istr.find('|')
            seekExMode = istr[0:index-1]
            value = int(seekExMode)
            return value
            #self._config._bSeekExMode = iseek
        else:
            logger.error("配置文件中没有找到%s的设置",substr)
            return -1

    def YUNModeSlot(self,filename,combObj,confObj):
        index = combObj.currentIndex()
        logger.info("获取YUN模式下拉框中的索引值为:%d", index)
        if index == 0:
            logger.info("将当前选择的YUN模式的索引值写入文件")
            self.writeFile(filename,"_bYUVMode=",0,1)
        else:
            self.writeFile(filename,"_bYUVMode=",1,1)
            if index == 1:
                logger.info("将当前选择的YUN模式的索引值写入文件")
                self.writeFile(filename,"_xProtocol=", 0, 1)
                self.writeFile(filename, "_nYUVxModeHD=", 0, 1)
            if index == 2:
                logger.info("将当前选择的YUN模式的索引值写入文件")
                self.writeFile(filename, "_xProtocol=", 1, 1)
                self.writeFile(filename, "_nYUVxModeHD=", 1, 1)
            if index == 3:
                logger.info("将当前选择的YUN模式的索引值写入文件")
                self.writeFile(filename, "_xProtocol=", 2, 1)
                self.writeFile(filename, "_nYUVxModeHD=", 2, 1)
            else:
                logger.error("选择了未知项，无法写入文件")

    def CardDualPlaySlot(self,filename,combObj,confObj):
        index = combObj.currentIndex()
        logger.info("获取板卡第二路输出下拉框中的索引值为:%d", index)
        if index == 0:
            logger.info("将当前选择的板卡第二路输出的索引值写入文件")
            self.writeFile(filename, "_bIsCardDualPlay=", 0, 1)
        else:
            self.writeFile(filename, "_bIsCardDualPlay=", 1, 1)
            if index == 1:
                logger.info("将当前选择的板卡第二路输出的索引值写入文件")
                self.writeFile(filename,"_bIsCardSameSize=", 1, 1)
            else:
                self.writeFile(filename,"_bIsCardSameSize=", 0, 1)
                if index == 2:
                    logger.info("将当前选择的板卡第二路输出的索引值写入文件")
                    self.writeFile(filename, "_nDualPlayScaleMode=", 0, 1)
                if index == 3:
                    logger.info("将当前选择的板卡第二路输出的索引值写入文件")
                    self.writeFile(filename, "_nDualPlayScaleMode=", 1, 1)
                if index == 4:
                    logger.info("将当前选择的板卡第二路输出的索引值写入文件")
                    self.writeFile(filename, "_nDualPlayScaleMode=", 2, 1)
                else:
                    logger.error("选择了未知项，无法写入文件")

    def UserAudioCacheSlot(self,filename,combObj,confObj):
        index = combObj.currentIndex()
        logger.info("获取视音频一体解码优化下拉框中的索引值为:%d", index)
        if index == 0:
            logger.info("将当前选择的视音频一体解码优化的索引值写入文件")
            self.writeFile(filename, "_bUseAudioCacheForAAFile=", 0, 1)
            self.writeFile(filename, "_bUseAudioCacheForVAFile=", 0, 1)
        elif index == 1:
            logger.info("将当前选择的视音频一体解码优化的索引值写入文件")
            self.writeFile(filename, "_bUseAudioCacheForAAFile=", 1, 1)
            self.writeFile(filename, "_bUseAudioCacheForVAFile=", 1, 1)
        else:
            logger.error("选择了未知项，无法写入文件")

    #备注：此函数只处理需求中的内容只有一个文件参数的情况，且确
    #      保函数索引值和文件中写入的值是相等的关系。
    def comboboxSlot(self,filename,item,combObj,confObj,flag):
        index = combObj.currentIndex()
        logger.info("获取当前下拉框中的索引值为:%s",index)
        num = combObj.count()
        if index < num:
            confObj = index
            logger.info("设置AnDParameter类中%s的值为:%d",item,index)
            logger.info("将当前选择的索引值写入文件")
            self.writeFile(filename, item, index,flag)
        else:
            logger.error("选择了未知项，无法写入文件")

    def setDecodeThreadValue(self,name,key,combObj,flag):
        self._config._nMaxDecodeThreadNum = combObj.value()
        self.writeFile(name,key,self._config._nMaxDecodeThreadNum,flag)

    def setVideoCodecValue(self,name,key,combObj, flag):
        self._config._nMaxVideoCodecNum = combObj.value()
        self.writeFile(name,key,self._config._nMaxVideoCodecNum ,flag)

    def setAudioCodecValue(self,name,key,combObj, flag):
        self._config._nMaxAudioCodecNum = combObj.value()
        self.writeFile(name,key,self._config._nMaxAudioCodecNum ,flag)

    def setDecodePrereadValue(self,name, combObj, flag):
        value = combObj.value()
        if value == 0:
            self._config._bDecodePreread  = 0
            self.writeFile(name,"_bDecodePreread=",value,flag)
        else:
            self._config._bDecodePreread = 1
            self._config._nDecodePreread = value
            self.writeFile(name,"_bDecodePreread=",1,flag)
            self.writeFile(name,"_nDecodePreread=",value,flag)

    def writeFile(self,name,key,value,flag):
        #备注：key传入的时候前面不包括"|"
        #flag为1表明是FxEditCommon.ini文件，为0表明是FxD3DXCommon.ini文件
        if flag == 1:
            tempstr = "|".join(self.editsettings)
        if flag == 0:
            tempstr = "|".join(self.d3dsettings)
        item = key
        key = "|" + key
        index = self.handleStr(tempstr,key)
        olditem = item + str(index) + "\n"
        logger.info("将修改的内容为：%s",olditem)
        if flag == 1:
            i = self.editsettings.index(olditem)
            if olditem in self.editsettings:
                newitem = item + str(value) + "\n"
                logger.info("修改后的内容为：%s", newitem)
                self.editsettings[i] = newitem
            else:
                logger.info("配置文件中没有找到%s的内容", olditem)
        if flag == 0:
            i = self.d3dsettings.index(olditem)
            if olditem in self.d3dsettings:
                newitem = item + str(value) + "\n"
                logger.info("修改后的内容为：%s",newitem)
                self.d3dsettings[i] = newitem
            else:
                logger.info("配置文件中没有找到%s的内容",olditem)
        logger.info("修改配置文件的内容")
        filePath = self.getDayangPath("SOFTWARE\\Dayang", "Path", name)
        file = open(filePath, 'w')
        if flag == 1:
            file.writelines(self.editsettings)
        if flag == 0:
            file.writelines(self.d3dsettings)
        file.close()

if __name__ == "__main__":
    logger.info("Enter into main func...")
    regTempPath = os.path.abspath(os.path.dirname(sys.argv[0]))
    dom = parse(regTempPath + '\\registryTemplate.xml')
    logger.info("QApplication...")
    app = QtWidgets.QApplication(sys.argv)
    logger.info("TabCompute页面数据加载")
    icomp = CommonInfo()
    icomp.hfillup()
    icomp.sfillup()
    icomp.rfillup()
    icomp.updateUI()
    logger.info("TabDayang页面数据加载")
    iDayang = TabDayang(icomp)
    logger.info("保存ini文件中的各个选项的数值")
    iDayang.initAnDParameter()
    logger.info("获取非编软件的版本")
    iDayang.getDayangVersion()
    logger.info("展现各个视频参数的数据")
    iDayang.setParameter()
    logger.info("页面显示")
    icomp.show()
    logger.info("QApplication exit...")
    RB6Version.releaseBoard()
    sys.exit(app.exec_())
