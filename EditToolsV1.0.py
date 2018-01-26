#-*- coding: UTF-8 -*-
import sys
import os
import logging
import wmi
from winreg import *
from xml.dom.minidom import parse
from mainwindow import *

#regPath = "SOFTWARE\\Dayang\\SoftCodec"
regTempPath = os.path.abspath(os.path.dirname(sys.argv[0]))
dom = parse(regTempPath+'\\registryTemplate.xml')

#创建一个logger
logger = logging.getLogger('mylogger')
logger.setLevel(logging.DEBUG)
#创建一个handler，用于写入日志文件
fh = logging.FileHandler('Log.log','w')
fh.setLevel(logging.DEBUG)
#定义handler的输出格式
formatter = logging.Formatter('[%(asctime)s][%(thread)d][%(filename)s][line: %(lineno)d][%(levelname)s] ## %(message)s')
fh.setFormatter(formatter)
#给logger添加handler
logger.addHandler(fh)

class MyApp(QtWidgets.QMainWindow, Ui_MainWindow):
    global s
    s = wmi.WMI()
    global div_gb_factor
    # CPU Type in the target machine
    global strCpuType
    # CPU Type list which supported currently
    global cpuTypeList
    #the default is 0(无需优化)
    global flag
    flag =0
    # The registry (key, value) which stored in the xml file
    global regList
    regList = {}
    global palette
    palette = QtGui.QPalette()

    def getCpuInfo(self):
        logger.info("获取CPU信息!")
        cpuNum = 0
        for cpu in s.Win32_Processor():
            cpuType = str(cpu.Name)
            cpuType = cpuType.split('@',1)
            result = float(cpuType[1].split('GHz',1)[0])
            logger.info("判断CPU主频是否满足需求！")
            if result < 2.0:
                logger.info("CPU主频小于2.0GHz！")
                self.isNotUsable()
                #self.updateUI()
                palette.setColor(QtGui.QPalette.WindowText,QtCore.Qt.red)
                self.label_hardware_CPUspeed_result.setPalette(palette)
                self.label_hardware_CPUspeed_result.setText(cpuType[1].strip() + u", 请更换主频更高的CPU!")
            else:
                logger.info("CPU主频大于等于2.0GHZ!")
                self.label_hardware_CPUspeed_result.setText(cpuType[1].strip())
            logger.info("判断多个CPU的型号是否一致")
            if cpuType is not None:
                logger.info("设置CPU型号为: %s",cpuType[0])
                self.label_hardware_CPUmodel_result.setText(cpuType[0])
                self.strCpuType = cpuType
            elif self.strCpuType != cpuType:
                logger.info("存在CPU型号不一致，设置CPU的类型为其他")
                palette.setColor(QtGui.QPalette.WindowText, QtCore.Qt.red)
                self.label_hardware_CPUmodel_result.setPalette(palette)
                self.label_hardware_CPUmodel_result.setText("CPU型号不一致！")
                self.strCpuType = "other"
            else:
                logger.info("存在多个相同型号的CPU")
                self.strCpuType = cpuType
            cpuNum += 1
        logger.info("获取CPU的核数！")
        coreNum = str(cpu.NumberOfCores * cpuNum)
        if int(coreNum) < 4:
            logger.info("CPU的核数小于4，要设置红色字提示当前计算机的硬件环境信息不可以使用非编软件！")
            self.isNotUsable()
            #self.updateUI()
            palette.setColor(QtGui.QPalette.WindowText, QtCore.Qt.red)
            self.label_hardware_CPUnum_result.setPalette(palette)
            self.label_hardware_CPUnum_result.setText(coreNum)
        logger.info("CPU核数满足使用非编软件的最低硬件环境需求！")
        self.label_hardware_CPUnum_result.setText(coreNum)
        #self.strCpuType = strCpuType

    def getMemoryInfo(self):
        logger.info("获取内存信息！")
        div_gb_factor = (1024.0 ** 3)
        for mem in s.Win32_ComputerSystem():
            memCap = round(int(mem.TotalPhysicalMemory)/div_gb_factor)
            logger.info("判断内存是否满足要求")
            if memCap < 8:
                logger.info("内存小于8G")
                self.isNotUsable()
                #self.updateUI()
                palette.setColor(QtGui.QPalette.WindowText, QtCore.Qt.red)
                self.label_hardware_memory_result.setPalette(palette)
                self.label_hardware_memory_result.setText(str(memCap) + 'GB')
            logger.info("内存大于等于8G")
            self.label_hardware_memory_result.setText(str(memCap)+'GB')

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
                if int(gpuMemCap) < 1:
                    logger.info("显存小于1GB")
                    self.isNotUsable()
                    # self.updateUI()
                    palette.setColor(QtGui.QPalette.WindowText, QtCore.Qt.red)
                    self.label_hardware_GPUmemory_result.setPalette(palette)
                    self.label_hardware_GPUmemory_result.setText(str(gpuMemCap) + 'GB')
                logger.info("显存大于等于1GB")
                self.label_hardware_GPUmemory_result.setText(str(gpuMemCap) + 'GB')
                gpuInfo = gpu.Name
                logger.info("设置显卡型号信息")
                self.label_hardware_GPUmodel_result.setText(gpuInfo)
                return
        self.isNotUsable()
        # self.updateUI()
        palette.setColor(QtGui.QPalette.WindowText, QtCore.Qt.red)
        self.label_hardware_GPUmodel_result.setPalette(palette)
        self.label_hardware_GPUmodel_result.setText("请安装独立显卡！")


    def isFileExisted(self, fPath, fileName, obj):
        logger.info("检查文件是否存在")
        isExisted = os.path.exists(fPath+"\\"+fileName)
        if not isExisted:
            logger.info("文件不存在")
            self.isNotUsable()
            #self.updateUI()
            palette.setColor(QtGui.QPalette.WindowText, QtCore.Qt.red)
            obj.setPalette(palette)
            obj.setText(r"未安装")
        else:
            logger.info("文件存在")
            obj.setText(r"已安装")

    def isSoftwareInstalled(self, rPath, softName, obj):
        logger.info("检查软件是否安装")
        value = self.QueryReg(rPath, softName)
        if value == True:
            logger.info("软件已经安装")
            obj.setText(r"已安装")
        else:
            logger.info("软件未安装")
            self.isNotUsable()
            #self.updateUI()
            palette.setColor(QtGui.QPalette.WindowText, QtCore.Qt.red)
            obj.setPalette(palette)
            obj.setText(r"未安装")

    def isNotUsable(self):
        logger.info("进入软件不可使用处理函数")
        strUsable = self.label_usable_result.text()
        if strUsable == r"可以使用":
            logger.info("可以使用-->无法使用")
            palette.setColor(QtGui.QPalette.WindowText, QtCore.Qt.red)
            self.label_usable_result.setPalette(palette)
            self.label_usable_result.setText(r"无法使用,请关注标红部分！")
            self.pushButton.setEnabled(False)

    def isOptimize(self):
        logger.info("进入软件优化处理函数")
        strOptimize = self.label_optimize_result.text()
        if strOptimize == r"无需优化":
            logger.info("无需优化-->需要")
            palette.setColor(QtGui.QPalette.WindowText, QtCore.Qt.red)
            self.label_optimize_result.setPalette(palette)
            self.label_optimize_result.setText(r"需要优化")


    def updateUI(self):
        logger.info("更新界面信息")
        logger.info("设置非编是否可以使用")
        strUsable = self.label_usable_result.text()
        logger.info("设置非编是否需要优化")
        strOptimize = self.label_optimize_result.text()
        logger.info("设置一键优化按钮状态")
        if (strUsable == r"可以使用") and (strOptimize == r"无需优化"):
            self.pushButton.setEnabled(False)
        #if (strUsable.find(r"无法使用") == 0) or (strOptimize == r"需要优化"):
        if (strOptimize == r"需要优化"):
            self.pushButton.setEnabled(True)
        if strOptimize == r"优化完毕！":
            self.pushButton.setEnabled(False)

    def checkRegistryInfo(self, rPath):
        logger.info("检测注册表信息")
        cpuTypeList = ['1620','2609','2620','2630','2650','other']
        self.cpuTypeList = cpuTypeList
        global regList
        for cpuType in cpuTypeList:
            #Judge the cputype of the target machine is in the cpuTypeList, if yes then save in the global variable
            if self.strCpuType[0].find(cpuType) != -1:
                logger.info("模板文件中指定CPU类型的预制注册表值读取")
                regList = self.getListData(cpuType)
            else:
                logger.info("模板文件中未找到指定的CPU类型")
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

    def QueryReg(self, rPath,value_name):
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
        logger.info("获取指定路径下所有子项的路径下的DisplayName项的值")
        for i in keylist:
            tPath = rPath + "\\" + i
            rValue = self.readFromReg(tPath, "DisplayName")
            rValue = str(rValue)
            if (rValue.find(value_name) != -1):
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
        global regList
        logger.info("一键优化按钮处理函数")
        for rKey in regList:
            self.updateReg(rKey, regList[rKey],rPath)
        palette.setColor(QtGui.QPalette.WindowText, QtCore.Qt.black)
        self.label_optimize_result.setPalette(palette)
        self.label_optimize_result.setText(r"优化完毕！")
        self.updateUI()

    def __init__(self):
        QtWidgets.QMainWindow.__init__(self)
        Ui_MainWindow.__init__(self)
        self.setupUi(self)
        self.getCpuInfo()
        self.getMemoryInfo()
        self.getGPUInfo()
        self.isFileExisted("C:\\WINDOWS\\system32", "D3DX11_41.dll", self.label_software_Directx10_result)
        self.isFileExisted("C:\\WINDOWS\\system32", "D3DX11_43.dll", self.label_software_Directx11_result)
        #check if the software is installed or not
        rPath = "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall"
        self.isSoftwareInstalled(rPath, "Microsoft Visual C++ 2010  x64 Redistributable", self.label_software_vc2010x64_result)
        rPath = "SOFTWARE\\Wow6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall"
        self.isSoftwareInstalled(rPath, "Microsoft Visual C++ 2010  x86 Redistributable", self.label_software_vc2010x86_result)
        rPath = "SOFTWARE\\Wow6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall"
        self.isSoftwareInstalled(rPath, "Setup1", self.label_software_setup1_result)
        rPath = "SOFTWARE\\Dayang\\SoftCodec"
        self.checkRegistryInfo(rPath)
        self.updateUI()
        self.pushButton.clicked.connect(lambda:self.optimizeOneKey(rPath))

if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    window = MyApp()
    window.show()
    sys.exit(app.exec_())
