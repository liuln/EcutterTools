#-*- coding: UTF-8 -*-
import sys
import os
import logging
import wmi
from winreg import *
from xml.dom.minidom import parse
from mainwindow import *

#创建一个logger
logger = logging.getLogger('mylogger')
logger.setLevel(logging.DEBUG)
#创建一个handler，用于写入日志文件
fh = logging.FileHandler('EcutterTools.log')
fh.setLevel(logging.DEBUG)
#定义handler的输出格式
formatter = logging.Formatter('[%(asctime)s][%(thread)d][%(filename)s][line: %(lineno)d][%(levelname)s] ## %(message)s')
fh.setFormatter(formatter)
#给logger添加handler
logger.addHandler(fh)

class MyApp(QtWidgets.QMainWindow, Ui_MainWindow):
    global s
    s = wmi.WMI()
    global div_gb_factor, strCpuType
        #, cpuTypeList
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

    def getCpuInfo(self, cpus):
        logger.info("获取CPU硬件信息！")
        self.strCpuType = "初始状态"
        for cpu in cpus:
            speed = self.getCPUSpeed(cpu)
            cpuType = self.getCPUType(cpu)
            logger.info("判断CPU主频是否满足需求！")
            if speed < 2.0:
                logger.info("CPU主频小于2.0GHz！")
                self.isNotUsable()
                # self.updateUI()
                palette.setColor(QtGui.QPalette.WindowText, QtCore.Qt.red)
                self.label_hardware_CPUspeed_result.setPalette(palette)
                self.label_hardware_CPUspeed_result.setText(cpuType[1].strip() + u", 请更换主频更高的CPU!")
            else:
                logger.info("CPU主频大于等于2.0GHZ!")
                self.label_hardware_CPUspeed_result.setText(cpuType[1].strip())
            logger.info("处理多个CPU的情况！")
            if self.strCpuType == "初始状态":
                logger.info("设置CPU型号为: %s", cpuType[0])
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
        logger.info("获取CPU的核数！")
        coreNum = self.getCPUNum(cpu,len(cpus))
        if int(coreNum) < 4:
            logger.info("CPU的核数小于4，要设置红色字提示当前计算机的硬件环境信息不可以使用非编软件！")
            self.isNotUsable()
            # self.updateUI()
            palette.setColor(QtGui.QPalette.WindowText, QtCore.Qt.red)
            self.label_hardware_CPUnum_result.setPalette(palette)
            self.label_hardware_CPUnum_result.setText(coreNum)
        logger.info("CPU核数满足使用非编软件的最低硬件环境需求！")
        self.label_hardware_CPUnum_result.setText(coreNum)

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
        if strUsable.find(r"无法使用") == 0 or strOptimize == r"需要优化":
            self.pushButton.setEnabled(True)
        if strOptimize == r"优化完毕！":
            self.pushButton.setEnabled(False)

    def checkRegistryInfo(self, rPath):
        logger.info("检测注册表信息")
        cpuTypeList = ['1620','2609','2620','2630','2650','other']
        self.cpuTypeList = cpuTypeList
        regList = {}
        for cpuType in cpuTypeList:
            #Judge the cputype of the target machine is in the cpuTypeList, if yes then save in the global variable
            if self.strCpuType[0].find(cpuType) != -1:
                logger.info("模板文件中指定CPU类型的预制注册表值读取")
                regList = self.getListData(cpuType)
                self.regList = regList
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
        palette.setColor(QtGui.QPalette.WindowText, QtCore.Qt.black)
        self.label_optimize_result.setPalette(palette)
        self.label_optimize_result.setText(r"优化完毕！")
        self.updateUI()

    def __init__(self):
        logger.info("QMainWindow初始化")
        QtWidgets.QMainWindow.__init__(self)
        logger.info("Ui_MainWindow初始化")
        Ui_MainWindow.__init__(self)
        logger.info("MyApp setupUi...")
        self.setupUi(self)
        logger.info("Load UI...")
        self.loadUI()

    def loadUI(self):
        logger.info("获取CPU信息...")
        self.getCpuInfo(s.Win32_Processor())
        logger.info("获取内存信息...")
        self.getMemoryInfo()
        logger.info("获取GPU信息...")
        self.getGPUInfo()
        logger.info("判断DirectX 10是否安装")
        self.isFileExisted("C:\\WINDOWS\\system32", "D3DX11_41.dll", self.label_software_Directx10_result)
        logger.info("判断DirectX 11是否安装")
        self.isFileExisted("C:\\WINDOWS\\system32", "D3DX11_43.dll", self.label_software_Directx11_result)
        #check if the software is installed or not
        logger.info("判断VC2005 x64是否安装")
        rPath = "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall"
        self.isSoftwareInstalled(rPath, "Microsoft Visual C++ 2005 Redistributable (x64)", self.label_software_vc2005x64_result)
        logger.info("判断Setup1是否安装")
        rPath = "SOFTWARE\\Wow6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall"
        self.isSoftwareInstalled(rPath, "Setup1", self.label_software_setup1_result)
        logger.info("检查注册表信息")
        rPath = "SOFTWARE\\Dayang\\SoftCodec"
        self.checkRegistryInfo(rPath)
        logger.info("更新UI")
        self.updateUI()
        logger.info("点击一键优化按钮")
        self.pushButton.clicked.connect(lambda:self.optimizeOneKey(rPath))

if __name__ == "__main__":
    logger.info("Enter into main func...")
    regTempPath = os.path.abspath(os.path.dirname(sys.argv[0]))
    dom = parse(regTempPath + '\\registryTemplate.xml')
    logger.info("QApplication...")
    app = QtWidgets.QApplication(sys.argv)
    logger.info("MyApp...")
    window = MyApp()
    logger.info("MyApp show...")
    window.show()
    logger.info("QApplication exit...")
    sys.exit(app.exec_())
