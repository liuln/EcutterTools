#-*- coding: UTF-8 -*-
import sys
import os
import wmi
from winreg import *
from xml.dom.minidom import parse
from mainwindow import *

#regPath = "SOFTWARE\\Dayang\\SoftCodec"
regTempPath = os.path.abspath(os.path.dirname(sys.argv[0]))
dom = parse(regTempPath+'\\registryTemplate.xml')

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
        cpuNum = 0
        for cpu in s.Win32_Processor():
            strCpuType = str(cpu.Name)
            strCpuType = strCpuType.split('@',1)
            self.label_hardware_CPUmodel_result.setText(strCpuType[0])
            result = float(strCpuType[1].split('GHz',1)[0])
            if result < 2.0:
                self.isNotUsable()
                #self.updateUI()
                palette.setColor(QtGui.QPalette.WindowText,QtCore.Qt.red)
                self.label_hardware_CPUspeed_result.setPalette(palette)
                self.label_hardware_CPUspeed_result.setText(strCpuType[1].strip() + u", 请更换主频更高的CPU!")
            else:
                self.label_hardware_CPUspeed_result.setText(strCpuType[1].strip())
            self.strCpuType = strCpuType
            cpuNum += 1
        coreNum = str(cpu.NumberOfCores * cpuNum)
        if int(coreNum) < 4:
            self.isNotUsable()
            #self.updateUI()
            palette.setColor(QtGui.QPalette.WindowText, QtCore.Qt.red)
            self.label_hardware_CPUnum_result.setPalette(palette)
            self.label_hardware_CPUnum_result.setText(coreNum)
        self.label_hardware_CPUnum_result.setText(coreNum)
        self.strCpuType = strCpuType

    def getMemoryInfo(self):
        div_gb_factor = (1024.0 ** 3)
        for mem in s.Win32_ComputerSystem():
            memCap = round(int(mem.TotalPhysicalMemory)/div_gb_factor)
            if memCap < 8:
                self.isNotUsable()
                #self.updateUI()
                palette.setColor(QtGui.QPalette.WindowText, QtCore.Qt.red)
                self.label_hardware_memory_result.setPalette(palette)
                self.label_hardware_memory_result.setText(str(memCap) + 'GB')
            self.label_hardware_memory_result.setText(str(memCap)+'GB')

    def getGPUInfo(self):
        div_gb_factor = (1024 ** 3)
        for gpu in s.Win32_VideoController():
            gpuInfo = gpu.Name
            self.label_hardware_GPUmodel_result.setText(gpuInfo)
            gpuMemCap = gpu.CurrentNumberOfColors
            gpuMemCap = int(gpuMemCap)/div_gb_factor
            if gpuMemCap is None:
                gpuMemCap = 0
            if int(gpuMemCap) < 1:
               self.isNotUsable()
               #self.updateUI()
               palette.setColor(QtGui.QPalette.WindowText, QtCore.Qt.red)
               self.label_hardware_GPUmemory_result.setPalette(palette)
               self.label_hardware_GPUmemory_result.setText(str(gpuMemCap) + 'GB')
            self.label_hardware_GPUmemory_result.setText(str(gpuMemCap)+'GB')

    def isFileExisted(self, fPath, fileName, obj):
        isExisted = os.path.exists(fPath+"\\"+fileName)
        if not isExisted:
            self.isNotUsable()
            #self.updateUI()
            palette.setColor(QtGui.QPalette.WindowText, QtCore.Qt.red)
            obj.setPalette(palette)
            obj.setText(r"未安装")
        else:
            obj.setText(r"已安装")

    def isSoftwareInstalled(self, rPath, softName, obj):
        value = self.QueryReg(rPath, softName)
        if value == True:
            obj.setText(r"已安装")
        else:
            self.isNotUsable()
            #self.updateUI()
            palette.setColor(QtGui.QPalette.WindowText, QtCore.Qt.red)
            obj.setPalette(palette)
            obj.setText(r"未安装")

    def isNotUsable(self):
        strUsable = self.label_usable_result.text()
        if strUsable == r"可以使用":
            palette.setColor(QtGui.QPalette.WindowText, QtCore.Qt.red)
            self.label_usable_result.setPalette(palette)
            self.label_usable_result.setText(r"无法使用,请关注标红部分！")
            self.pushButton.setEnabled(False)

    def isOptimize(self):
        strOptimize = self.label_optimize_result.text()
        if strOptimize == r"无需优化":
            palette.setColor(QtGui.QPalette.WindowText, QtCore.Qt.red)
            self.label_optimize_result.setPalette(palette)
            self.label_optimize_result.setText(r"需要优化")

    def updateUI(self):
        strUsable = self.label_usable_result.text()
        strOptimize = self.label_optimize_result.text()
        if (strUsable == r"可以使用") and (strOptimize == r"无需优化"):
            self.pushButton.setEnabled(False)
        if strUsable.find(r"无法使用") == 0 or strOptimize == r"需要优化":
            self.pushButton.setEnabled(True)
        if strOptimize == r"优化完毕！":
            self.pushButton.setEnabled(False)

    def checkRegistryInfo(self, rPath):
        cpuTypeList = ['1620','2609','2620','2630','2650','other']
        self.cpuTypeList = cpuTypeList
        regList = {}
        for cpuType in cpuTypeList:
            #Judge the cputype of the target machine is in the cpuTypeList, if yes then save in the global variable
            if self.strCpuType[0].find(cpuType) != -1:
                regList = self.getListData(cpuType)
                self.regList = regList
        for rKey in regList:
            regData = self.readFromReg(rPath, rKey)
            if regData == None or regData is None:
                self.isOptimize()
                #self.updateUI()
            elif str(regData) != str(regList[rKey]):
                self.isOptimize()

    def getListData(self,cpuType):
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

    def QueryReg(self, rPath,value_name):
        key = OpenKey(HKEY_LOCAL_MACHINE, rPath, 0, KEY_READ)
        keylist = []
        subKey = QueryInfoKey(key)[0]
        for i in range(int(subKey)):
            keylist.append(EnumKey(key, i))
        CloseKey(key)
        for i in keylist:
            tPath = rPath + "\\" + i
            rValue = self.readFromReg(tPath, "DisplayName")
            if value_name == rValue:
                return True
        return None

    def readFromReg(self,rPath, value_name):
        key = OpenKey(HKEY_LOCAL_MACHINE, rPath, 0, KEY_READ)
        try:
            value,type = QueryValueEx(key, value_name)
            return value
        except:
            return

    def updateReg(self, value_name, value, regPath):
        # Read from registry to see if the value_name is existed
        rValue = self.readFromReg(regPath, value_name)
        try:
            newKey = CreateKeyEx(HKEY_LOCAL_MACHINE, regPath, 0, KEY_WRITE)
        except:
            return
        value = int(value)
        if rValue == None or rValue is None:
            SetValueEx(newKey, value_name, 0, REG_DWORD, value)
        if rValue != value:
            SetValueEx(newKey, value_name, 0, REG_DWORD, value)

    def optimizeOneKey(self, rPath):
        for rKey in self.regList:
            self.updateReg(rKey, self.regList[rKey],rPath)
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
        self.isSoftwareInstalled(rPath, "Microsoft Visual C++ 2005 Redistributable (x64)", self.label_software_vc2005x64_result)
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
