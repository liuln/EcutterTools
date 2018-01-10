#-*- coding: UTF-8 -*-
import unittest
import threading
from parameterized import parameterized
from EcutterTools import *

class EcutterToolTest(unittest.TestCase):
    app = QtWidgets.QApplication(sys.argv)
    window = MyApp()

    @parameterized.expand([
        (window, "Intel(R) Xeon(R) CPU E5-1620 v3 @ 3.50GHz", "Intel(R) Xeon(R) CPU E5-1620 v3"),
        (window, "Intel(R) Xeon(R) CPU E5-2609 v3 @ 3.50GHz", "Intel(R) Xeon(R) CPU E5-2609 v3"),
        (window, "Intel(R) Xeon(R) CPU E5-2620 v3 @ 3.50GHz", "Intel(R) Xeon(R) CPU E5-2620 v3"),
        (window, "Intel(R) Xeon(R) CPU E5-2630 v3 @ 3.50GHz", "Intel(R) Xeon(R) CPU E5-2630 v3"),
        (window, "Intel(R) Xeon(R) CPU E5-2650 v3 @ 3.50GHz", "Intel(R) Xeon(R) CPU E5-2650 v3"),
        (window, "Intel(R) Xeon(R) CPU E5-4603 v2 @ 3.50GHz", "Intel(R) Xeon(R) CPU E5-4603 v2"),
        (window,None, "Intel(R) Xeon(R) CPU E5-4603 v2"),
    ])
    def test_CPUType(self,window, cType, expType):
        cpu = CPUInfo(cType)
        m_type = window.getCPUType(cpu)
        cpuType = str.strip(m_type[0])
        self.assertEqual(cpuType,expType)

    @parameterized.expand([
        (window, "Intel(R) Xeon(R) CPU E5-1620 v3 @ 3.50GHz", 3.5),
        (window, "Intel(R) Xeon(R) CPU E5-1620 v3 @ 4.50GHz", 4.5),
    ])
    def test_CPUSpeed(self,window,cSpeed, expSpeed):
        cpu = CPUInfo(cSpeed)
        m_speed = window.getCPUSpeed(cpu)
        self.assertEqual(m_speed,float(expSpeed))

class CPUInfo():
    Name = None
    def __init__(self, Name):
        self.Name = Name

if __name__ == '__main__':
    #cpu类型单元测试
    suite = unittest.TestSuite()
    suite.addTest(EcutterToolTest("test_CPUType"))
    suite.addTest(EcutterToolTest("test_CPUSpeed"))
    runner = unittest.TextTestRunner()
    runner.run(suite)

