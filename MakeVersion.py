#-*- coding: UTF-8 -*-
import sys
import os
import time
import traceback
from string import Template

"""
-------------------------------------------------------------------------------------------------
@@@ Use Example: python ChangeVersion.py F F F T 
@@@ Argument
      @1 是否改变基础版本的编号，若是T则在原来基础版本号的基础上加1，若是F则保留原版本号
      @2 是否改变发布版本的编号，若是T则在原来发布版本号的基础上加1，若是F则保留原版本号
      @3 是否改变发布版本次数的编号，若是T则在原来发布版本次数编号的基础上加1，若是F则保留原版本号
      @4 是否改变开发版本次数的编号，若是T则在原来开发版本次数编号的基础上加1，若是F则保留原版本号

-------------------------------------------------------------------------------------------------
"""
if __name__ == "__main__":
    if len(sys.argv) == 5:
        majvers = sys.argv[1]
        relvers = sys.argv[2]
        relcvers = sys.argv[3]
        devcvers = sys.argv[4]
    VersionText = ''
    #print('('+majvers+', '+relvers+', '+relcvers+', '+devcvers+')')
    try:
        MajorVersion = 1
        ReleaseVersion = 1
        ReleaseCVersion = 1
        DevelopCVersion = 3
        BuildTime = time.strftime("%m%d", time.localtime(time.time()))
        try:
            f = open('file_version.txt','r',encoding='utf-8')
            for line in f.readlines():
                if "filevers=" in line:
                    temp = line.split(", ")
                    if len(temp) == 4:
                        MajorVersion = int(temp[0].split('(')[1])
                        ReleaseVersion = int(temp[1][:-2])
                        ReleaseCVersion = int(temp[1][-2:])
                        DevelopCVersion = int(temp[2])
                        #DevelopCVersion = int(temp[3].split(')')[0])
                    else:
                        print("命令行参数不匹配，请重新检查输入！")
                        break
                    if majvers == 'T':
                        MajorVersion = MajorVersion + 1
                    if relvers == 'T':
                        ReleaseVersion = ReleaseVersion + 1
                    if relcvers == 'T':
                        ReleaseCVersion = ReleaseCVersion + 1
                    if devcvers == 'T':
                        DevelopCVersion = DevelopCVersion + 1
                    if ReleaseCVersion < 10:
                        MinorVersion = str(ReleaseVersion) + '0' + str(ReleaseCVersion)
                    else:
                        MinorVersion = str(ReleaseVersion) + str(ReleaseCVersion)
                    MicroVersion = DevelopCVersion
                    BuildTime = str(BuildTime)
                    if (BuildTime.find('0',0,1) == 0):
                        BuildTime = BuildTime.strip('0')
                    fileVersion = '(' + str(MajorVersion) + ', ' + str(MinorVersion) + ', ' + str(MicroVersion) + ', ' + str(BuildTime) + ')'
                    temp = line[line.find("filevers=")+9:-2]
                    line = line.replace(temp,fileVersion)
                VersionText += line
            f.close()
        except IOError as e:
            print(e.message)

        fwrite = open("file_version.txt", "w", encoding="utf-8")
        fwrite.flush()
        fwrite.write(VersionText)
        fwrite.close()
        print("成功生成版本号")

    except:
        traceback.print_exc(file=sys.stdout)
        print("无法生成版本号")