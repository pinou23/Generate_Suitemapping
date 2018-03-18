# -*- coding: UTF-8 -*-
'''
Created on 2015年11月5日
功能：把某个路径下所有的TA case的信息（id、name）放到一个Excel表格里
@author: pacao
'''

import os
import sys
import shutil
from robot.utils import asserts
from robot.api import logger
import re
import types
from robot.api import TestData
from xlwt import *
from xlrd import open_workbook
from xlutils.copy import copy


LOGLIST = []
info_list=[]

'''创建一个excel'''
def create_excel(path):
    excel_file = Workbook()
    ws = excel_file.add_sheet('Case Info')
    ws.write(0,0,'Test Suite')
    ws.write(0,1,'Case Name')
    ws.write(0,2,'Run')
    ws.write(0,3,'Timeout(min)')
    ws.write(0,4,'Instance_id')
    ws.write(0,5,'Responsible Tester')
    ws.write(0,6,'Case Tag')
    ws.write(0,7,'ENV')
    ws.write(0,8,'PC IP')
    ws.write(0,9,'Comment')
    
    file_name = 'Case Info.csv'
    file_path = path+'\\'+file_name
    #print file_path
    excel_file.save(file_path)
    return file_path

def write_excel(file_path):
    rb = open_workbook(file_path)
    wb = copy(rb)
    ws = wb.get_sheet(0)
    flag = 0
    for info in info_list:
        flag = flag+1
        print info[1]
        ws.write(flag,0,info[2])
        ws.write(flag,1,info[3])
        ws.write(flag,2,1)
        ws.write(flag,3,80)
        ws.write(flag,4,int(info[0]))
        ws.write(flag,5,info[1])
      
        
    wb.save(file_path)
    print 'DONE!'
    return wb

#def getTestPath(path):
    



    
def parseTestcase(file_):
    """This KW is used for change QCID tags in a Testcase file
    This KW should be used after connected DB

     | Input Parameters | Man. | Description |
     | file_ | Y | Testcase html file |
     | build | Y | the build which you want to porting QCID to |
     | cur | Y | get cur from KW connectDB() |

    return true or false
    """
#    recordLogsToList('Transforming QCID of %s to build %s' % (file_,build))
    #resetid = re.compile('qc_test_set_id: *?([\d-]+)')
    reinsid = re.compile('QC_ *?([\d-]+)')
#    rehwtag = re.compile(r'TL16.*|Release.*|TL15A.*',re.I)
    global DELETECASELIST
    global MODIFYFOLDERLIST

    try:
        suite = TestData(source = "%s" %(file_))
    except Exception:
        recordLogsToList('Warning: TestData analyze file [%s] Failed' % file_)
        return False
    for mytestcase in suite.testcase_table:
        if not mytestcase.tags.value:
            recordLogsToList('%s QCID is missed,Please input it in your script!' %file_)
            return False
        #print mytestcase.name
        else:
            case_name = mytestcase.name
            print case_name
            for tag in mytestcase.tags.value:
                if 'QC_' in tag:
                    #setid = resetid.findall(tag)
                    insid = reinsid.findall(tag)
                    TAG_LOCATION = mytestcase.tags.value.index(tag)
                    
                 
                    if insid:
                        qcid = insid[0]
                    else:
                        qcid = -1
                    
                    for ftag in suite.setting_table.force_tags.value:
                        if 'Owner-'in ftag or 'owner-' in ftag:
                                owner = re.findall(r'[O|o]wner-(.*)@.*.com',ftag)
                                print owner
                    info = (qcid,owner,file_[14:],case_name)
                    info_list.append(info)
                    break
    return True

def TraversalScriptPath(apath):
    """This KW is used for Transform all testcase files in a folder to new QCID of the build.
    This KW will delete all SVN path in parameter 'path'

    | Input Parameters | Man. | Description |
    | path | Y | testcase files path |
    | build | Y | the build which you want to porting QCID to |

    return true or false
    """
    
    path = apath.replace('\\','/')
#    recordLogsToList(r'Transforming path----%s'%path)
    print path
    if not os.path.exists(path):
        
        recordLogsToList('%s is not exist!' % path)
        return False
    if os.path.isfile(path):
        if '.html' in path:
            print 'path:',path[14:]
            parseTestcase(path)
    elif os.path.isdir(path):
        if '.svn' in path:
            pass
        else:
            searchfile = os.listdir(path)
            for vpath in searchfile:
                childpath = path + '/' + vpath
                TraversalScriptPath(childpath)
    else:
        recordLogsToList('%s is an unknown object,I can not handle it!' % path)
    

    return True



def recordLogsToList(log):
    """This KW is used for recordlogs to global log list and print it

    """
    print log
#    global LOGLIST
    LOGLIST.append(log)



excel_path = r'D:\test'
file_path = create_excel(excel_path)
path = r'D:\TA_Scripts\TL18\CIT'
    
TraversalScriptPath(path)
write_excel(file_path)

#     file_path = create_excel(excel_path)
#     write_excel(file_path,info_list)

