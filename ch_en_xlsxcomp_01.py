#!/usr/bin/python
#encoding=utf-8

import os
import os.path
from openpyxl import load_workbook
from openpyxl import Workbook


def ch_en_xlsxcomp():
    """
    根据给定的中英文对应表格，完成中英文表头的对应输出
    """
    conf_file = open("conf.data","r")
    conf_lines = conf_file.readlines()
    ch_dir = ""
    en_dir = ""
    if len(conf_lines) == 2:
        ch_dir = conf_lines[0].strip()
        en_dir = conf_lines[1].strip()
    else:
        print "conf.data error!"
    #获取文件名    
    filename_list = []
    for parent,dirnames,filenames in os.walk(ch_dir):
        for filename in filenames:
            filename_list.append(filename)
    
    #
    for filename in filename_list:
        ch_xlsx_name = ch_dir + filename
        en_xlsx_name = en_dir + filename
        print "当前处理文件：",filename
        
        
        
if __name__ == "__main__":
    ch_en_xlsxcomp()
        
    

