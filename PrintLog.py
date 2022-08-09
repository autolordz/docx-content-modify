# -*- coding: utf-8 -*-
"""
Created on Sun Nov 21 20:20:49 2021

@author: autol
"""

import os,re,sys,time
import pandas as pd
from glob import glob
import logging
import datetime
from Global import current_path

#%%

class PrintLog:
    def __init__(self, *args, **kwargs):
        
        self.isLogFile = kwargs.get('isLogFile',0)
        if self.isLogFile:
            logpath = os.path.join(current_path,'log')
            os.makedirs(logpath,exist_ok=True)
            timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M')
            #  + datetime.timedelta(5)
            self.logname = '{}\log_{}.txt'.format(logpath,timestamp)

    def log(self, msg, *args, **kwargs):
        if msg:
            msg = '%s: %s'%(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'),msg)
        print(msg, *args, **kwargs) # print in console
        if self.isLogFile:
            with open(self.logname,'a',encoding='utf-8') as file:
                print(msg, *args, **kwargs, file=file) # print to file

class PrintLog1:
    def __init__(self):
        Log = 'log'
        today = datetime.date.today().strftime("%Y%m%d")
        logpath = '%s//%s'%(os.path.join(current_path,Log),today)
        os.makedirs(logpath,exist_ok=True)
        logfile = '%s//log.txt'%logpath
        self.logger = logging.getLogger(Log)
        self.logger.setLevel(logging.INFO)
        handler=logging.FileHandler(logfile)
        formatter=logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
        console=logging.StreamHandler()
        handler.setFormatter(formatter)
        console.setFormatter(formatter)
        self.logger.addHandler(handler)
        self.logger.addHandler(console)
        self.logger.setLevel(logging.DEBUG)
    def LogMsg(self,string):
        self.logger.info(string)
    def removeHandler(self):
        self.logger.handlers = []
    def print_log(*args, **kwargs):
        print(*args, **kwargs)
        with open('log.txt', "a",encoding='utf-8') as file:
            print(*args, **kwargs, file=file)
            
PrintLog = PrintLog(isLogFile=1)

if __name__ == '__main__':
    print(os.path.realpath('.'))
    print(os.path.join(os.path.realpath('.'),'log'))
    PrintLog1 = PrintLog(isLogFile=1)
    PrintLog1.log('asdasad')
    PrintLog1.log('2323')

#%%

# temp = sys.stdout 
# print('console')

# sys.stdout = open('output.txt', 'a')
# print('file')

# sys.stdout = temp
# print('console')
