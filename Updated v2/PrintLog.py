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
Log = 'Log'

#%%

class PrintLog:
    def __init__(self, *args, **kwargs):
        
        self.isLogFile = kwargs.get('isLogFile',0)
        if self.isLogFile:
            curpath = os.path.realpath('.')
            logpath = curpath+'\logfolder'
            os.makedirs(logpath,exist_ok=True)
            timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M')
            self.logname = '{}\log_{}.txt'.format(logpath,timestamp)
        pass

    def log(self,*args, **kwargs):
        print(*args, **kwargs) # print in console
        if self.isLogFile:
            with open(self.logname, "a",encoding='utf-8') as file:
                print(*args, **kwargs, file=file) # print to file
        

class PrintLog1:
    def __init__(self):
        today = datetime.date.today().strftime("%Y%m%d")
        logpath = '%s//%s'%(Log,today)
        logfile = '%s//log.txt'%logpath
        os.makedirs(logpath,exist_ok=True)
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
            
PrintLog = PrintLog(isLogFile=0)

# PrintLog.LogMsg('asdasad')
# PrintLog.LogMsg('2323')
# PrintLog.removeHandler()

