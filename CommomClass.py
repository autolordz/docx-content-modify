# -*- coding: utf-8 -*-
"""
Created on Mon Jun 27 20:42:50 2022

@author: Autozhz
"""
import pandas as pd
import re,os,sys
import rsa,datetime
import win32api,win32con
from PrintLog import PrintLog
from Global import current_path

class CommomClass:
    
    def __init__(self, *args, **kwargs):
        
        try:
            with open(os.path.join(current_path, 'verification.txt'),'rb') as verif, \
                open(os.path.join(current_path, 'private.pem'),'rb') as privatefile:
                content = verif.read()
                p = privatefile.read()
                privkey = rsa.PrivateKey.load_pkcs1(p)
                
                self.dectex = rsa.decrypt(content, privkey).decode()
                if datetime.datetime.now() > \
                    datetime.datetime.strptime(self.dectex,'%Y%m%d_%H%M'):
                    # print("APP Expired, Please Contact Admin, Exit!! ")
                    self.exit_msg('APP过期, 请联系管理员, 退出！')
        except Exception as e:
            PrintLog.log(e)
            sys.exit()
        
        self.df = None
        pass
    
    def check_expired(self):
        diff_t1 = datetime.datetime.strptime(self.dectex,'%Y%m%d_%H%M')
        PrintLog.log('APP过期时间',diff_t1)
    
    def exit_msg(self, msg):
        PrintLog.log(msg)
        win32api.MessageBox(0,msg,'提示',win32con.MB_OK)
        sys.exit()
        
    def pop_msg(self, msg):
        PrintLog.log(msg)
        win32api.MessageBox(0,msg,'提示',win32con.MB_OK)
            
    def read_file(self, **kwargs):
        df = kwargs.get('df_input',None)
        df_name = kwargs.get('df_input_name',None)
        try:
            if df is None or not df.any().any() and df_name:
                df = pd.read_excel(df_name,na_filter=False)
                # df = pd.read_csv(self.df_name,na_filter=False)
        except Exception as e:
            PrintLog.log(e)
            df = pd.DataFrame()
        self.df = df.copy()
        return df
    
    def open_path(self, path, **kwargs):
        isOpenPath = kwargs.get('isOpenPath',0)
        if isOpenPath:
            if os.path.isdir(path):
                os.system('explorer %s'%path)
            else:
                os.system('explorer %s'%os.path.dirname(path))
    
    def sort_duplicates_data(self,df,sortlist=['']):
        if set(sortlist).issubset(set(df.columns)):
            df.drop_duplicates(sortlist,keep='first',inplace=True)
            df.sort_values(sortlist,ascending=False,inplace=True)
        df.reset_index(drop=True,inplace=True)
        df.fillna('',inplace=True)
        return df
    
    def save_file(self, **kwargs):
        isSave = kwargs.get('isSave',0)
        # dfi = kwargs.get('df_input',None)
        dfi = self.df
        df = kwargs.get('df_output',None)
        save_name  = kwargs.get('df_output_name',None)
        try:
            if isSave and save_name and df is not None:
                if dfi is not None and df.equals(dfi):
                    print('No Updated with File %s'%os.path.relpath(save_name))
                else:
                    df.to_excel(save_name,index=0)
                    # df.to_csv(self.df_name,encoding='utf_8_sig',index=0)
                    PrintLog.log('更新 %s 数据 【%s】 条'%(os.path.basename(save_name),len(df)))
                    return 1
        except Exception as e:
            PrintLog.log(e)
        return 0
    
    def check_df_expand(self, **kwargs):
        df = self.read_file(**kwargs)
        if not df['联系'].any():
            df_name = kwargs.get('df_input_name',None)
            self.pop_msg('正在打开 df_expand，请补充代理人、联系电话等信息')
            ret = os.system('start excel %s'%df_name)
            if ret > 0:
                self.pop_msg('没有安装打开表格的 excel，请手动打开df_expand')
                os.system('explorer %s'%os.path.dirname(df_name))
            sys.exit()
        else:
            print('Continute to Export')
            
    
    def split_list(self, regex,L):
        return list(filter(None,re.split(regex,L)))

if __name__ == '__main__':
    CommomClass1 = CommomClass()
    CommomClass1.check_expired()
    
    
    