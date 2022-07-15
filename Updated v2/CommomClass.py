# -*- coding: utf-8 -*-
"""
Created on Mon Jun 27 20:42:50 2022

@author: Autozhz
"""
import pandas as pd
import re,os,sys
import rsa,datetime
import win32api,win32con


class CommomClass:
    
    def __init__(self, *args, **kwargs):
        
        content = b'\x13\xbe\x1a\xd9\xed\xe9(\x1f\xe9\xe6U\xc0\x82\xe0\xc5\xef0=&^P\x18\xf9\x90\xa6\xf4\xd7\xd0\x96V:H?.D\xfe\xc2k\xa3\xa8C\x84C\xe6\x8epO\x98@\xabm\xb9\x8d\x07\xd9\x85"=d\xd1\x13\x98\xe5\xe9\xb8\x1ex{\x18\x13\x1d\xc1D\x1c3\xae\x9a}\xd5\xf4\xf5\x15l>\x0bP\x17\xb81\xe0\xcb\xca\xa8{\xfa\xf5\xc3\xbd7\xe3\xe3\x1bITi\xecla\xc0>7\xa6Le<\xf6"\x91q\x8f3\xbc\x86 %\x91\x9a\x86'
        
        with open('private.pem','rb') as privatefile:
             p = privatefile.read()
        privkey = rsa.PrivateKey.load_pkcs1(p)
        dectex = rsa.decrypt(content, privkey).decode()
        
        if datetime.datetime.now() > \
            datetime.datetime.strptime(dectex,'%Y%m%d_%H%M'):
            # print("APP Expired, Please Contact Admin, Exit!! ")
            self.exit_msg('APP过期, 请联系管理员, 退出！')
        self.df = None
        pass
    
    def exit_msg(self, msg):
        win32api.MessageBox(0,msg,'提示',win32con.MB_OK)
        sys.exit()
        
    def pop_msg(self, msg):
        win32api.MessageBox(0,msg,'提示',win32con.MB_OK)
            
    def read_file(self, **kwargs):
        df = kwargs.get('df_input',None)
        df_name = kwargs.get('df_input_name',None)
        try:
            if df is None or not df.any().any() and df_name:
                df = pd.read_excel(df_name,na_filter=False)
                # df = pd.read_csv(self.df_name,na_filter=False)
        except Exception as e:
            print(e)
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
                    print('No Updated with File %s'%save_name)
                else:
                    df.to_excel(save_name,index=0)
                    # df.to_csv(self.df_name,encoding='utf_8_sig',index=0)
                    print('保存 %s 数据 【%s】 条'%(os.path.basename(save_name),len(df)))
                    return 1
        except Exception as e:
            print(e)
        return 0
    
    def check_df_expand(self, **kwargs):
        df = self.read_file(**kwargs)
        if not df['联系'].any():
            df_name = kwargs.get('df_input_name',None)
            self.pop_msg('正在打开 df_expand，请补充代理人、联系电话等信息')
            try:
                os.system('start excel %s'%df_name)
            except Exception:
                pass
            sys.exit()
        else:
            print('Continute to Export')
            
    
    def split_list(self, regex,L):
        return list(filter(None,re.split(regex,L)))