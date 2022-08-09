# -*- coding: utf-8 -*-
"""
Created on Mon Jun 20 09:55:44 2022

@author: Autozhz
"""

import os,re,glob,sys,time,random,string

import pandas as pd
from ProcessOA import ProcessOA
from CommomClass import CommomClass
from PrintLog import PrintLog

class AppendOA(CommomClass):
    
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
    
    def _rename_oa(self,file,**kwargs):
        if not kwargs.get('isRename',0): return

        bn = os.path.basename(file)
        if len(bn) != len('data_oa_XXXXXX.xlsx'): # 'data_oa' not in file and
            # t1 = time.ctime(os.path.getmtime(file))
            # t2 = time.strftime('%Y%m%d%H%M%S',time.strptime(t1))
            t2 = ''.join(random.choice(string.ascii_letters) for _ in range(6))
            nd = '%s\\data_oa_%s.xlsx'%(os.path.split(file)[0],t2)
            try:
                os.rename(file,nd)
            except Exception as e:
                PrintLog.log('文件 %s 重命名错误 %s'%(os.path.relpath(file),e))
    
    def combine_record(self,oa_all_path,**kwargs):
        print('\n'+'Start 合并新增案件'.center(30,'*'))
        oa_files = glob.glob(oa_all_path+'\\*.xlsx')
        df = self.read_file(**kwargs)
        count_len = 0
        for file in oa_files:
            # print(file)
            is_version2 = kwargs.get('is_version2',0)
            if is_version2:
                cols = ['立案日期','承办法官','案号','原审案号','当事人']
            else:
                cols = ['立案日期','承办法官','案号','原一审案号','当事人']
            try:
                df0 = pd.read_excel(file,usecols=cols,
                                    na_filter=False) # 记得关闭
            except Exception as e:
                PrintLog.log('文件 %s 格式错误 %s'%(os.path.relpath(file),e))
                continue
            df = pd.concat([df0,df])
            count_len += df0.shape[0]
            PrintLog.log('File %s, count read rows to %s'%(os.path.relpath(file),count_len))
            self._rename_oa(file,**kwargs)
        df.rename(columns = {'原一审案号':'原审案号'}, inplace=True)
        df = self.sort_duplicates_data(df,sortlist=['立案日期','案号'])
        self.save_file(df_output=df,
                       **kwargs)
        print('END'.center(30,'*') + '\n')

        return df
    
if __name__ == '__main__':
    from Global import *
    AppendOA1 = AppendOA()
    df = AppendOA1.combine_record(oa_all_path,
                                  is_version2 = is_version2,
                                  df_input_name=df_oa_path,
                                  df_output_name=df_oa_path,
                                  isRename=0,
                                  isSave=1)

