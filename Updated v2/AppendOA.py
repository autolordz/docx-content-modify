# -*- coding: utf-8 -*-
"""
Created on Mon Jun 20 09:55:44 2022

@author: Autozhz
"""

import os,re,glob,sys,time,random,string

import pandas as pd
from ProcessOA import ProcessOA
from CommomClass import CommomClass

class AppendOA(CommomClass):
    
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
    
    def _rename_oa(self,file):
        bn = os.path.basename(file)
        if len(bn) != len('data_oa_XXXXXX.xlsx'): # 'data_oa' not in file and
            # t1 = time.ctime(os.path.getmtime(file))
            # t2 = time.strftime('%Y%m%d%H%M%S',time.strptime(t1))
            t2 = ''.join(random.choice(string.ascii_letters) for _ in range(6))
            nd = '%s\\data_oa_%s.xlsx'%(os.path.split(file)[0],t2)
            try:
                os.rename(file,nd)
            except Exception as e:
                print(e)
    
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
            df0 = pd.read_excel(file,usecols=cols,
                                na_filter=False) # 记得关闭
            df = pd.concat([df0,df])
            count_len += df0.shape[0]
            print('Read oa data rows',count_len)
            self._rename_oa(file)
        df.rename(columns = {'原一审案号':'原审案号'}, inplace=True)
        df = self.sort_duplicates_data(df,sortlist=['立案日期','案号'])
        self.save_file(df_output=df,
                       **kwargs)
        print('END'.center(30,'*') + '\n')

        return df



