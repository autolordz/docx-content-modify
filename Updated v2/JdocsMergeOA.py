# -*- coding: utf-8 -*-
"""
Created on Wed Jun 22 16:45:09 2022

@author: Autozhz
"""

import os,re,sys
import pandas as pd
from CommomClass import CommomClass

class JdocsMergeOA(CommomClass):
    
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
    
    def fill_jdoc_oa(self,**kwargs):
        print('\n'+'Start 填充判决书内容'.center(30,'*'))

        df_jdoc = self.read_file(
            df_input = kwargs.get('df1',None),
            df_input_name = kwargs.get('df1_path',None))
        
        df_expand = self.read_file(
            df_input = kwargs.get('df2',None),
            df_input_name = kwargs.get('df2_path',None))

        if not (df_jdoc.any().any() and df_expand.any().any()):
            # msg = 'No df_jdoc or df_expand xlsx found, please run step 1 to regenerate it.'
            msg = '找不到文件 df_jdoc_extract 或 df_oa_extract, 请运行 Step 1 重新生成.'
            self.exit_msg(msg)            
        
        aa = pd.merge(df_expand,
                         df_jdoc[['原审案号','当事人','地址']],
                         how='left',on=['原审案号','当事人'])
        bb = pd.merge(df_expand,
                         df_jdoc[['当事人','地址']],
                        how='left',on=['当事人'])
        df = pd.concat([aa,bb])
        # df.drop_duplicates(['原审案号','当事人'],keep='first',inplace=True)
        # df.reset_index(drop=True,inplace=True)
        if '地址_x' in df.columns:
            df['地址_x'].replace('',None,inplace=True)
            df['地址'] = df['地址_x'].combine_first(df['地址_y'])
            df.drop(['地址_x','地址_y'],axis=1,inplace=True)
        
        df = self.sort_duplicates_data(df,sortlist=['立案日期','案号','当事人'])
        self.save_file(df_output = df,**kwargs)
        print('END'.center(30,'*') + '\n')

        return df
