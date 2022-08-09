# -*- coding: utf-8 -*-
"""
Created on Mon Jun 20 10:11:36 2022

@author: Autozhz
"""
import os,re,sys
import pandas as pd
from CommomClass import CommomClass

class ProcessOA(CommomClass):
    
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
    
    def append_record(self,**kwargs):
        print('\n'+'Start 展开案件内容'.center(30,'*'))

        df_oa = self.read_file(
                    df_input = kwargs.get('df1',None),
                    df_input_name = kwargs.get('df1_path',None))
        df = self.read_file(
                    df_input = kwargs.get('df2',None),
                    df_input_name = kwargs.get('df2_path',None))

        if df_oa.any().any():
            is_version2 = kwargs.get('is_version2',0)
            if is_version2:
                df_oa = self._process_data_ver2(df_oa)
            else:
                df_oa = self._process_data(df_oa)
            if df.any().any():
                expandcases = list(set(df['案号'].to_list())) # 展开案号,普通案号不用展开,保证唯一
                expanduname = list(set(df['当事人'].to_list())) # 展开案号,普通案号不用展开,保证唯一
                add = df_oa[~(df_oa['案号'].isin(expandcases) &
                              df_oa['当事人'].isin(expanduname))].reset_index(drop=True) # 新增条目
                if not add.any().any():
                    print('Nothing to update.')
                    return df
            else:
                add = df_oa.copy()
            df = pd.concat([df,add]).reset_index(drop=True)
            df = self.sort_duplicates_data(df,sortlist=['立案日期','案号','当事人'])
            self.save_file(df_output = df,**kwargs)
        else:
            print('Nothing for df_oa data.')
        print('END'.center(30,'*') + '\n')
        return df
    
    def _process_data_ver2(self,df):
        user = df['当事人'].str.strip().str.split(r'[,\/]',expand=True).stack().reset_index() #  # divide user by slash, old is [,，。]
        user['角色'] = user.iloc[:,2].str.extract(r'(\w+)')
        user['当事人'] = user.iloc[:,2].str.extract(r'((?<=\]).*)')
        df.reset_index(inplace=True)
        df.rename(columns={'index': 'level_0'},inplace=True)  
        df_x = pd.merge(df[['level_0','立案日期','案号','原审案号','承办法官']],
                        user[['level_0','角色','当事人']],
                        how='left',on=['level_0'])
        df_x.drop('level_0',axis=1,inplace=True)
        df_x['代理人'] = '';df_x['律所'] = '';df_x['联系'] = '';df_x['地址'] = '';
        return df_x

if __name__ == '__main__':
    from Global import *
    ProcessOA1 = ProcessOA()
    df_expand = ProcessOA1.append_record(
                            df1_path = df_oa_path,
                            df2_path = df_expand_path,
                            df_output_name = df_expand_path,
                            is_version2 = is_version2,
                            isSave=1)


    
    # def _process_data(self,df):
    #     '''获取 datetime|number 获取所有用户名包括曾用名'''
    #     user = df[['立案日期','案号','原审案号','承办法官','当事人']].reset_index()
    #     user.columns.values[0] = 'level_0'
    #     userx = user['当事人'].str.strip().str.split(r'[,，。]',expand=True).stack().reset_index() #  # divide user by slash, old is [,，。]
    #     userx[['角色','当事人']] = userx.iloc[:,2].str.strip().str.split(r'[:]',expand=True)# divide character only with [:] 
    #     userx = userx.reset_index()
    #     userx_sub = userx.loc[:,'当事人'].str.strip().str.split(r'[、]',expand=True).stack().reset_index()
    #     userx_sub.columns.values[0] = 'index'
    #     userx_sub.columns.values[2] = '当事人'
    #     userx = pd.merge(userx[['index','level_0','角色']],
    #                      userx_sub[['index','当事人']],
    #                      how='left',on=['index'])
    #     df_x = pd.merge(user[['level_0','立案日期','案号','原审案号']],
    #                     userx[['level_0','角色','当事人']],
    #                     how='left',on=['level_0'])
    #     # df_x['当事人'] = df_x['当事人'].str.strip().apply(lambda x: re.sub(r'\[.*\]','',x)).apply(lambda x: re.sub(r'等','',x)) #去掉 等
    #     df_x.drop('level_0',axis=1,inplace=True)
    #     df_x['代理人'] = '';df_x['律所'] = '';df_x['联系'] = '';df_x['地址'] = '';
    #     return df_x


