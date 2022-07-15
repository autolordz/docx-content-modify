# -*- coding: utf-8 -*-
"""
Created on Fri Jun 24 11:49:14 2022

@author: Autozhz
"""
import pandas as pd
import sys
from CommomClass import CommomClass

class CombineDuplicate(CommomClass):
    
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
    
    def combine_codes_user(self,**kwargs):
        print('\n'+'Start 合并案号、当事人'.center(30,'*'))

        df = self.read_file(**kwargs)
        if df.any().any():
            df['代理人'].replace('',None,inplace=True)
            df['代理人'] = df['代理人'].combine_first(df['当事人'])
            if df.duplicated('代理人').any():
                df = self.combine_codes_user_x(df,keys = ['代理人'],combine_key = '当事人',add_key = ['案号','承办法官'])
                df = self.combine_codes_user_x(df,keys = ['代理人','承办法官'],combine_key = '案号',add_key = ['当事人'])
                df = df[['案号','承办法官','当事人','代理人','律所','联系','地址']]
                for i, item in df.iterrows():
                    if item['当事人'] == item['代理人']:
                        df.at[i, '代理人'] = ''
                df = self.sort_duplicates_data(df,sortlist=['案号','当事人'])
                self.save_file(df_output = df,**kwargs)
        else:
            print('Nothing for df_expand data.')

        print('END'.center(30,'*') + '\n')

        return df
    
    def combine_codes_user_x(self,df,keys = [],combine_key = '',add_key = []):
        lasttext = dict(zip(keys,list(range(len(keys)))))
        kk = keys.copy()
        kk.append(combine_key)
        df.sort_values(kk,ascending=True,inplace=True)
        count = 0; 
        for i, item in df.iterrows():
            l1 = item[keys].values.tolist()
            l2 = list(lasttext.values())
            l1.sort();l2.sort()
            if l1 == l2:
                count = count + 1
            else:
                for key in lasttext.keys():
                    lasttext[key] = item[key]
                count = 0
            df.at[i, 'level_2'] = count

        user = df[['level_2']+keys+[combine_key]].set_index(['level_2']+keys)
        user = user.unstack('level_2') #.fillna('')
        user = user.droplevel(0,axis=1)
        user[combine_key] = user[user.columns] \
            .apply(lambda x: ','.join(sorted(list(set(x.dropna().astype(str))),reverse=1)), axis=1)   
        user.reset_index(inplace=True)
        df = pd.merge(
            df[keys+['律所','联系','地址']+add_key],
                          user[keys+[combine_key]],
                          how='right',on=keys)
        df.drop_duplicates(inplace=True)
        return df
    
# from Global import *
# CombineDuplicate1 = CombineDuplicate()
# df = CombineDuplicate1.combine_codes_user(
#                                 df_input_name = df_expand_path,
#                                 df_output_name = df_combine_path,
#                                 isSave=0)

