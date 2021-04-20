# -*- coding: utf-8 -*-
"""
Created on Wed Sep 11 15:23:05 2019

@author: autol
"""
#%%
import os,re,sys
import pandas as pd
from dcm_configure import write_config,read_config

#%%
pd.set_option('max_colwidth',500)
pd.set_option('max_rows', 50)
pd.set_option('max_columns',50)

#%% global variable

titles_main = ['立案日期','适用程序','案号','原一审案号','判决书源号','主审法官','当事人','诉讼代理人','地址',]
titles_oa = ['立案日期','案号','原一审案号','承办人','当事人','适用程序']
#titles_oa = ['立案日期','案号','原审案号','承办法官','当事人','适用程序']

titles_cn = ['立案日期','案号','当事人','诉讼代理人','地址']
titles_en = ['datetime','number','uname','aname','address']
path_names_clean = re.compile(r'[^A-Za-z\u4e00-\u9fa5（）()：]') # 保留用户名和旧名 包括括号冒号
search_names_phone = lambda x: re.search(r'[\w（）()：:]+\_\d+',x)  # tel numbers 电话号
path_code_ix = re.compile(r'[(（][0-9]+[)）].*?号') # case numbers 案号
postal_path = os.path.join('.','postal')
jdocs_path = os.path.join('.','jdocs')
adr_tag = '/地址：' # 地址标识，分割用
done_tag = '_集合'
usrtag = r'申请人|被申请人|原告|被告|原审被告|上诉人|被上诉人|第三人|原审诉讼地位|申请再审人|被申请再审人' # 当事人抬头标识

#    dr = dict((reversed(item) for item in dd.items()))
#%%

def getcolums_en(df_columns):
    dd = dict(zip(titles_cn,titles_en))
    return list(filter(None,(dd.get(x) for x in df_columns.tolist())))


# sample
#y = getcolums_en(df.columns)
#
#from functools import reduce
#reduce((lambda x,y: x + 2), [1, 1, 1, 1])
#
#fib = lambda n:reduce(lambda x,n:[x[1],x[0]+x[1]], range(n),[0,1])[0]
#for x in range(1,100):
#    print(fib(x))


#.copy()
#%% print_log log

logname = 'log.txt'

def print_log(*args, **kwargs):
    print(*args, **kwargs)
    with open(logname, "a",encoding='utf-8') as file:
        print(*args, **kwargs, file=file)

def input_exit(*args, **kwargs):
    '''输入并退出'''
    input(*args, **kwargs);
    sys.exit()
    return 1

if os.path.exists(logname):
    os.remove(logname)

#%% read configure global variable

def init_var():
    cfgfile = 'conf.txt'
    try:
#        if not os.path.exists(cfgfile): write_config(cfgfile) # 生成默认配置
        var = pd.Series(read_config(cfgfile));var
    except Exception as e:
        print_log('>>> 配置文件出错 %s ,删除...'%e)
        if os.path.exists(cfgfile):
            os.remove(cfgfile)
        try:
            write_config()
            var = pd.Series(read_config());var
        except Exception as e:
            '''这里可以添加配置问题预判问题'''
            input_exit('>>> 配置文件再次生成失败 %s ...'%e)
#    print(var)
    return var

var = init_var()
#locals().update(var.to_dict()) # 设置读取的全局变量
