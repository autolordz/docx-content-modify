# -*- coding: utf-8 -*-
# Copyright (c) 2018 Autoz https://github.com/autolordz

# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:

# The above copyright notice and this permission notice shall be included in all
# copies or substantial portions of the Software.

# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
# SOFTWARE.

"""
Created on Wed Sep 11 17:33:29 2019

@author: autol
"""

#%%
import os
import pandas as pd

import dcm_util as ut
from dcm_globalvar import *

locals().update(var.to_dict()) # 设置读取的全局变量

from dcm_df_progress import df_fill_infos,df_oa_append,df_read_fix,merge_group_cases
from dcm_df_transform import df_check_format,df_transform_stream
from dcm_df_transform import make_adr,make_agent,merge_user,merge_usr_agent_adr
from dcm_df_transform import reclean_data,sort_data
from dcm_print_postal import fill_postal_save


from dcm_df_transform import clean_rows_adr,clean_rows_aname


import datetime

print('''
邮单机，自动填充判决书

Postal Notes Automatically Generate App

Updated on %s

Depends on: python-docx,pandas,StyleFrame,configparser

@author: Autoz (autolordz@gmail.com)
'''%datetime.datetime.now().strftime("%Y-%m-%d %H:%M"))


print_log('''>>> 正在处理...
    主表路径 = %s
    指定案件 = %s
    指定日期 = %s
    指定条数 = %s
    '''%(os.path.abspath(data_xlsx),
        data_case_codes,
        data_date_range,
        data_last_lines,
        )
    )

if not os.path.exists(data_xlsx):
    r = ut.save_adjust_xlsx(pd.DataFrame(columns=titles_main),data_xlsx,width=40)
    print_log('>>> %s 文件不存在...重新生成'%(data_xlsx) + r)


#try:
#%% 
print('读取数据'.center(30, '*'))
df = pd.read_excel(data_xlsx).dropna(how='all');df  # 删除全空的行
df = df_read_fix(df);df # fix empty data columns
print('>>> 有Main数据共%s条'%len(df))

#%%
dfoa = pd.read_excel(data_oa_xlsx)[titles_oa].dropna(how='all');dfoa# only some columns
print('>>> 有OA数据共%s条'%len(dfoa))

#%%
print('新增OA记录'.center(30, '*'))
df,ocodes = df_oa_append(df) # append oa data and sava # 新增OA记录

#%%
#print('合并系列案逻辑部分'.center(30, '*'))
#df = merge_group_cases(df,ocodes);df # 合并系列案逻辑部分 and save 

#%%
#from glob import glob
#import dcm_util as ut
#from dcm_get_jdocs import get_all_jdocs
#docs = glob(ut.parse_subpath(jdocs_path,'*.docx')) # get jdocs
#dfj = get_all_jdocs(docs)

#%%
print('填充判决书内容'.center(30, '*'))
df = df_fill_infos(df) # 填充判决书内容 # filled and save

#%% df tramsfrom stream 数据转换流程
print('数据转换流程'.center(30, '*'))
#df_print = df_transform_stream(df)

df = ut.titles_trans_columns(df,titles_cn);df # 中译英方便后面处理

#if flag_check_postal:
df.apply(lambda x:df_check_format(x), axis=1)


if 0<len(df)<10:
    print_log('>>> 将要打印【%s条】=> %s '%(len(df),
                                       df['number'].to_list()))

#if len(df) and flag_to_postal:
    
    #%%
print_log('\n>>> 开始生成新数据 data_main_temp... ')

'''获取 datetime|number'''
number = df[titles_en[:2]]
number = number.reset_index()
number.columns.values[0] = 'idx0'

#%%

'''获取所有用户名包括曾用名'''

print_log('>>> 正在处理...【用户】.....')

user = df[['number','uname']]
user = user[user.number != '']
user = user.reset_index()
user.columns.values[0] = 'idx0'

# user = user.str.strip().str.split(r'[,，。]',expand=True).stack() # divide user
#user = user.str.strip().str.split(r'[:]',expand=True)# divide character

userx = user['uname'].str.strip().str.split(r'[\/]',expand=True).stack().reset_index() #  # divide user by slash, old is [,，。]
userx.columns.values[0] = 'idx0'
userx.columns.values[2] = 'uname'
userx = userx.drop(['level_1'],axis=1)

userx1 = pd.merge(user[['idx0','number']],userx,how='left',on=['idx0']).fillna('')
userx1['uname'] = userx1['uname'].str.strip().apply(lambda x: re.sub(r'\[.*\]','',x)).apply(lambda x: re.sub(r'等','',x)) #去掉 等

#%%

# agent and address
# agent_adr = df[['aname','address']]
# opt = agent_adr.any()

#%%

'''获取用户 或 代理人 的地址'''

print_log('>>> 正在处理...【地址】.....')

user = userx1.copy()

adr = df[['number','address']]

adr = adr[adr.address != '']

adrx = adr.address.str.strip().str.split(r'[,，。]',expand=True).stack()

adrx = adrx.str.strip().apply(lambda x:clean_rows_adr(x))

adrx = adrx.str.strip().str.split(r'\/地址[:：]',expand=True).fillna('')

adrx.columns = ['tmp_name','address']

fix_aname=user['uname'].tolist()

adrx['clean_tmpname'] =  adrx['tmp_name'].str.strip().apply(lambda x:clean_rows_aname(x,fix_aname)) # clean adr

adrx = adrx.reset_index().drop(['level_1','tmp_name'],axis=1)
    
adrx.columns.values[0] = 'idx0'

#make_adr(adr,fix_aname=user['uname'].tolist())

#%%
adr = adrx.copy()
adr.rename(columns={'clean_tmpname': 'uname'},inplace=True)
ad_u = pd.merge(user,adr,how='left',on=['idx0','uname']).fillna('')   


#%%

'''获取 代理人 的地址'''

print_log('>>> 正在处理...【代理人】.....')


agent = df[['number','aname']]
agent = agent[agent.aname != '']
if agent.size:
    agentx = agent.aname.str.strip().str.split(r'[,，、。]',expand=True).stack() #Series
    agentx = agentx.str.strip().str.split(r'\/',expand=True).fillna('') #DataFrame
    agentx.columns = ['uname','aname']
    agentx['clean_aname'] = agentx['aname'].str.strip().apply(lambda x: clean_rows_aname(x,fix_aname))
    dd_l = agentx['uname'].str.strip().str.split(r'、',expand=True).stack().to_frame(name = 'uname').reset_index()
    dd_r = agentx[agentx.columns.difference(['uname'])].reset_index()
    agentx = pd.merge(dd_l,dd_r,how='outer',on=['level_0','level_1']).drop(['level_1','level_2'],axis=1).fillna('')
    agentx.columns.values[0] = 'idx0'

#%%
    '''合并 用户 代理人 地址'''

if agent.size:
    agent = agentx.copy()
else:
    agent = user[['idx0','uname']]
    agent['aname'] = user['uname']
    agent['clean_aname'] = user['uname']


#%%
adr = adrx.copy()
adr.rename(columns={'clean_tmpname': 'clean_aname'},inplace=True)
ag_ad = pd.merge(agent,adr,how='left',on=['idx0','clean_aname']).fillna('')   

ag_u = pd.merge(user,agent,how='left',on=['idx0','uname']).fillna('')

ag_ad_u = pd.merge(ag_u,ag_ad,how='left',on=['idx0','uname','aname','clean_aname']).fillna('')

ag_ad_u1 = pd.merge(ag_ad_u,ad_u,how='left',on=['idx0','number','uname']).fillna('')

#%%

ag_ad_u2 =ag_ad_u1.copy()

for (i,ag_ad_u_x) in ag_ad_u2.iterrows():
    if len(ag_ad_u_x.address_x)==0 and len(ag_ad_u_x.address_y)>0:
        ag_ad_u2.at[i,'address_x'] = ag_ad_u2.at[i,'address_y']
        

ag_ad_u2.rename(columns={'address_x': 'address'},inplace=True)
ag_ad_u2.drop(['address_y'],axis=1,inplace=True) 

ag_ad_u = pd.merge(ag_ad_u2,df[['datetime','number']],how='left',on=['number']).fillna('')

        
#%%
    

#if not agent.size:
#    input_exit('>>> !!! 代理人为空，退出.... !!!')

#%%
  
 #%%

 
#%%   
#ag_u['clean_aname'].replace('',float('nan'),inplace=True)
#ag_u['clean_aname'] = ag_u['clean_aname'].fillna(ag_u['uname']).replace(path_names_clean,'')

#ad_u['clean_aname'] = ad_u['uname'].apply(lambda x: clean_rows_aname(x,u_a['clean_aname'].tolist()))

#ag_ad_u = pd.merge(ag_u,ad_u,how='left',on=['idx0','number','uname','clean_aname']).fillna('')
#ag_ad_u.dropna(how='all',subset=['uname', 'aname'],inplace=True)

#ag_ad_u = pd.merge(df[['datetime','number']] ,ag_ad_u,how='right',on=['number']).fillna('')
#.sort_values(by=['idx0'])

#%%

df_x = ag_ad_u.copy()
df_print = df_x.copy()

#%%

'''保存数据'''


data_tmp = os.path.splitext(data_xlsx)[0]+"_tmp.xlsx"
df_save = df_x.copy()
df_save.columns = ut.titles_switch(df_save.columns.tolist())
df_save = ut.save_adjust_xlsx(df_save,data_tmp,width=40)

#%%
'''打印内容'''

print('邮单输出过程'.center(30, '*'))
if len(df) and flag_to_postal:
    if not os.path.exists(sheet_docx):
        input_exit('>>> 没有找到邮单模板 %s...任意键退出'%sheet_docx)
    df_ret = df_print.apply(fill_postal_save,axis = 1) # 重复处理并保存邮单 实际等于for循环

    # 计算邮单并显示
    count = len(df_ret[df_ret != ''])
    codes = df_print['number'].astype(str)
    dates = df_print['datetime'].astype(str)
    codesrange = codes.iloc[0] if codes.iloc[0] == codes.iloc[-1] else ('%s--%s'%(codes.iloc[0],codes.iloc[-1]))
    datesrange = dates.iloc[0] if dates.iloc[0] == dates.iloc[-1] else ('%s--%s'%(dates.iloc[0],dates.iloc[-1]))
    print_log('\n>>> 最终生成邮单【%s条】范围: 【%s】日期:【%s】'%(count,codesrange,datesrange))

#%%
    
input_exit('>>> !!! 全部完成,可以回顾记录...任意键退出 !!!')
