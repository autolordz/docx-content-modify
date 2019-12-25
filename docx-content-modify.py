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

#%%
import dcm_util as ut
from dcm_globalvar import *
from dcm_df_progress import df_fill_infos,df_oa_append,df_read_fix,merge_group_cases
from dcm_df_transform import df_check_format,df_transform_stream
from dcm_df_transform import make_adr,make_agent,merge_user,merge_usr_agent_adr
from dcm_df_transform import reclean_data,sort_data
from dcm_print_postal import fill_postal_save

#%%
import datetime

print('''
Postal Notes Automatically Generate App

Updated on %s

Depends on: python-docx,pandas,StyleFrame,configparser

@author: Autoz (autolordz@gmail.com)
'''%datetime.datetime.now().strftime("%Y-%m-%d %H:%M"))

#%%

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

#import sys
#sys.exit()

#%%
print('读取数据'.center(30, '*'))
df = pd.read_excel(data_xlsx,sort=False).dropna(how='all')  # 删除全空的行
df = df_read_fix(df);df # fix empty data columns

print('新增OA记录'.center(30, '*'))
df,ocodes = df_oa_append(df) # append oa data and sava # 新增OA记录

#print('合并系列案逻辑部分'.center(30, '*'))
#df = merge_group_cases(df,ocodes);df # 合并系列案逻辑部分 and save

print('填充判决书内容'.center(30, '*'))
df = df_fill_infos(df) # 填充判决书内容 # filled and save

#%% df tramsfrom stream 数据转换流程
print('数据转换流程'.center(30, '*'))
df_print = df_transform_stream(df)

#%%
print('邮单输出过程'.center(30, '*'))

if len(df) and flag_to_postal:
    if not os.path.exists(sheet_docx):
        input_exit('>>> 没有找到邮单模板 %s...任意键退出'%sheet_docx)
    df_ret = df_print.apply(fill_postal_save,axis = 1) # 重复处理并保存邮单

    # 计算邮单并显示
    count = len(df_ret[df_ret != ''])
    codes = df_print['number'].astype(str)
    dates = df_print['datetime'].astype(str)
    codesrange = codes.iloc[0] if codes.iloc[0] == codes.iloc[-1] else ('%s:%s'%(codes.iloc[0],codes.iloc[-1]))
    datesrange = dates.iloc[0] if dates.iloc[0] == dates.iloc[-1] else ('%s:%s'%(dates.iloc[0],dates.iloc[-1]))
    print_log('\n>>> 最终生成邮单【%s条】范围: 【%s】日期:【%s】'%(count,codesrange,datesrange))

input_exit('>>> 全部完成,可以回顾记录...任意键退出')
