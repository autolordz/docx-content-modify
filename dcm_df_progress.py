# -*- coding: utf-8 -*-
"""
Created on Wed Sep 11 15:11:39 2019

@author: autol
"""

import os,re
from glob import glob
import pandas as pd
import dcm_util as ut
from dcm_copy_infos import copy_rows_user_func
from dcm_get_jdocs import get_all_jdocs,rename_jdocs_codes
from dcm_globalvar import *

#%%

def df_oa_append(dfo):
    '''main fill OA data into df data'''

    if not flag_append_oa:
        return dfo,[]
    if not os.path.exists(data_oa_xlsx):
        print_log('>>> 没有找到OA模板 %s...不处理！！'%data_oa_xlsx)
        return dfo,[]
    dfoa = pd.read_excel(data_oa_xlsx,sort=False)[titles_oa].dropna(how='all')  # only some columns
    print('>>> OA数据共%s条'%len(dfoa))

    df0 = dfo.copy()

    if '适用程序' not in dfo.columns:
        dfo['适用程序'] = '' # 新建一栏用于系列案
    dfoa = df_make_subset(dfoa) # subset by n lines
    dfoa = df_read_fix(dfoa) # fix empty data columns

    dfoa['add_index'] = 'new'
    dfo['add_index'] = 'old'

    ec = list(set(ut.expand_codes(dfo['案号'].to_list()))) # 展开案号
    add =  dfoa[~dfoa['案号'].isin(ec)] # 新增条目
    dfn = pd.concat([dfo,add],sort=1).fillna('')
    dfn.sort_values(by=['立案日期','案号'],inplace=True)
#        dfn.drop_duplicates(['立案日期','案号'],keep='first',inplace=True)

    df_t0 = dfn[dfn['add_index'] == 'old']
    df_t1 = dfn[dfn['add_index'] == 'new']
    df_t1l = df_t1['立案日期'].to_list()
    print_log('>>> 截取OA【%s条】-Data记录 old【%s条】-new【%s条】...'%(len(dfoa),len(df_t0),len(df_t1)))
    if df_t1l:
        print_log('>>> 实际添加【%s条】【%s】共【%s】条...'%(len(df_t1),
                                               str(df_t1l[0])+':'+str(df_t1l[-1]),
                                               len(dfn)))
    save_df(df0,dfn) # 新旧对比保存
    return dfn,dfoa['原一审案号'].to_list() # 添加n条OA的原一审案号

def merge_group_cases(dfo,ocodes):

    '''合并系列案 input dfo return dfn '''

    dfn = dfo.copy() # [dfo['适用程序'].str.len()>2]
    ds = dfo[['适用程序','当事人']].drop_duplicates().copy();ds # 依据'适用程序','当事人'定性系列案
    for tag1,tag2 in zip(ds['适用程序'].to_list(),ds['当事人'].to_list()): # ds是系列案标签和内容
        if tag2:
            dgroup = dfo[dfo['当事人']==tag2] # 查找dfo拥有的系列案
        elif tag1:
            dgroup = dfo[dfo['适用程序']==tag1] # 先查 '当事人' 后查 '适用程序'
        if len(dgroup) > 1:
            ss = dgroup.iloc[0].copy() # 系列案选一个
            if done_tag not in ss['适用程序']: # 处理系列案
                sn0 = dgroup['案号'].to_list()
                sn = [re.search(r'\d+(?=号)|$',x).group(0) for x in sn0]
                if sn[0] != sn[-1]:
#                    ss['案号'] = re.sub(r'\d+(?=号)','%s-%s'%(sn[0],sn[-1]),sn0[0])
                    ss['案号'] = re.sub(r'\d+(?=号)','、'.join(sn),sn0[0])
                    print_log('>>> 发现并合并系列案：', ss['案号'])
                    if not ss['适用程序']: ss['适用程序'] = '、'.join(sn)
                    ss['适用程序'] += done_tag
                dfn = pd.concat([dfn[~dfn.isin(dgroup).all(1)], ss.to_frame().T]) #合并系列并过滤原来条目
    save_df(dfo,dfn)

    dfn = dfn[dfn['原一审案号'].isin(ocodes)]; dfn # 合并后找回记录案号

    return dfn

#%% df process steps

def fill_infos_func(dfj,dfo):
    '''填充判决书内容'''
    dfn = copy_rows_user_func(dfj,dfo)
    rename_jdocs_codes(dfn)
    return dfn


def df_read_fix(df):
    '''fix codes remove error format 处理案号格式'''
    df.rename(columns={'承办人':'主审法官'},inplace=True)
    x_col_tags = ['立案日期','案号','主审法官','当事人']
    df.dropna(how='any',subset=x_col_tags,inplace=True)
    df = df.fillna('')
    df[['案号','原一审案号']] = df[['案号','原一审案号']].applymap(ut.case_codes_fix)
    return df

def df_fill_infos(dfo):
    '''main fill jdocs infos'''
    if len(dfo) == 0:  return dfo
    docs = glob(ut.parse_subpath(jdocs_path,'*.docx')) # get jdocs
    if not docs: return dfo
    dfj = get_all_jdocs(docs)
    if len(dfj) == 0: print_log('>>> 没有找到判决书...不处理！！') ; return dfo
    dfn = dfo.copy()
    dfn = fill_infos_func(dfj,dfn)
    if flag_fill_jdocs_infos:
        save_df(dfo,dfn)
    return dfn

def save_df(df_old,df_new): # 内容相同就不管
    '''保存并对比记录'''
    try:
        df_old = ut.titles_resort(df_old,titles_main)
        df_new = ut.titles_resort(df_new,titles_main)
        pd.testing.assert_frame_equal(df_old,df_new)
        print_log('\n>>> 内容没变,不用保存 ..\n')
        return 0
    except Exception: # 不同则保存
        r = ut.save_adjust_xlsx(df_new,data_xlsx) ; print(r)
        return 1

def df_make_subset(df):

    '''
    cut orgin data into subset by conditions
    d_codes: 多个指定案号例如: （2018）哈哈1234号,（2018）哈哈3333号
    d_range: 2019-08-13：2019-08-27
    '''
    d_codes,d_range,d_lines = data_case_codes,data_date_range,0
    if data_last_lines:  d_lines = int(data_last_lines)

    ct,dats = ut.check_time(d_range);dats
#        ct,dats = check_time('2019-08-13：2019-08-27');dats
    if d_codes:
        dcc = ut.split_list(r'[,，;；]',d_codes)
        dcc = list(filter(None,[ut.case_codes_fix(x) for x in dcc]))
        df = df[df['案号'].isin(dcc) | df['原一审案号'].isin(dcc)]
#        df1 = df[df['案号'].isin(dcc) | df['原一审案号'].isin(dcc)];df1
    elif ct:
        print_log('\n>>> 预定读取【%s】'%d_range)
        df['立案日期'] = pd.to_datetime(df['立案日期'])
        df.sort_values(by=['立案日期'],inplace=True)
        try:
            x = dats[0]
            if len(dats) == 1:
                y = dats[0] #  y = str(datetime.date.today())
            else:
                y = dats[1]
            x,y = ut.parse_datetime(x),ut.parse_datetime(y)
            x1 = df['立案日期'].iloc[0].to_pydatetime()
            y1 = df['立案日期'].iloc[-1].to_pydatetime()
            t1 = min(x,y); t2 = max(x,y)
            t1 = max(t1,x1);t2 = min(t2,y1)
            date_start = t1 if t1 else x1
            date_end = t2 if t2 else y1
            df = df[(df['立案日期']>=date_start)&(df['立案日期']<=date_end)].copy() #这里数据分片有警告
            df['立案日期'] = df['立案日期'].astype(str)
        except Exception as e:
            print_log('>>> 日期异常',e)
    elif d_lines:
        df = df.tail(d_lines)
    return df



