# -*- coding: utf-8 -*-
"""
Created on Wed Sep 11 17:26:56 2019

@author: autol
"""
import re
import pandas as pd
import util as ut
from globalvar import *

#%% df tramsfrom functions
def clean_rows_aname(x,names):
    '''Clean agent name for agent to match address's agent name'''
    if names:
        for name in names:
            if not ut.check_cn_str(name):continue # 非中文名跳过
            if name in x:
                x = name;break
    x = re.sub(r'_.*','',x)
    x = re.sub(path_names_clean,'',x)
    return x

def clean_rows_adr(adr):
    '''clean adr format'''
    y = ut.split_list(r'[,，]',adr)
    if y:
        y = list(map(lambda x: x if re.search(r'\/地址[:：]',x) else adr_tag + x,y))
        adr = '，'.join(list(filter(None, y)))
    return adr

def make_adr(adr,fix_aname=[]):
    '''
    clean_aname:合并标识,此处如果没律师，则代理人就是自己
    fix_aname:修正名字错误
    Returns:
          level_0       address        clean_aname
    0       44      XX市XX镇XXX村          张三
    1       44      XXX市XX区XXX          B律师
    '''
    adr = adr[adr != '']
    adr = adr.str.strip().str.split(r'[,，。]',expand=True).stack()
    adr = adr.str.strip().apply(lambda x:clean_rows_adr(x))
    adr = adr.str.strip().str.split(r'\/地址[:：]',expand=True).fillna('')
    adr.columns = ['aname','address']
    adr['clean_aname'] = adr['aname'].str.strip().apply(lambda x:clean_rows_aname(x,fix_aname)) # clean adr
    adr = adr.reset_index().drop(['level_1','aname'],axis=1)
    return adr

def make_agent(agent,fix_aname=[]):
    '''
    fix_aname:修正名字错误,假如律师(aname)有多个,则选择第一个律师作为合并标识(clean_aname)，注意没有律师的合并就是自己(uname)做代理人
    Returns:
       level_0       uname            aname              clean_aname
    0       44         张三          A律师_123213123                A律师
    1       44         李四
    2       44         王五       B律师_123123132123、C律师_123123   B律师
    '''
    agent = agent[agent != '']
    agent = agent.str.strip().str.split(r'[,，。]',expand=True).stack() #Series
    agent = agent.str.strip().str.split(r'\/',expand=True).fillna('') #DataFrame
    agent.columns = ['uname','aname']
    agent['clean_aname'] = agent['aname'].str.strip().apply(lambda x: clean_rows_aname(x,fix_aname))
    dd_l = agent['uname'].str.strip().str.split(r'、',expand=True).stack().to_frame(name = 'uname').reset_index()
    dd_r = agent[agent.columns.difference(['uname'])].reset_index()
    agent = pd.merge(dd_l,dd_r,how='outer',on=['level_0','level_1']).drop(['level_1','level_2'],axis=1).fillna('')
    return agent

def merge_user(user,agent):
    '''合并后以uname为主,clean_aname是律师标识
    Returns:
       level_0       uname            aname              clean_aname
    0       44         张三          A律师_123213123                A律师
    2       44         王五       B律师_123123132123、C律师_123123   B律师
    '''
    return pd.merge(user,agent,how='left',on=['level_0','uname']).fillna('')

def merge_usr_agent_adr(agent,adr):
    ''' clean_aname 去除nan,保留曾用名'''

    agent['clean_aname'].replace('',float('nan'),inplace=True)
    agent['clean_aname'] = agent['clean_aname'].fillna(agent['uname']).replace(path_names_clean,'')
    adr['clean_aname'] = adr['clean_aname'].apply(lambda x: clean_rows_aname(x,agent['clean_aname'].tolist()))
    tb = pd.merge(agent,adr,how='outer',on=['level_0','clean_aname']).fillna('')
    tb.dropna(how='all',subset=['uname', 'aname'],inplace=True)
    return tb

def reclean_data(tb):
    tg = tb.groupby(['level_0','clean_aname','aname','address'])['uname'].apply(lambda x: '、'.join(x.astype(str))).reset_index()
    glist = tg['uname'].str.split(r'、',expand=True).stack().values.tolist()
    rest = tb[tb['uname'].isin(glist) == False]
    x = pd.concat([rest,tg],axis=0,sort=True)
    return x

def sort_data(x,number):
    x = x[['level_0','uname','aname','address']].sort_values(by=['level_0'])
    x = pd.merge(number,x,how='right',on=['level_0']).drop(['level_0'],axis=1).fillna('')
    return x

def df_check_format(x):
    '''check data address and agent format with check flag'''
    if x['aname']!='' and not re.search(r'[\/_]',x['aname']):
        ut.print_log('>>> 记录\'%s\'---- 【诉讼代理人】格式 \'%s\' 不正确,如无请留空,请自行修改...'%(x['number'],x['aname']))
    if x['address']!='' and not re.search(r'\/地址[:：]',x['address']):
        ut.print_log('>>> 记录\'%s\'---- 【地址】格式 \'%s\' 不正确,如无请留空,请自行修改...'%(x['number'],x['address']))
    return x