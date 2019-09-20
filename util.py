# -*- coding: utf-8 -*-
"""
Created on Wed Sep 11 11:29:47 2019

@author: autol
"""

import os,re,datetime
from StyleFrame import StyleFrame, Styler
from globalvar import *

#%% base utils

def split_list(regex,L):
    return list(filter(None,re.split(regex,L)))

def user_to_list(u):
    '''get name list from user string
    Usage: '申请人:张xx, 被申请人:李xx, 原审被告:罗xx（又名罗aa）'
    -> ['张xx', '李xx', '罗xx（又名罗aa）']
    '''
    u = split_list(r'[:、,，]',u)
    return [x for x in u if not re.search(usrtag,x)]

def check_codes(x):
    '''check cases codes here'''
    return bool(re.search(path_code_ix.pattern,str(x)))

def case_codes_fix(x):
    '''fix string with chinese codes format
    Usage: 'dsfdsf(2018)中文中文248号sdfsdf' -> '（2018）中文中文248号'
    '''
    x = str(x)
    x = re.search(path_code_ix.pattern+r'|$',x).group().strip().replace(' ','')
    x = x.replace('(','（').replace(')','）')
    return x

def expand_codes(xxx):
    '''
    对于系列案号处理，展开案号
    Usage: ['(2018)中文中文111、248号','(2018)中文中文333、444号']
    -> ['(2018)中文中文111号', '(2018)中文中文248号', '(2018)中文中文333号', '(2018)中文中文444号'] '''
    cc =[]
    for xx in xxx:
        aa = re.split('、',re.search(r'\d+、.*\d+(?=号)|$',xx).group(0))
        bb = [re.sub(r'\d+、.*\d+(?=号)',x,xx) for x in aa]
        cc+=bb
    return cc

def parse_subpath(path,file):
    '''make subpath'''
    if not os.path.exists(path):
        os.mkdir(path)
    return os.path.join(path,file)

def check_cn_str(x):
    '''check if string contain chinese'''
    return bool(re.search(r'[\u4e00-\u9fa5]',str(x)))

def parse_datetime(date):
    '''datetime transform'''
    try:date = datetime.datetime.strptime(date,'%Y-%m-%d')
    except ValueError:print_log('时间范围格式有误,默认选取全部日期');date = ''
    return date

def titles_switch(df_list):
    '''switch titles between Chinese and English'''
    titles_cn2en = dict(zip(titles_cn, titles_en))
    titles_en2cn = dict(zip(titles_en, titles_cn))
    trans_cn_en = list(map(lambda x,y:(titles_cn2en if y else titles_en2cn).get(x),
                           df_list,list(map(check_cn_str,df_list))))
    return trans_cn_en

def titles_trans_columns(df,titles):
    '''sub-replace columns titles you want'''
    titles_rest = df.drop(titles,axis=1).columns.tolist()
    df = df[titles + titles_rest]
    df.columns = titles_switch(titles) + titles_rest
    return df

def titles_resort(df,titles):
    '''resort titles with orders'''
    titles_rest = df.drop(titles,axis=1).columns.tolist()
    return df[titles + titles_rest]

#%% read func

isStyleFrame = 1

def save_adjust_xlsx(df,file,textfit=('当事人', '诉讼代理人', '地址'),width=60):
    '''save and re-adjust excel format
    with StyleFrame or not
    '''
    try:
        print_log('>>> 保存文件 => 文件名 \'%s\''%file)
        df = df.reset_index(drop='index').fillna('')
        if isStyleFrame:
            StyleFrame.A_FACTOR = 5
            StyleFrame.P_FACTOR = 1.2
            sf = StyleFrame(df,Styler(wrap_text = False, shrink_to_fit=True, font_size= 12))
            if('add_index' in df.columns.tolist()):
                sf.apply_style_by_indexes(indexes_to_style=sf[sf['add_index'] == 'new'],
                                          styler_obj=Styler(bg_color='yellow'),
                                          overwrite_default_style=False)
                sf.apply_column_style(cols_to_style = textfit,
                                      width = width,
                                      styler_obj=Styler(wrap_text=False,shrink_to_fit=True))
            else:
                sf.set_column_width_dict(col_width_dict={textfit: width})
            if len(df):
                sf.to_excel(file,best_fit=sf.data_df.columns.difference(textfit).tolist()).save()
            else:
                sf.to_excel(file).save()
        else:
            df.to_excel(file,index=0)
    except PermissionError:
        print_log('！！！！！%s被占用，不能覆盖记录！！！！！'%file)
    return df

def check_time(dlist):
    '''split and check configure times'''
    if dlist:
        if isinstance(dlist,str):
            if re.search(r'[:：]',dlist):
                dlist = split_list(r'[:：]',dlist)
            else:
                dlist = [dlist]
        for date in dlist:
            try:
                datetime.datetime.strptime(date, '%Y-%m-%d')
            except ValueError as e:
                print("Incorrect data format, should be YYYY-MM-DD",e)
                return 0,None
        return 1,dlist
    return 0,None
