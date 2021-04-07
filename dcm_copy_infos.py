# -*- coding: utf-8 -*-
"""
Created on Wed Sep 11 12:08:08 2019

@author: autol
"""


#%%

import re
from collections import Counter
from dcm_util import split_list,user_to_list,save_adjust_xlsx
from dcm_globalvar import *

#%%

def copy_users_compare(jrow,df,errs=list('    ')):
    '''
    对比OA和data的用户记录来填充信息
    errs=['【OA无用户记录】','【用户错别字】','【字段重复】','【系列案】']
    如下对比：
    不相交，OA无用户记录
    判断字段重复,输出重复的内容
    比例确定怀疑用户错别字，判别不了直接正常输出
    判决书多于当前案件,认为是系列案
    判决书少于当前案件,当前案件缺部分地址
    '''

    code0 = str(df['案号']).strip()
    code1 = str(df['原一审案号']).strip()
    jcode = str(jrow['判决书源号']).strip()
    x = Counter(user_to_list(df['当事人'])) # 当前案件
    y = Counter(list(jrow['new_adr'].keys())) # 判决书
    rxy = len(list((x&y).elements()))/len(list((x|y).elements()))
    rxyx = len(list((x&y).elements()))/len(list(x.elements()))
    rxyy = len(list((x&y).elements()))/len(list(y.elements()))
#    print('x=',x);print('y=',y);print('rxy=',rxy)
#    print('rxyx=',rxyx);print('rxyy=',rxyy)
    if rxy == 0: # 不相交，完全无关
        return errs[0]
    if max(x.values()) > 1 or max(y.values()) > 1: # 有字段重复
        xdu = [k for k,v in x.items() if v > 1] # 重复的内容
        ydu = [k for k,v in y.items() if v > 1]
        print_log('>>> %s 用户有字段重复【%s】-【案件:%s】 vs 【判决书:%s】'
                  %(code0,'{0:.0%}'.format(rxy),xdu,ydu))
        return errs[2]
    if rxy == 1: # 完全匹配
        return df['当事人']
    if 0 < rxy < 1: # 错别字
        dx = list((x-y).elements())
        dy = list((y-x).elements())
        xx = Counter(''.join(dx))
        yy = Counter(''.join(dy))
        rxxyy = len(list(xx&yy.keys()))/len(list(xx|yy.keys()))
#        print('rxxyy=',rxxyy)
        if rxxyy >= .6:
            print_log('>>> %s 认为【错别字率 %s】->【案件:%s vs 判决书:%s】'
                      %(code0,'{0:.0%}'.format(1-rxxyy),dx,dy))
            return errs[1]
        elif rxxyy >= .2:
            print_log('>>> %s 认为【不好判断当正常处理【差异率 %s】vs【相同范围:%s】->【差异范围:案件:%s vs 判决书:%s】 '
                          %(code0,'{0:.0%}'.format(1-rxxyy),
                            list((x&y).elements()),
                            dx,dy))
            return df['当事人']
    if rxyx > .8:
        print_log('>>> %s 案件 %s人 < 判决书  %s人'%(code0,len(x),len(y)))
        if jcode != code1:# 系列案
            print_log('>>> %s 认为【系列案,判决书人员 %s 多出地址】'%(code0,list((y-x).elements())))
            return errs[3]
        else:
            return df['当事人']
    elif rxyy > .8:
        print_log('>>> %s 案件 %s人 > 判决书 %s人'%(code0,len(x),len(y)))
        print_log('>>> %s 认为【当前案件人员 %s 缺地址】'%(code0,list((x-y).elements())))
        return df['当事人']
    return errs[0]


def copy_rows_adr1(x,n_adr):
    '''
        复制判决书内容到地址栏
        格式:['当事人','诉讼代理人','地址','new_adr','案号']
        同时排除已有代理人的信息
    '''
    user = x['当事人'];agent = x['诉讼代理人'];adr = x['地址']; codes = x['案号']
    if not isinstance(n_adr,dict):
        return adr
    else:
        y = split_list(r'[,，]',adr)
        adr1 = y.copy()
        for i,k in enumerate(n_adr):
            by_agent = any([k in ag for ag in re.findall(r'[\w+、]*\/[\w+]*',agent)]) # 找到代理人格式 'XX、XX/XX_123123'
            if by_agent and k in adr: # remove user's address when user with agent 用户有代理人就不要地址
                y = list(filter(lambda x:not k in x,y))
            if type(n_adr) == dict and not k in adr and k in user and not by_agent:
                y += [k+adr_tag+n_adr.get(k)] # append address by rules 输出地址格式
        adr2 = y.copy()
        adr =  '，'.join(list(filter(None, y)))
        if Counter(adr1) != Counter(adr2) and adr and flag_check_jdocs:
            print_log('>>> 【%s】成功复制判决书地址=>【%s】'%(codes,adr))
    return adr

def copy_rows_user_func(dfj,dfo):

    '''
    根据地址用户，复制每行用户信息
    '''
    errs = ['【OA无用户记录】','【用户错别字】','【字段重复】','【系列案】']

    dfo['判决书源号'] = ''

    def find_source():
        print_log('\n>>> 判决书信息 | 案号=%s | 源号=%s | 判决书源号=%s'%(code0,code1,jcode))
        dfo.loc[i,'地址'] = copy_rows_adr1(dfor,n_adr)
        dfo.loc[i,'判决书源号'] = jcode

    for (i,dfor) in dfo.iterrows():
        for (j,dfjr) in dfj.iterrows():
            code0 = str(dfor['案号']).strip()
            code1 = str(dfor['原一审案号']).strip()
            jcode = str(dfjr['判决书源号']).strip()
            n_adr = dfjr['new_adr']
            if isinstance(n_adr,dict):
                if not n_adr:continue# 提取jdocs字段失败
                if code1 == jcode:# 同案号，则找到内容
                    find_source() ; break
                else:#[::-1] # 没案号
                    tag1 = copy_users_compare(dfjr,dfor,errs)
                    if tag1 not in errs:
                        find_source() ; break
                    else: pass
    dfj = dfj.fillna('')
    save_adjust_xlsx(dfj,'address_tmp.xlsx',textfit=('判决书源号','new_adr')) # 保存临时提取信息
    return dfo