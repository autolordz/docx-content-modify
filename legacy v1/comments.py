# -*- coding: utf-8 -*-
"""
Created on Wed Sep 11 11:46:29 2019

@author: autol
"""

#def new_adr_format(n_adr):
#    y=[]
#    for i,k in enumerate(n_adr):
#        y += [k+adr_tag+n_adr.get(k)]
#    return ('，'.join(list(filter(None, y))))


#def copy_rows_adr(x):
#    ''' copy jdocs address to address column'''
#    '''格式:['当事人','诉讼代理人','地址','new_adr','案号']'''
#    x[:3] = x[:3].astype(str)
#    user = x[0];agent = x[1];adr = x[2];n_adr = x[3];codes = x[4]
#    if not isinstance(n_adr,dict):
#        return adr
#    else:
#        y = split_list(r'[,，]',adr)
#        adr1 = y.copy()
#        for i,k in enumerate(n_adr):
#            by_agent = any([k in ag for ag in re.findall(r'[\w+、]*\/[\w+]*',agent)]) # 找到代理人格式 'XX、XX/XX_123123'
#            if by_agent and k in adr: # remove user's address when user with agent 用户有代理人就不要地址
#                y = list(filter(lambda x:not k in x,y))
#            if type(n_adr) == dict and not k in adr and k in user and not by_agent:
#                y += [k+adr_tag+n_adr.get(k)] # append address by rules 输出地址格式
#        adr2 = y.copy()
#        adr =  '，'.join(list(filter(None, y)))
#        if Counter(adr1) != Counter(adr2) and flag_check_jdocs and adr:print_log('>>> 【%s】成功复制判决书地址=>【%s】'%(codes,adr))
#    return adr



#if any(dfo['add_index'] == 'new'):
#            dfo = titles_resort(dfo,titles_main)
#            if len(add) > 0:
#                print(1121212)
#                dfo = save_adjust_xlsx(dfo,data_xlsx)
#            else:
#                print_log('>>> 内容没变,不用保存 ..')


#%%
#import os
#import pandas as pd
#from util import print_log,input_exit
#
#def init_var():
#    cfgfile = 'conf.txt'
#    try:
#        if not os.path.exists(cfgfile): write_config(cfgfile) # 生成默认配置
#        var = pd.Series(read_config(cfgfile));var
#    except Exception as e:
#        print_log('>>> 配置文件出错 %s ,删除...'%e)
#        if os.path.exists(cfgfile):
#            os.remove(cfgfile)
#        try:
#            write_config()
#            var = pd.Series(read_config());var
#        except Exception as e:
#            '''这里可以添加配置问题预判问题'''
#            input_exit('>>> 配置文件再次生成失败 %s ...'%e)
#        print(var)
#        return var


#data_xlsx = var.data_xlsx
#data_oa_xlsx = var.data_oa_xlsx
#sheet_docx = var.sheet_docx

#%%

#
#write_config()
#
#conf_list = read_config();conf_list
#
##locals().update(conf_list)
#
#var = pd.Series(conf_list)
#
#var.data_xlsx
#
#var.sheet_docx


