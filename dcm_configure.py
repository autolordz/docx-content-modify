# -*- coding: utf-8 -*-
"""
Created on Wed Sep 11 11:41:46 2019

@author: autol
"""

import configparser

#%% config and default values

def write_config(cfgfile):
    cfg = configparser.ConfigParser(allow_no_value=1,
                                    inline_comment_prefixes=('#', ';'))

    cfg['config'] = dict(
            data_xlsx = 'data_main.xlsx    # 数据模板地址',
            data_oa_xlsx = 'data_oa.xlsx    # OA数据地址',
            sheet_docx = 'sheet.docx    # 邮单模板地址',
            flag_fill_jdocs_infos = '1    # 是否填充判决书地址',
            flag_append_oa = '1    # 是否导入OA数据',
            flag_to_postal = '1    # 是否打印邮单',
            flag_check_jdocs = '0    # 是否检查用户格式,输出提示信息',
            flag_check_postal = '0    # 是否检查邮单格式,输出提示信息',
            data_case_codes = '   # 指定打印案号,可接多个,示例:AAA,BBB,优先级1',
            data_date_range = '  # 指定打印数据日期范围示例:2018-09-01:2018-12-01,优先级2',
            data_last_lines = '3    # 指定打印最后行数,优先级3',
        )

    with open(cfgfile, 'w',encoding='utf-8-sig') as configfile:
        cfg.write(configfile)
    print('>>> 重新生成配置 %s ...'%cfgfile)
    return cfg['config']


#%%
def read_config(cfgfile):
    cfg = configparser.ConfigParser(allow_no_value=True,
                                    inline_comment_prefixes=('#', ';'))
    cfg.read(cfgfile,encoding='utf-8-sig')
    ret = dict(
            data_xlsx = cfg['config']['data_xlsx'],
            data_oa_xlsx = cfg['config']['data_oa_xlsx'],
            sheet_docx = cfg['config']['sheet_docx'],
            data_case_codes = cfg['config']['data_case_codes'],
            data_date_range = cfg['config']['data_date_range'],
            data_last_lines = cfg['config']['data_last_lines'],
            flag_fill_jdocs_infos = int(cfg['config']['flag_fill_jdocs_infos']),
            flag_append_oa = int(cfg['config']['flag_append_oa']),
            flag_to_postal = int(cfg['config']['flag_to_postal']),
            flag_check_jdocs = int(cfg['config']['flag_check_jdocs']),
            flag_check_postal = int(cfg['config']['flag_check_postal']),
        )
    return ret
#    return dict(cfg.items('config'))

