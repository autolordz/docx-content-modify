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

# -*- coding: utf-8 -*-

#%%
import os,re,sys,datetime,configparser,shutil
import pandas as pd
from pandas import DataFrame, read_excel, merge, concat, set_option, to_datetime
isStyleFrame = 1
from StyleFrame import StyleFrame, Styler
from collections import Counter
from docx import Document
from glob import glob
set_option('max_colwidth',500)
set_option('max_rows', 50)
set_option('max_columns',50)

flag_print = 0
flag_output_log = 1

cfgfile = 'conf.txt'
logname = 'log.txt'
data_xlsx = 'data_main.xlsx'
data_oa_xlsx = 'data_oa.xlsx'
sheet_docx = 'sheet.docx'
address_tmp_xlsx = 'address_tmp.xlsx'
postal_path = os.path.join('.','postal')
jdocs_path = os.path.join('.','jdocs')
flag_fill_jdocs_infos = 1
flag_append_oa = 1
flag_to_postal = 1
flag_check_jdocs = 0
flag_check_postal = 0
data_case_codes = 'AAA,BBB'
data_date_range = '2018-09-01:2018-12-01'
data_last_lines = 10
conf_list = 0

#%% print_log log

if os.path.exists(logname):
    os.remove(logname)
def print_log(*args, **kwargs):
    print(*args, **kwargs)
    if flag_output_log:
        with open(logname, "a",encoding='utf-8') as file:
            print(*args, **kwargs, file=file)
    else:
        if os.path.exists(logname):
            os.remove(logname)

#%%
print_log('''
Postal Notes Automatically Generate App

Updated on Thu Jun 19 2019

Depends on: python-docx,pandas,StyleFrame,configparser

@author: Autoz
''')
#%% config and default values


def set_default_value(**kwargs):
    global data_date_range
    data_date_range = kwargs.get('data_date_range') if kwargs.get('data_date_range') != None else '# 2018-01-01:2018-12-01'
    
def write_config():
    cfg = configparser.ConfigParser(allow_no_value=True,
                                    inline_comment_prefixes=('#', ';'))
    cfg['config'] = {'data_xlsx': data_xlsx+'    # 数据模板地址',
                     'data_oa_xlsx': data_oa_xlsx+'    # OA数据地址',
                     'sheet_docx': sheet_docx+'    # 邮单模板地址',
                     'flag_fill_jdocs_infos': str(int(flag_fill_jdocs_infos))+'    # 是否填充判决书地址',
                     'flag_append_oa': str(int(flag_append_oa))+'    # 是否导入OA数据',
                     'flag_to_postal': str(int(flag_to_postal))+'    # 是否打印邮单',
                     'flag_check_jdocs': str(int(flag_check_jdocs))+'    # 是否检查用户格式,输出提示信息',
                     'flag_check_postal': str(int(flag_check_postal))+'    # 是否检查邮单格式,输出提示信息',
                     'flag_output_log': str(flag_output_log)+'    # 是否保存打印',
                     'data_case_codes': '   # 指定打印案号,可接多个,示例:AAA,BBB,优先级1',
                     'data_date_range': '  # 指定打印数据日期范围示例:%s,优先级2'%(data_date_range),
                     'data_last_lines': str(data_last_lines)+'    # 指定打印最后行数,优先级3',
                     }
    with open(cfgfile, 'w',encoding='utf-8-sig') as configfile:
        cfg.write(configfile)
    print_log('>>> 重新生成配置 %s ...'%cfgfile)

def read_config():
    global data_xlsx,data_oa_xlsx,sheet_docx,address_tmp_xlsx,postal_path
    global jdocs_path,data_last_lines,data_date_range,data_case_codes
    global flag_fill_jdocs_infos,flag_append_oa
    global flag_to_postal,flag_check_jdocs,flag_check_jdocs,flag_check_postal,flag_output_log
    cfg = configparser.ConfigParser(allow_no_value=True,
                                    inline_comment_prefixes=('#', ';'))
    cfg.read(cfgfile,encoding='utf-8-sig')
    data_xlsx = cfg['config']['data_xlsx']
    data_oa_xlsx = cfg['config']['data_oa_xlsx']
    sheet_docx = cfg['config']['sheet_docx']
    data_case_codes = cfg.get('config', 'data_case_codes',fallback=data_case_codes)
    data_date_range = cfg.get('config', 'data_date_range',fallback=data_date_range)
    data_last_lines = int(cfg.get('config','data_last_lines',fallback=data_last_lines))
    flag_fill_jdocs_infos = int(cfg.get('config', 'flag_fill_jdocs_infos',fallback=flag_fill_jdocs_infos))
    flag_append_oa = int(cfg.get('config', 'flag_append_oa',fallback=flag_append_oa))
    flag_to_postal = int(cfg.get('config', 'flag_to_postal',fallback=flag_to_postal))
    flag_check_jdocs = int(cfg.get('config', 'flag_check_jdocs',fallback=flag_check_jdocs))
    flag_check_postal = int(cfg.get('config', 'flag_check_postal',fallback=flag_check_postal))
    flag_output_log = int(cfg.get('config', 'flag_output_log',fallback=flag_output_log))
    return dict(cfg.items('config'))
#%% global variable

titles_cn = ['立案日期','案号','当事人','诉讼代理人','地址']
titles_en = ['datetime','number','uname','aname','address']
titles_oa = ['立案日期','案号','原一审案号','承办人','当事人','适用程序']
titles_main = ['立案日期','适用程序','案号','原一审案号','判决书源号','主审法官','当事人','诉讼代理人','地址',]

path_names_clean = re.compile(r'[^A-Za-z\u4e00-\u9fa5（）()：]') # remain only name including old name 包括括号冒号
search_names_phone = lambda x: re.search(r'[\w（）()：:]+\_\d+',x)  # phone numbers
path_code_ix = re.compile(r'[(（][0-9]+[)）].*?号') # case numbers
adr_tag = '/地址：'

#%% read func
def split_list(regex,L):
    return list(filter(None,re.split(regex,L)))

def user_to_list(u):
    '''get name list from user string
    Usage: '申请人:张xx, 被申请人:李xx, 原审被告:罗xx（又名罗aa）' 
    -> ['张xx', '李xx', '罗xx（又名罗aa）']
    '''
    u = split_list(r'[:、,，]',u)
    return [x for x in u if not re.search(r'申请人|被申请人|原告|被告|原审被告|上诉人|被上诉人|第三人|原审诉讼地位',x)]

def check_codes(x):
    return bool(re.search(path_code_ix.pattern,str(x)))

def case_codes_fix(x):
    '''fix string with chinese codes format
    Usage: 'dsfdsf(2018)中文中文248号sdfsdf' -> '（2018）中文中文248号'
    '''
    x = str(x)
    x = re.search(path_code_ix.pattern+r'|$',x).group().strip().replace(' ','')
    x = x.replace('(','（').replace(')','）')
    return x

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

def titles_trans(df_list):
    '''change titles between Chinese and English'''
    titles_cn2en = dict(zip(titles_cn, titles_en))
    titles_en2cn = dict(zip(titles_en, titles_cn))
    trans_cn_en = list(map(lambda x,y:(titles_cn2en if y else titles_en2cn).get(x),
                           df_list,list(map(check_cn_str,df_list))))
    return trans_cn_en

def titles_trans_columns(df,titles):
    '''sub-replace columns titles you want'''
    titles_rest = df.drop(titles,axis=1).columns.tolist()
    df = df[titles + titles_rest]
    df.columns = titles_trans(titles) + titles_rest
    return df

def titles_resort(df,titles):
    '''resort titles with orders'''
    titles_rest = df.drop(titles,axis=1).columns.tolist()
    return df[titles + titles_rest]

def save_adjust_xlsx(df,file='test.xlsx',textfit=('当事人', '诉讼代理人', '地址'),width=60):
    '''save and re-adjust excel format'''
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
    print_log('>>> 保存文件 => 文件名 \'%s\' => 数据保存成功...' %(file))
    return df

#%%
    
def read_jdocs_table(tables):
    codes = ''
    for table in tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    x = paragraph.text
                    if re.search(path_code_ix,x) and len(x) < 25:
                        codes = case_codes_fix(x)
                        break
    return codes

def get_jdocs_infos(doc,lines = 20):# search at least 20 lines docs
    '''get pre address from judgment docs, return docs pre code and address'''
    adrs = {};codes = ''
    try:tables = Document(doc).tables
    except Exception as e: 
        print('读取错误 %s ,docx文档问题,请重新另存为,或关闭已打开的docx文档'%e)
        return codes,adrs
    if tables: codes = read_jdocs_table(tables)
    paras = Document(doc).paragraphs
    if not paras:
        return codes,adrs
    if len(paras) > 20: # 多于20行就扫描一般内容
        lines = int(len(paras)/2)
    parass = paras[:lines]
    for i,para in enumerate(parass):
        x = para.text.strip()
        if len(x) > 150: continue # 段落大于150字就跳过
        if re.search(path_code_ix,x) and len(x) < 25:
            codes = case_codes_fix(x);continue # codes
        cond3 = re.search(r'法定代表|诉讼|代理人|判决|律师|请求|证据|辩称|辩论|不服',x) # 跳过非人员信息
        cond4 = re.search(r'上市|省略|区别|借款|保证|签订',x) # 跳过非人员信息,模糊 
        cond1 = re.search(r'(?<=[：:]).*?(?=[,，。])',x)
        cond2 = re.search(r'.*?[省市州县区乡镇村]',x)
        if cond3:continue
        if cond4:continue
        if cond1 and cond2:
            '''
            Todo: get user and address
            Usage: '被上诉人（原审被告）：张三，男，1977年7月7日出生，汉族，住XX自治区(省)XX市XX区1212。现住XX省XX市XX区3434'
            -> {'张三': 'XX省XX市XX区3434'}
            '''
            try:
                name = re.search(r'(?<=[：:]).*?(?=[,，。])|$',x).group(0).strip()
                name = re.sub(r'[(（][下称|原名|反诉|变更前].*?[）)]','',name) # filter some special names,notice here will add some words for filter
                z = split_list(r'[,，:：.。]',x)
                z = [re.sub(r'户[籍口]|居住|身份证|所在地|住所地?|住址?|^[现原]住?','',y) for y in z if re.search(r'.*?[省市州县区乡镇村]',y)][-1] # 几个地址选最后一个 remain only address
                adr = {name:''.join(z)}
                adrs.update(adr)
            except Exception as e:
                print_log('获取信息失败 =>',e)
    return codes,adrs

def rename_jdoc_x(doc,codes):
    '''rename only judgment doc files'''
    jdoc_name = os.path.join(os.path.split(doc)[0],'判决书_'+codes+'.docx')
    if not codes in doc:# os.path.exists(jdoc_name)
        try:
            os.rename(doc,jdoc_name)
            return True
        except Exception as e:
            print_log(e)
            os.remove(doc)
            return False
    return False

def get_all_jdocs(docs):
    numlist=[]; nadr = []
    for doc in docs:
        codes,adrs = get_jdocs_infos(doc)
        if codes:
            rename_jdoc_x(doc,codes)
        numlist.append(codes)
        nadr.append(adrs)
        if flag_check_jdocs and codes:
            print_log('>>> 判决书信息 【%s】-【%s人】-%s \n'%(codes,len(adrs),adrs))
    numlist = list(map(case_codes_fix,numlist))
    return DataFrame({'判决书源号':numlist,'new_adr':nadr})

#%%
def copy_rows_adr(x):
    ''' copy jdocs address to address column''' 
    '''格式:['当事人','诉讼代理人','地址','new_adr','案号']'''
    x[:3] = x[:3].astype(str)
    user = x[0];agent = x[1];adr = x[2];n_adr = x[3];codes = x[4]
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
        if Counter(adr1) != Counter(adr2) and flag_check_jdocs and adr:print_log('>>> 【%s】成功复制判决书地址=>【%s】'%(codes,adr))
    return adr

def copy_users_compare(jrow,df,errs=list('    ')):
    '''copy users and check users completement
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
    if flag_print:
        print('x=',x);print('y=',y);print('rxy=',rxy)
        print('rxyx=',rxyx);print('rxyy=',rxyy)
    if rxy == 0: # 不相交，完全无关
        return errs[0]
    if max(x.values()) > 1 or max(y.values()) > 1: # 有字段重复
        xdu = [k for k,v in x.items() if v > 1] # 重复的内容
        ydu = [k for k,v in y.items() if v > 1]
        print_log('>>> 用户有字段重复【%s】-【案件:%s】 vs 【判决书:%s】'
                  %("{0:.0%}".format(rxy),xdu,ydu))
        return errs[2]
    if rxy == 1: # 完全匹配
        return df['当事人']
    if 0 < rxy < 1: # 错别字
        dx = list((x-y).elements())
        dy = list((y-x).elements())
        xx = Counter(''.join(dx))
        yy = Counter(''.join(dy))
        rxxyy = len(list(xx&yy.keys()))/len(list(xx|yy.keys()))
        if flag_print:print('rxxyy=',rxxyy)
        if rxxyy >= .6:
            print_log('>>> 觉得有【错别字率 %s】->【案件:%s vs 判决书:%s】'
                      %("{0:.0%}".format(1-rxxyy),dx,dy))
            return errs[1]
        elif rxxyy >= .2:
            print_log('>>> 觉得不好判断当正常处理【差异率 %s】vs【相同范围:%s】->【差异范围:案件:%s vs 判决书:%s】 '
                          %("{0:.0%}".format(1-rxxyy),
                            list((x&y).elements()),
                            dx,dy))
            return df['当事人']
    if rxyx > .8:
        print_log('>>> 案件 %s人 < 判决书  %s人'%(len(x),len(y)))
        if jcode != code1:# 系列案
            print_log('>>> 觉得是【系列案,判决书人员 %s 多出地址】'%(list((y-x).elements())))
            return errs[3]
        else:
            return df['当事人']
    elif rxyy > .8:
        print_log('>>> 案件 %s人 > 判决书 %s人'%(len(x),len(y)))
        print_log('>>> 觉得有【当前案件人员 %s 缺地址】'%(list((x-y).elements())))
        return df['当事人']
    return errs[0]
    
#%%

def save_jdocs_infos(x):
    '''save remane jdocs'''
    try:
        x = x.fillna('')
        save_adjust_xlsx(x,file=address_tmp_xlsx,textfit=('判决书源号','new_adr'))
#        x.to_excel(address_tmp_xlsx,index=False)
    except Exception as e:
        print_log('%s <= 保存失败,请检查... %s'%(address_tmp_xlsx,e))
  

def new_adr_format(n_adr):
    y=[]
    for i,k in enumerate(n_adr):
        y += [k+adr_tag+n_adr.get(k)]
    return ('，'.join(list(filter(None, y))))
      
def copy_rows_user_func(dfj,dfo):
    
    def copy_rows_adr1(x,n_adr):
        ''' copy jdocs address to address column
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
            if Counter(adr1) != Counter(adr2) and flag_check_jdocs and adr:print_log('>>> 【%s】成功复制判决书地址=>【%s】'%(codes,adr))
        return adr
    
    '''copy users line regard adr user'''
    errs = ['【OA无用户记录】','【用户错别字】','【字段重复】','【系列案】']
    
    dfo['判决书源号'] = ''
    
    for (i,dfor) in dfo.iterrows():
        for (j,dfjr) in dfj.iterrows():
            code0 = str(dfor['案号']).strip()
            code1 = str(dfor['原一审案号']).strip()
            jcode = str(dfjr['判决书源号']).strip()
            n_adr = dfjr['new_adr']
            if isinstance(n_adr,dict):
                if not n_adr:continue# 提取jdocs字段失败
                if code1 == jcode:# 同案号，则找到内容
                    print_log('\n>>> 找到信息_案号=%s__源号=%s__判决书源号=%s'%(code0,code1,jcode))
                    dfo.loc[i,'地址'] = copy_rows_adr1(dfor,n_adr)
                    dfo.loc[i,'判决书源号'] = jcode
                    break
                else:#[::-1] # 没案号
                    tag1 = copy_users_compare(dfjr,dfor,errs)
                    if tag1 not in errs:
                        print_log('\n>>> 找到信息_案号=%s__源号=%s__判决书源号=%s'%(code0,code1,jcode))
                        dfo.loc[i,'地址']= copy_rows_adr1(dfor,n_adr)
                        dfo.loc[i,'判决书源号'] = jcode
                        break
                    else:
                        pass
    save_jdocs_infos(dfj)
    return dfo
    

#%%

def rename_jdocs_codes_x(d,r,old_codes):
    '''add jdoc current case codes for reference 判决书改名，包括源案号'''
    if str(r[old_codes]) in str(d):
        nd = os.path.join(os.path.split(d)[0],'判决书_'+str(r['案号']) +'_原_'+ str(r[old_codes]) + '.docx')
        if(d == nd):
            return d
        try:
            if os.path.exists(nd):
                os.remove(nd)
#            if '_原_' in d:
#                shutil.copyfile(d,nd)
            else:
                os.rename(d,nd)
                print('>>> 重命名判决书 => ',nd)
        except Exception as e:
            print_log(e)
        return nd
    return d

def rename_jdocs_codes(dfo):
    '''rename with new codes'''
    old_codes='判决书源号'
    docs = glob(parse_subpath(jdocs_path,'判决书_*.docx'))
    df = dfo[dfo[old_codes] != '']
    if docs:
        for doc in docs:
            for (i,dfr) in df.iterrows():
                if check_codes(dfr[old_codes]) and str(dfr[old_codes]) in doc:
                    rename_jdocs_codes_x(doc,dfr,old_codes)
                    break
    return None

def fill_infos_func(dfj,dfo):
    '''填充信息并处理系列案'''
    dd = dfo[['适用程序','当事人']][dfo['适用程序'].str.len()>2].drop_duplicates().copy()
    dfoo = dfo.copy()
    for tag1,tag2 in zip(dd['适用程序'].to_list(),dd['当事人'].to_list()):
        serise = dfo[(dfo['适用程序']==tag1)&(dfo['当事人']==tag2)]
        if len(serise) > 0:
            ss = serise.iloc[0].copy()
            if '_集合' not in ss['适用程序']:
                print_log('>>> 发现系列案：',serise['案号'].to_list())
                sn0 = serise['案号'].to_list()
                sn = [re.search(r'\d+(?=号)|$',x).group(0) for x in sn0]
                if sn[0] != sn[-1]:
                    ss['案号'] = re.sub(r'\d+(?=号)','%s-%s'%(sn[0],sn[-1]),sn0[0])
                ss['适用程序'] = ss['适用程序']+'_集合'
                dfoo = pd.concat([ dfoo[~dfoo.isin(serise).all(1)],
                                               ss.to_frame().T])
    dfo = dfoo
    dfo = copy_rows_user_func(dfj,dfo)
    rename_jdocs_codes(dfo)
    return dfo

#%% df process steps

def df_read_fix(df):
    '''fix codes remove error format 处理案号格式'''
    df[['立案日期','案号','主审法官','当事人']] = df[['立案日期','案号','主审法官','当事人']].replace('',float('nan'))
    df.dropna(how='any',subset=['立案日期','案号','主审法官','当事人'],inplace=True)
    df['原一审案号'] = df['原一审案号'].fillna('')
    df[['案号','原一审案号']] = df[['案号','原一审案号']].applymap(case_codes_fix)
    return df

def df_fill_infos(dfo):
    '''main fill jdocs infos'''
    if len(dfo) == 0:return dfo
    docs = glob(parse_subpath(jdocs_path,'*.docx')) # get jdocs
    if not docs:return dfo
    dfj = get_all_jdocs(docs)
    global dd
    dd = dfj
    if len(dfj) == 0:
        print_log('>>> 没有找到判决书...不处理！！');return dfo
    dfn = fill_infos_func(dfj,dfo)
    dfn = titles_resort(dfn,titles_main)
    try:
        if flag_fill_jdocs_infos:
            dfo = save_adjust_xlsx(dfn,data_xlsx)
    except PermissionError:
        print_log('>>> %s 文件已打开...填充判决书地址失败！！...请关闭并重新执行'%data_xlsx)
    return dfo

def df_make_subset(df,oa_new=0):
    '''
    cut orgin data into subset by conditions
    '''
    dcn = case_codes_fix(data_case_codes)
    date_range = data_date_range
    last_lines = data_last_lines
    if dcn:  # 多个指定案号例如: （2018）哈哈1234号,（2018）哈哈3333号
        df = df[df['案号'].isin(split_list('[,，;；]',dcn)) | df['原一审案号'].isin(split_list('[,，;；]',dcn))]
    elif ':' in date_range:
        print_log('\n>>> 预定读取【%s】'%date_range)
        df['立案日期'] = to_datetime(df['立案日期'])
        df.sort_values(by=['立案日期'],inplace=True)
        try:
            dats = date_range.split(':')
            x = parse_datetime(dats[0]);y = parse_datetime(dats[1])
            x1 = df['立案日期'].iloc[0].to_pydatetime()
            y1 = df['立案日期'].iloc[-1].to_pydatetime()
            t1 = min(x,y); t2 = max(x,y)
            t1 = max(t1,x1);t2 = min(t2,y1)
            date_start = t1 if t1 else x1
            date_end = t2 if t2 else y1
            df = df[(df['立案日期']>=date_start)&(df['立案日期']<=date_end)].copy() #这里数据分片有警告
            df['立案日期'] = df['立案日期'].astype(str)
            return df
        except Exception as e:
            print_log('>>> 日期异常',e)
    elif last_lines:
        df = df.tail(last_lines)
    return df

#%%
def df_oa_append(dfo):
    '''main fill OA data into df data and mark new add'''
    if flag_append_oa:
        if not os.path.exists(data_oa_xlsx):
            print_log('>>> 没有找到OA模板 %s...不处理！！'%data_oa_xlsx);return dfo
        dfoa = read_excel(data_oa_xlsx,sort=False)[titles_oa].fillna('') # only oa columns
        df1 = dfo.copy()
        df2 = dfoa.copy()
        
        if '适用程序' not in dfo.columns:
            dfo['适用程序'] = 0
        dfoa = df_make_subset(dfoa,oa_new=1) # subset by columns
        dfoa.rename(columns={'承办人':'主审法官'},inplace=True)
        dfoa = df_read_fix(dfoa) # fix empty data columns
        dfoa['add_index'] = 'new'
        dfo['add_index'] = 'old'
        dfors = dfo['适用程序']
        
        tags = list(dfors[dfors.str.len()>2&dfors.apply(lambda x:'Done' in x)].unique())
        tags = [t.replace('_集合','') for t in tags]
        
        for i,df2r in dfoa.iterrows():
            if df2r['适用程序'] in tags:continue
            dfo = dfo.append(df2r,sort=False)
                
        dfo.fillna('',inplace=True)
        dfo.drop_duplicates(['立案日期','案号'],keep='first',inplace=True)
        
        dfo.sort_values(by=['立案日期','案号'],inplace=True)
        df_noa = dfo[dfo['add_index'] == 'new']
        print_log('>>> 所有OA记录【%s条】...'%len(dfoa))
        print_log('>>> 原Data记录【%s条】...'%len(dfo))
        print_log('>>> 实际添加【%s条】新OA记录...'%len(df_noa))
        if len(df_noa):
            dd = str(df_noa['立案日期'].iloc[0]) +':'+df_noa['立案日期'].iloc[-1]
            print_log('>>> 实际添加【%s】'%dd)
        if any(dfo['add_index'] == 'new'):
            dfo = titles_resort(dfo,titles_main)
            try:dfo = save_adjust_xlsx(dfo,data_xlsx)
            except PermissionError:print_log('>>> %s 文件已打开...填充OA数据失败！！。。。请关闭并重新执行'%data_xlsx)
    return dfo

def df_check_format(x):
    '''check data address and agent format with check flag'''
    if x['aname']!='' and not re.search(r'[\/_]',x['aname']):
        print_log('>>> 记录\'%s\'---- 【诉讼代理人】格式 \'%s\' 不正确,如无请留空,请自行修改...'%(x['number'],x['aname']))
    if x['address']!='' and not re.search(r'\/地址[:：]',x['address']):
        print_log('>>> 记录\'%s\'---- 【地址】格式 \'%s\' 不正确,如无请留空,请自行修改...'%(x['number'],x['address']))
    return x



#%% df tramsfrom functions
def clean_rows_aname(x,names):
    '''Clean agent name for agent to match address's agent name'''
    if names:
        for name in names:
            if not check_cn_str(name):continue
            if name in x:
#                if flag_print: print('A=%s,B=%s'%(x,name))
                x = name;break
    x = re.sub(r'_.*','',x)
    x = re.sub(path_names_clean,'',x)
    return x

def clean_rows_adr(adr):
    '''clean adr format'''
    y = split_list(r'[,，]',adr)
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
    agent = merge(dd_l,dd_r,how='outer',on=['level_0','level_1']).drop(['level_1','level_2'],axis=1).fillna('')
    return agent

def merge_user(user,agent):
    '''合并后以uname为主,clean_aname是律师标识
    Returns:
       level_0       uname            aname              clean_aname
    0       44         张三          A律师_123213123                A律师
    2       44         王五       B律师_123123132123、C律师_123123   B律师
    '''
    return merge(user,agent,how='left',on=['level_0','uname']).fillna('')

def merge_usr_agent_adr(agent,adr):
    ''' clean_aname 去除nan,保留曾用名
    '''
    agent['clean_aname'].replace('',float('nan'),inplace=True)
    agent['clean_aname'] = agent['clean_aname'].fillna(agent['uname']).replace(path_names_clean,'')
    adr['clean_aname'] = adr['clean_aname'].apply(lambda x: clean_rows_aname(x,agent['clean_aname'].tolist()))
    tb = merge(agent,adr,how='outer',on=['level_0','clean_aname']).fillna('')
    tb.dropna(how='all',subset=['uname', 'aname'],inplace=True)
    return tb

def reclean_data(tb):
    tg = tb.groupby(['level_0','clean_aname','aname','address'])['uname'].apply(lambda x: '、'.join(x.astype(str))).reset_index()
    glist = tg['uname'].str.split(r'、',expand=True).stack().values.tolist()
    rest = tb[tb['uname'].isin(glist) == False]
    x = concat([rest,tg],axis=0,sort=True)
    return x

def sort_data(x,number):
    x = x[['level_0','uname','aname','address']].sort_values(by=['level_0'])
    x = merge(number,x,how='right',on=['level_0']).drop(['level_0'],axis=1).fillna('')
    return x

#%% main processing stream 主数据流程

try:
    if not os.path.exists(cfgfile):
        '''生成默认配置'''
        write_config()
    conf_list = read_config()
except Exception as e:
    print_log('>>> 配置文件出错 %s ,删除...'%e)
    if os.path.exists(cfgfile):
        os.remove(cfgfile)
    try:
        write_config()
        conf_list = read_config()
    except Exception as e:
        '''这里可以添加配置问题预判问题'''
        print_log('>>> 配置文件再次生成失败 %s ...'%e)
        set_default_value(data_date_range = '')
        
print_log('''>>> 正在处理...
    主表路径 = %s
    指定案件 = %s
    指定日期 = %s
    指定条数 = %s
    '''%(os.path.abspath(data_xlsx),
        conf_list.get('data_case_codes'),
        conf_list.get('data_date_range'),
        conf_list.get('data_last_lines'),
        )
    )
    

if not os.path.exists(data_xlsx):
    save_adjust_xlsx(DataFrame(columns=titles_main),data_xlsx,width=40)
    print_log('>>> %s 记录文件不存在...重新生成'%(data_xlsx))

dfo = read_excel(data_xlsx,sort=False).fillna('') #真正读取记录位置
dfo = df_read_fix(dfo) # fix empty data columns
dfo = df_oa_append(dfo) # append oa data

dfo = df_fill_infos(dfo) # fill jdocs infos
dfo = df_make_subset(dfo)
df = titles_trans_columns(dfo,titles_cn) # 中译英方便后面处理

if flag_check_postal:
    df.apply(lambda x:df_check_format(x), axis=1)
    
print_log('\n>>> ***将要打印Data记录【---%s条----】...'%len(df))

if 0<len(df)<10:
    print_log('>>> ***将要打印 => %s '%df['number'].to_list())

#%% df tramsfrom stream 数据转换流程

if len(df) and flag_to_postal:  
    try:
        print_log('\n>>> 开始生成新数据 data_main_temp... ')
        '''获取 datetime|number'''
        number = df[titles_en[:2]]
        number = number.reset_index()
        number.columns.values[0] = 'level_0'
        # user
        '''获取所有用户名包括曾用名'''
        user = df['uname']
        user = user[user != '']
        user = user.str.strip().str.split(r'[,，。]',expand=True).stack() # divide user
        user = user.str.strip().str.split(r'[:]',expand=True)# divide character
        user = user[1].str.strip().str.split(r'[、]',expand=True).stack().to_frame(name = 'uname')
        user = user.reset_index().drop(['level_1','level_2'],axis=1)
        # agent and address
        agent_adr = df[['aname','address']]
        opt = agent_adr.any()
        agent = df['aname']
        adr = df['address']

        if all(opt):
            print_log('>>> 有【代理人】和【地址】...正在处理...')
            adr = make_adr(adr,fix_aname=user['uname'].tolist())
            agent = make_agent(agent,fix_aname=adr['clean_aname'].tolist()) #获取代理人
            usr_agent = merge_user(user,agent)
            df_x = reclean_data(merge_usr_agent_adr(usr_agent,adr))
            df_x = sort_data(df_x,number)
        elif opt.address:
            print_log('>>> 只有【地址】...正在处理...')
            adr = make_adr(adr,fix_aname=user['uname'].tolist())
            adr['uname'] = adr['clean_aname']
            adr = merge_user(user,adr)
            adr = adr.assign(aname='')
            df_x = reclean_data(adr)
            df_x = sort_data(df_x,number)
        elif opt.aname:
            print_log('>>> 只有【代理人】...正在处理...')
            agent = make_agent(agent)
            agent = merge_user(user,agent)
            agent = agent.assign(address='')
            df_x = reclean_data(agent)
            df_x = sort_data(df_x,number)
        else:
            print_log('>>> 缺失【代理人】和【地址】...正在处理...')
            agent_adr.index.name = 'level_0'
            agent_adr.reset_index(inplace=True)
            df_x = merge(user,agent_adr,how='left',on=['level_0']).fillna('')
            df_x = sort_data(df_x,number)

        if len(df_x):
            data_tmp = os.path.splitext(data_xlsx)[0]+"_temp.xlsx"
            df_save = df_x.copy()
            df_save.columns = titles_trans(df_save.columns.tolist())
            try:df_save = save_adjust_xlsx(df_save,data_tmp,width=40)
            except PermissionError: print_log('>>> %s 文件已打开...请手动关闭并重新执行...保存失败'%data_tmp)

        
    except Exception as e:
        raise e
        print_log('>>> 错误 \'%s\' 生成数据失败,请检查源 \'%s\' 文件...退出...'%(e,data_xlsx));sys.exit()

#%% generate postal sheets 生成邮单流程

def re_write_text(x):
    '''re-write postal sheet content from df rows'''

    doc = Document(sheet_docx)
    doc.styles['Normal'].font.bold = True
    uname = str(x['uname']);aname = str(x['aname'])
    agent_text = aname if aname else uname
    user_text = '' if uname in agent_text else '代 '+ uname
    number_text = str(x['number'])
    address_text = str(x['address'])

    try:
        para = doc.paragraphs[9]  # No.9 line is agent name
        text = re.sub(r'[\w（）()]+',agent_text,para.text)
        para.clear().add_run(text)

        para = doc.paragraphs[11]  # No.11 line is user name
        text = re.sub(r'代 \w+',user_text,para.text)
        para.clear().add_run(text)

        para = doc.paragraphs[13]  # No.13 line is number and address
        text = re.sub(path_code_ix,number_text,para.text)
        para.clear().add_run(text)
        text = re.sub(r'(?<=\s)\w+市.*',address_text,para.text)
        para.clear().add_run(text)
    except Exception as e:
        print_log('错误 \'%s\' 替换文本 => \'%s\' 失败！！！' %(e,para.text))

    sheet_file = number_text+'_'+agent_text+'_'+user_text+'_'+address_text+'.docx'
    sheet_file = re.sub(r'[\/\\\:\*\?\"\<\>]',' ',sheet_file) # keep rename legal

    if os.path.exists(parse_subpath(postal_path,sheet_file)):
        if flag_check_postal:print_log('>>> 邮单已存在！！！ <= %s'%sheet_file)
        return ''

    if not agent_text:
        if flag_check_postal:print_log('>>> 【代理人】暂缺！！！ <= %s'%sheet_file)
        return ''
    
    if not address_text:
        if flag_check_postal:print_log('>>> 【地址】暂缺！！！ <= %s'%sheet_file)
        return ''
    try:
        doc.save(parse_subpath(postal_path,sheet_file))
        print_log('>>> 已生成邮单 => %s'%sheet_file)
        return sheet_file
    except Exception as e:
        print_log('>>> 生成失败！！！ => %s'%e)
    return ''

if len(df) and flag_to_postal:
    print_log('\n>>> 正在输出邮单...')
    if not os.path.exists(sheet_docx):
        input('>>> 没有找到邮单模板 %s...任意键退出'%sheet_docx);sys.exit()
    df_p = df_x.apply(re_write_text,axis = 1)
    count = len(df_p[df_p != ''])
    codes = df_x['number'].astype(str)
    dates = df_x['datetime'].astype(str)
    codesrange = codes.iloc[0] if codes.iloc[0] == codes.iloc[-1] else ('%s:%s'%(codes.iloc[0],codes.iloc[-1]))
    datesrange = dates.iloc[0] if dates.iloc[0] == dates.iloc[-1] else ('%s:%s'%(dates.iloc[0],dates.iloc[-1]))
    print_log('>>> 最终生成邮单【%s条】范围: 【%s】日期:【%s】'%(count,codesrange,datesrange))
    
    del df_x,df_p,codes,dates
    del user,number,agent,adr,df,agent_adr,opt
    
#%% main finish 结束所有

print_log('>>> 全部完成,可以回顾记录...任意键退出')
#input('>>> 全部完成,可以回顾记录...任意键退出');sys.exit()



