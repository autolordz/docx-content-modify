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
from pandas import DataFrame, read_excel, merge, concat, set_option, to_datetime
isStyleFrame = 0
#from StyleFrame import StyleFrame, Styler
from collections import Counter
from docx import Document
from glob import glob
set_option('max_colwidth',500)
set_option('max_rows', 50)
set_option('max_columns',50)

#%% print_log log

flag_print = 0
flag_output_log = 1
        
cfgfile = 'conf.txt'
logname = 'log.txt'

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
Auto generate word:docx from excel:xslx file.

Updated on Thu Nov 7 2018

Depends on: python-docx,pandas.

@author: Autoz
''')
#%% config and default values

data_xlsx = 'data_main.xlsx'
data_oa_xlsx = 'data_oa.xlsx'
sheet_docx = 'sheet.docx'
address_tmp_xlsx = 'address_tmp.xlsx'
postal_path = os.path.join('.','postal')
jdocs_path = os.path.join('.','jdocs')
flag_rename_jdocs = 1
flag_fill_jdocs_infos = 1
flag_append_oa = 1
flag_to_postal = 1
flag_check_jdocs = 0
flag_check_postal = 0
data_case_codes = 'AAA,BBB'
data_date_range = '2018-09-01:2018-12-01'
data_last_lines = 10

def set_default_value(**kwargs):
    global data_date_range
    data_date_range = kwargs.get('data_date_range') if kwargs.get('data_date_range') != None else '# 2018-01-01:2018-12-01'
    
def write_config():
    cfg = configparser.ConfigParser(allow_no_value=True,
                                    inline_comment_prefixes=('#', ';'))
    cfg['config'] = {'data_xlsx': data_xlsx+'    # 数据模板地址',
                     'data_oa_xlsx': data_oa_xlsx+'    # OA数据地址',
                     'sheet_docx': sheet_docx+'    # 邮单模板地址',
                     'flag_rename_jdocs': str(int(flag_rename_jdocs))+'    # 是否重命名判决书',
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
    global flag_rename_jdocs,flag_fill_jdocs_infos,flag_append_oa
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
    flag_rename_jdocs = int(cfg.get('config', 'flag_rename_jdocs',fallback=flag_rename_jdocs))
    flag_fill_jdocs_infos = int(cfg.get('config', 'flag_fill_jdocs_infos',fallback=flag_fill_jdocs_infos))
    flag_append_oa = int(cfg.get('config', 'flag_append_oa',fallback=flag_append_oa))
    flag_to_postal = int(cfg.get('config', 'flag_to_postal',fallback=flag_to_postal))
    flag_check_jdocs = int(cfg.get('config', 'flag_check_jdocs',fallback=flag_check_jdocs))
    flag_check_postal = int(cfg.get('config', 'flag_check_postal',fallback=flag_check_postal))
    flag_output_log = int(cfg.get('config', 'flag_output_log',fallback=flag_output_log))

#%% global variable

titles_cn = ['立案日期','案号','原一审案号','主审法官','当事人','诉讼代理人','地址']
titles_en = ['datetime','number','prenumber','judge','uname','aname','address']
titles_oa = ['立案日期','案号','原一审案号','承办人','当事人']

path_names_clean = re.compile(r'[^A-Za-z\u4e00-\u9fa5（）()：]') # remain only name including old name 包括括号冒号
search_names_phone = lambda x: re.search(r'[\w（）()：:]+\_\d+',x)  # phone numbers
path_code_ix = re.compile(r'[(（][0-9]+[)）].*?号') # case numbers
adr_tag = '/地址：'

#%% read func
def user_to_list(u):
    '''get name list from user string
    Usage: '申请人:张xx, 被申请人:李xx, 原审被告:罗xx（又名罗aa）' 
    -> ['张xx', '李xx', '罗xx（又名罗aa）']
    '''
    u = re.split(r'[:、,，]',u)
    return [x for x in u if not re.search(r'申请人|被申请人|原告|被告|原审被告|上诉人|被上诉人|第三人|原审诉讼地位',x)]

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

def save_adjust_xlsx(df,file='test.xlsx',width=60):
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
            sf.apply_column_style(cols_to_style=['当事人', '诉讼代理人', '地址'],
                                  width = width,
                                  styler_obj=Styler(wrap_text=False,shrink_to_fit=True))
        else:
            sf.set_column_width_dict(col_width_dict={('当事人', '诉讼代理人', '地址'): width})
        if len(df):
            sf.to_excel(file,best_fit=sf.data_df.columns.difference(['当事人', '诉讼代理人', '地址']).tolist()).save()
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
    if len(paras) > 20:lines = int(len(paras)/2)
    parass = paras[:lines]
    for i,para in enumerate(parass):
        x = para.text.strip()
        if len(x) > 150: continue # 段落大于150字就跳过
        if re.search(path_code_ix,x) and len(x) < 25:
            codes = case_codes_fix(x);continue # codes
        if re.search(r'法定代表|诉讼|代理|律师|请求|证据|辩|不服',x): continue # 跳过代理人或者没地址的人员
        if re.search(r'(?<=[：:]).*?(?=[,，。])',x) and re.search(r'(户[籍口]|居住|所在地?|住所地?|住址?|原住|现住?).*?[省市州县区乡镇村]',x):
            '''
            Todo: get user and address
            Usage: '被上诉人（原审被告）：张三，男，1977年7月7日出生，汉族，住XX自治区(省)XX市XX区1212。现住XX省XX市XX区3434'
            -> {'张三': 'XX省XX市XX区3434'}
            '''
            try:
                name = re.search(r'(?<=[：:]).*?(?=[,，。])|$',x).group(0).strip()
                name = re.sub(r'[(（][下称|原名|反诉|变更前].*?[）)]','',name) # filter some special names,notice here will add some words for filter
                z = re.split(r'[,，:：.。]',x)
                z = [re.sub(r'户[籍口]|居住|身份证|所在地|住所地?|住址?|^[现原]住?','',y) for y in z if re.search(r'.*?[省市州县区乡镇村]',y)][-1] # 几个地址选最后一个 remain only address
                adr = {name:''.join(z)}
                adrs.update(adr)
            except Exception as e:
                print_log('获取信息失败 =>',e)
    if flag_check_jdocs:print_log('>>> 获取判决书【%s】【%s人】%s'%(codes,len(adrs),list(adrs.keys())))
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
    numlist=[]; nadr = [];jcount = 0
    for doc in docs:
        codes,adrs = get_jdocs_infos(doc)
        if codes and flag_rename_jdocs:
            rename_jdoc_x(doc,codes)
            jcount += 1
        numlist.append(codes)
        nadr.append(adrs)
    numlist = list(map(case_codes_fix,numlist))
    x = DataFrame({'原一审案号':numlist,'new_adr':nadr})
    if flag_check_jdocs:print_log('>>> 获取判决书信息【%s】条'%jcount)
    return x

#%%
def copy_rows_adr(x):
    ''' copy jdocs address to address column''' 
    '''格式:['当事人','诉讼代理人','地址','new_adr','案号']'''
    x[:3] = x[:3].astype(str)
    user = x[0];agent = x[1];adr = x[2];n_adr = x[3];codes = x[4]
    if not isinstance(n_adr,dict):
        return adr
    else:
        y = re.split(r'[,，]',adr);adr1 = y.copy()
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

def copy_rows_agent(x):
    '''copy example phone-numbers to agent column if empty''' 
    '''格式:['当事人','诉讼代理人','地址','new_adr','案号']'''
    x = x.astype(str)
    user = x[0];agent = x[1];adr = x[2];codes = x[4]
    users = user_to_list(user)
    y = re.split(r'[,，]',agent);agent1 = y.copy()
    for i,u in enumerate(users):
        if not u in agent and u in adr:
            y += [u+'_12345678910'] # fake tel
    agent2 = y.copy()
    agent = '，'.join(list(filter(None, y)))
    if Counter(agent1) != Counter(agent2) and flag_check_jdocs and agent:print_log('>>> 【%s】成功复制伪手机=>【%s】'%(codes,agent))
    return agent

def copy_users_compare(jrow,df,errs=list('    ')):
    '''copy users and check users completement
    errs:['【OA无用户记录】','【怀疑用户错别字】','【用户字段重复】','【相似】','【系列】']
    如下对比：
    不相交，OA无用户记录
    判断字段重复,输出重复的内容
    比例确定怀疑用户错别字，判别不了直接正常输出
    判决书多于当前案件,认为是系列案
    判决书少于当前案件,当前案件缺部分地址
    '''
    x = Counter(list(jrow['new_adr'].keys())) # 判决书
    y = Counter(user_to_list(df['当事人'])) # 当前案件
    rxy = len(list((x&y).elements()))/len(list((x|y).elements()))
    rxyx = len(list((x&y).elements()))/len(list(x.elements()))
    rxyy = len(list((x&y).elements()))/len(list(y.elements()))
    if flag_print:
        print('x=',x);print('y=',y);print('rxy=',rxy)
    if rxy == 0: # 不相交，完全无关
        return errs[0]
    if max(x.values()) > 1 or max(y.values()) > 1: # 有字段重复
        xdu = [k for k,v in x.items() if v > 1] # 重复的内容
        ydu = [k for k,v in y.items() if v > 1]
        print_log('>>> 判决书或OA【用户重复】【%s】【判决书:%s:%s】-【OA或Data:%s:%s】'
                  %("{0:.0%}".format(rxy),jrow['原一审案号'],xdu,df['原一审案号'],ydu))
        return errs[2]
    if rxy == 1: # 完全匹配
        return df['当事人']
    if 0 < rxy < 1: # 错别字
        dx = list((x-y).elements())
        dy = list((y-x).elements())
        xx = Counter(''.join(dx))
        yy = Counter(''.join(dy))
        rxxyy = len(list((xx&yy).elements()))/len(list((xx|yy).elements()))
        if rxxyy >= .5:
            print_log('>>> 觉得有【错别字率 %s】->【判决书:%s:%s vs OA或Data:%s:%s】'
                      %("{0:.0%}".format(1-rxxyy),jrow['原一审案号'],dx,df['原一审案号'],dy))
            return errs[1]
        if rxxyy >= .2:
            print_log('>>> 觉得有【差异率 %s】vs【相同人员:%s】不好判断 ->【差异人员:判决书:%s:%s vs OA或Data:%s:%s】'
                          %("{0:.0%}".format(1-rxxyy),list((x&y).elements()),jrow['原一审案号'],dx,df['原一审案号'],dy))
            return df['当事人']
    if rxyy > .8: # 判决书>当前案件
        if str(jrow['原一审案号']) != str(df['原一审案号']):# 系列案
            if flag_check_jdocs:print_log('>>> 觉得是【系列案, 判决书人员 %s 有地址】->【判决书:%s vs OA或Data:%s】'
                                          %(list((x&y).elements()),jrow['原一审案号'],df['原一审案号']))
            return errs[4]
        else:return df['当事人']
    if rxyx > .8: # 判决书<当前案件
        print_log('>>> 觉得有【判决书 < 当前案件, 当前案件人员 %s 缺地址】-【判决书:%s vs OA或Data:%s】'
                  %(list((x&y).elements()),jrow['原一审案号'],df['原一审案号']))
        return df['当事人']
    return errs[0]
    
#%%

def save_jdocs_infos(x):
    '''save remane jdocs'''
    try:
        x = x.fillna('')
        global address_tmp;address_tmp = x.copy()
        x.to_excel(address_tmp_xlsx,index=False)
        # print_log('保存地址文件到 => %s...'%address_tmp_xlsx)
    except Exception as e:
        print_log('%s <= 保存失败,请检查... %s'%(address_tmp_xlsx,e))
        
def copy_rows_user_func(dfj,dfo):
    '''copy users line regard adr user'''
    errs = ['【OA无用户记录】','【怀疑用户错别字】','【用户字段重复】','【相似】','【系列】'];urow = errs[0]
    for (i,dfjr) in dfj.iterrows():
        dfj.loc[i,'当事人'] = errs[0]
        if isinstance(dfjr['new_adr'],dict):
            if not dfjr['new_adr']:# 提取jdocs字段失败
                dfj.loc[i,'当事人'] = '【检查判决书用户(冒号)】';continue
            dfs = dfo[dfo['原一审案号']==dfjr['原一审案号']]
            if len(dfs) > 0:# 同案号，继续
                dfj.loc[i,'当事人'] = copy_users_compare(dfjr,dfs.iloc[0],errs);continue
            else:#[::-1] # 没案号则遍历源数据dfo
                for (j,dfor) in dfo[['原一审案号','当事人']][::-1].iterrows():
                    urow = copy_users_compare(dfjr,dfor,errs)
                    if urow not in errs:
                        dfj.loc[i,'当事人'] = urow;break
                    elif urow == errs[4]:# 假如相似数据则添加一行jdocs数据
                        dfj = dfj.append(DataFrame({'原一审案号':[dfor['原一审案号']],
                                                '当事人':[dfor['当事人']],
                                                'new_adr':[dfjr['new_adr']]}),
                                                sort=False)
    save_jdocs_infos(dfj)
    return dfj

#%%   
def rename_jdocs_codes_x(d,r):
    '''add jdoc current case codes for reference 判决书改名，包括源案号'''
    if str(r['原一审案号_y']) in str(d):
        nd = os.path.join(os.path.split(d)[0],'判决书_'+str(r['案号']) +'_原_'+ str(r['原一审案号_y']) + '.docx')
        if(d == nd):return d
        try:
            if os.path.exists(nd):
                os.remove(nd)
            if '_原_' in d:
                shutil.copyfile(d,nd)
            else:
                os.rename(d,nd)
        except Exception as e:
            print_log(e)
        return nd
    return d

def rename_jdocs_codes(docs,df):
    '''rename jdoc now case codes for reference'''
    df = df[df['原一审案号_y'] != '']
    for (i,r) in df.iterrows():
        docs = list(map(lambda x:rename_jdocs_codes_x(x,r),docs))
    return docs

def rename_jdocs_codes_func(x):
    '''rename with new codes'''
    docs = glob(parse_subpath(jdocs_path,'判决书_*.docx'))
    if docs:rename_jdocs_codes(docs,x)
    return x
#%%
def merge_dfj_dfo(dfj,dfo):
    x = merge(dfo,dfj,how='left',on=['当事人'])
    x = titles_resort(x,['立案日期','案号','原一审案号_x','原一审案号_y'])
    x = x.drop_duplicates(['案号','当事人'])
    x['原一审案号_y'] = x['原一审案号_y'].fillna(x['原一审案号_x'])
    return x
    
def copy_rows_agent_adr_func(x):
    xx = x[['当事人','诉讼代理人','地址','new_adr','案号']].copy()
    xx['地址'] = xx.apply(lambda x:copy_rows_adr(x), axis=1)# copy address
#    if flag_fill_phone: xx['诉讼代理人'] = xx.apply(lambda x:copy_rows_agent(x), axis=1)# 填充伪手机
    x[['地址','诉讼代理人']] =  xx[['地址','诉讼代理人']]
    return x

def fill_infos_clean(x):
    x['原一审案号'] = x['原一审案号_x']
    x = x.drop(['原一审案号_x','原一审案号_y','new_adr'],axis=1).fillna('')
    return x
    
def fill_infos_func(dfj,dfo):
    '''combine address between data and judgment docs and delete duplicate'''
    dfj = copy_rows_user_func(dfj,dfo)
    x = merge_dfj_dfo(dfj,dfo)
    x = rename_jdocs_codes_func(x)
    x = copy_rows_agent_adr_func(x)
    x = fill_infos_clean(x)
    return x

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
    docs = glob(parse_subpath(jdocs_path,'*.docx')) # get jdocs
    if not docs:return dfo
    dfj = get_all_jdocs(docs)
    if len(dfj) == 0:print_log('>>> 没有找到判决书...不处理！！');return dfo
    if len(dfo) == 0:return dfo
    dfn = fill_infos_func(dfj,dfo)
    dfn = titles_resort(dfn,titles_cn)
    try:
        if flag_fill_jdocs_infos:dfo = save_adjust_xlsx(dfn,data_xlsx)
    except PermissionError:
        print_log('>>> %s 文件已打开...填充判决书地址失败！！...请关闭并重新执行'%data_xlsx)
    return dfo
#%%
def df_make_subset(df):
    '''
    cut orgin data into subset by conditions
    '''
    dcn = data_case_codes #（2018）哈哈1234号
    date_range = data_date_range
    last_lines = data_last_lines
    
    if dcn:
        df = df[df['案号'].isin(re.split('[,，;；]',dcn))]
    elif ':' in date_range:
        print_log('>>> 预定读取【%s】'%date_range)
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
            df = df[(df['立案日期']>=date_start)&(df['立案日期']<=date_end)]
            df['立案日期'] = df['立案日期'].astype(str)
            return df
        except Exception as e:
            print_log('>>> 日期异常',e)
    elif last_lines:
        df = df.tail(last_lines)
    return df

#%%
def df_oa_append(df_orgin):
    '''main fill OA data into df data and mark new add'''
    if flag_append_oa:
        if not os.path.exists(data_oa_xlsx):
            print_log('>>> 没有找到OA模板 %s...不处理！！'%data_oa_xlsx);return df_orgin
        df_oa = read_excel(data_oa_xlsx,sort=False)[titles_oa].fillna('') # only oa columns
        df1 = df_orgin.copy();df2 = df_oa.copy()
        df2 = df_make_subset(df2) # subset by columns
        df2.rename(columns={'承办人':'主审法官'},inplace=True)
        df2 = df_read_fix(df2) # fix empty data columns
        df2['add_index'] = 'new'
        df1['add_index'] = 'old'
        df3 = df1.append(df2,sort=False).drop_duplicates(['立案日期','案号'],keep='first')
        df3.sort_values(by=['立案日期','案号'],inplace=True)
        df_noa = df3[df3['add_index'] == 'new']
        print_log('>>> 所有OA记录【%s条】...'%len(df_oa))
        print_log('>>> 原Data记录【%s条】...'%len(df_orgin))
        print_log('>>> 实际添加【%s条】新OA记录...'%len(df_noa))
        if len(df_noa):
            dd = str(df_noa['立案日期'].iloc[0]) +':'+df_noa['立案日期'].iloc[-1]
            print_log('>>> 实际添加【%s】'%dd)
        if any(df3['add_index'] == 'new'):
            df3 = titles_resort(df3,titles_cn)
            try:df_orgin = save_adjust_xlsx(df3,data_xlsx)
            except PermissionError:print_log('>>> %s 文件已打开...填充OA数据失败！！。。。请关闭并重新执行'%data_xlsx)
    return df_orgin

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
                if flag_print: print('A=%s,B=%s'%(x,name))
                x = name;break
    x = re.sub(r'_.*','',x)
    x = re.sub(path_names_clean,'',x)
    return x

def clean_rows_adr(adr):
    '''clean adr format'''
    y = re.split(r'[,，]',adr)
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
    read_config()
except Exception as e:
    print_log('>>> 配置文件出错 %s ,删除...'%e)
    if os.path.exists(cfgfile):
        os.remove(cfgfile)
    try:
        write_config()
        read_config()
    except Exception as e:
        '''这里可以添加配置问题预判问题'''
        print_log('>>> 配置文件再次生成失败 %s ...'%e)
        set_default_value(data_date_range = '')
        
print_log('''>>> 正在处理...
    主表路径 = %s
    指定日期 = %s
    重命名判决书 = %s
    填充判决书地址 = %s
    导入OA数据 = %s
    打印邮单 = %s
    输出判决书提示 = %s
    输出邮单提示 = %s
    '''%(os.path.abspath(data_xlsx),str(data_date_range),
    str(flag_rename_jdocs),str(flag_fill_jdocs_infos),str(flag_append_oa),str(flag_to_postal),
    str(flag_check_jdocs),str(flag_check_postal)))

if not os.path.exists(data_xlsx):
    save_adjust_xlsx(DataFrame(columns=titles_cn),data_xlsx,width=40)
    print_log('>>> %s 记录文件不存在...重新生成'%(data_xlsx))

df_orgin = read_excel(data_xlsx,sort=False).fillna('') #真正读取记录位置
df_orgin = df_read_fix(df_orgin) # fix empty data columns
df_orgin = df_oa_append(df_orgin) # append oa data
df_orgin = df_fill_infos(df_orgin) # fill jdocs infos
df_orgin = df_make_subset(df_orgin)
df = titles_trans_columns(df_orgin,titles_cn) # 中译英方便后面处理

if flag_check_postal:df.apply(lambda x:df_check_format(x), axis=1)
print_log('>>> 将要打印Data记录【%s条】...'%len(df))

#%% df tramsfrom stream 数据转换流程

if len(df) and flag_to_postal:
    try:
        print_log('>>> 开始生成新数据 data_main_temp...')
        '''获取 datetime|number|prenumber|judge'''
        number = df[titles_en[:4]]
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

        del user,number,agent,adr,df,agent_adr,opt
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

if flag_to_postal:
    print_log('>>> 正在输出邮单...')
    if not os.path.exists(sheet_docx):
        input('>>> 没有找到邮单模板 %s...任意键退出'%sheet_docx);sys.exit()
    df_p = df_x.apply(re_write_text,axis = 1)
    count = len(df_p[df_p != ''])
    codes = df_x['number'].astype(str)
    dates = df_x['datetime'].astype(str)
    print_log('>>> 最终生成邮单【%s条】范围: 【%s:%s】日期:【%s:%s】'%(count,codes.iloc[0],codes.iloc[-1],dates.iloc[0],dates.iloc[-1]))
    del df_x,df_p,codes,dates
    
#%% main finish 结束所有

print_log('>>> 全部完成,可以回顾记录...任意键退出')
#input('>>> 全部完成,可以回顾记录...任意键退出');sys.exit()



