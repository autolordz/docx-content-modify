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
from StyleFrame import StyleFrame, Styler
from collections import Counter
from docx import Document
from glob import glob
set_option('max_colwidth',500)
set_option('max_rows', 50)
set_option('max_columns',50)

#%% print_log log
logname = 'hist.txt'
if os.path.exists(logname):
    os.remove(logname)
def print_log(*args, **kwargs):
    print(*args, **kwargs)
    # with codecs.open(logname, "a", "utf-8-sig") as file:
    with open(logname, "a",encoding='utf-8') as file:
        print(*args, **kwargs, file=file)

#%%
print_log('''
Auto generate word:docx from excel:xslx file.

Updated on Thu Oct 25 2018

Depends on: python-docx,pandas.

@author: Autoz
''')
#%% config and default values
cfgfile = 'conf.txt'
data_xlsx = 'data_main.xlsx'
data_oa_xlsx = 'data_oa.xlsx'
sheet_docx = 'sheet.docx'
address_tmp_xlsx = 'address_tmp.xlsx'
postal_path = os.path.join('.','postal')
jdocs_path = os.path.join('.','jdocs')
last_lines_oa = 50
last_lines_data = 50
date_range = '2018-09-01:2018-12-01'
flag_rename_jdocs = True
flag_fill_jdocs_adr = True
flag_fill_phone = False
flag_append_oa = True
flag_to_postal = True
flag_check_jdocs = False
flag_check_postal = False
cut_tail_lines = True
#%%
def write_config(cfgfile):
    cfg = configparser.ConfigParser(allow_no_value=True,
                                    inline_comment_prefixes=('#', ';'))
    cfg['config'] = {'data_xlsx': data_xlsx+' # 数据模板地址',
                     'data_oa_xlsx': data_oa_xlsx+' # OA数据地址',
                     'sheet_docx': sheet_docx+' # 邮单模板地址',
                     'flag_rename_jdocs': str(flag_rename_jdocs)+' # 是否重命名判决书',
                     'flag_fill_jdocs_adr': str(flag_fill_jdocs_adr)+' # 是否填充判决书地址',
                     'flag_fill_phone': str(flag_fill_phone)+' # 是否填充伪手机',
                     'flag_append_oa': str(flag_append_oa)+' # 是否导入OA数据',
                     'flag_to_postal': str(flag_to_postal)+' # 是否打印邮单',
                     'flag_check_jdocs': str(flag_check_jdocs)+' # 是否检查用户格式,输出提示信息',
                     'flag_check_postal': str(flag_check_postal)+' # 是否检查邮单格式,输出提示信息',
                     'date_range': '# '+date_range+' # 打印数据日期范围,比行数优先,去掉注释后读取,井号注释掉',
                     'last_lines_oa': str(last_lines_oa)+' # 导入OA数据的最后几行,当flag_append_oa开启才有效',
                     'last_lines_data': str(last_lines_data)+' # 打印数据的最后几行'
                     }
    # with codecs.open(cfgfile, 'w', 'utf-8-sig') as configfile:
    with open(cfgfile, 'w',encoding='utf-8-sig') as configfile:
        cfg.write(configfile)
# write_config(cfgfile)
#%%
# cfg.read(cfgfile,'utf-8-sig')

#%% global variable
titles_cn = ['立案日期','案号','原一审案号','主审法官','当事人','诉讼代理人','地址']
titles_en = ['datetime','number','prenumber','judge','uname','aname','address']
titles_oa = ['立案日期','案号','原一审案号','承办人','当事人']
titles_cn2en = dict(zip(titles_cn, titles_en))
titles_en2cn = dict(zip(titles_en, titles_cn))

path_names_clean = re.compile(r'[^A-Za-z\u4e00-\u9fa5（）()：]') # remain only name including old name 包括括号冒号
search_names_phone = lambda x: re.search(r'[\w（）()：:]+\_\d+',x)  # phone numbers
path_code_ix = re.compile(r'[(（][0-9]+[)）].*?号') # case numbers
adr_tag = '/地址：'

#%% read func
def user_to_list(u):
    '''get name from user list'''
    u = re.split(r'[:、,，]',u)
    return [x for x in u if not re.search(r'申请人|被申请人|原告|被告|原审被告|上诉人|被上诉人|第三人|原审诉讼地位',x)]

def case_codes_fix(x):
    '''fix string with chinese codes format'''
    x = str(x)
    x = re.search(path_code_ix.pattern+r'|$',x).group().strip().replace(' ','')
    x = x.replace('(','（').replace(')','）')
    return x

def parse_subpath(path,file):
    '''make subpath'''
    if not os.path.exists(path):
        os.mkdir(path)
    return os.path.join(path,file)

def check_contain_chinese(x):
    '''check if string contain chinese'''
    return bool(re.search(r'[\u4e00-\u9fa5]',str(x)))

def parse_datetime(date):
    try:date = datetime.datetime.strptime(date,'%Y-%m-%d')
    except ValueError:print_log('时间范围格式有误,默认选取全部日期');date = ''
    return date

def titles_trans(df_list):
    '''change titles between Chinese and English'''
    flag = check_contain_chinese(df_list[0])
    tdict = titles_cn2en if flag else titles_en2cn
    # print_log('Change titles to:', 'Chinese' if flag else 'English')
    return [tdict.get(x) for x in df_list if x in tdict.keys()]

def titles_resort(df,titles):
    '''resort titles with orders'''
    titles_rest = df.drop(titles,axis=1).columns.tolist()
    return df[titles + titles_rest]

def titles_combine_en(df,titles):
    '''refer to titles to sub-replace with columns titles'''
    titles_rest = df.drop(titles,axis=1).columns.tolist()
    df = df[titles + titles_rest]
    df.columns = titles_trans(titles) + titles_rest
    return df

def save_adjust_xlsx(df,file='test.xlsx',width=60):
    '''save and re-adjust excel format'''
    df.reset_index(drop='index',inplace=True)
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
    sf.to_excel(file,best_fit=sf.data_df.columns.difference(['当事人', '诉讼代理人', '地址']).tolist()).save()
    print_log('>>> 保存文件 => 文件名 \'%s\' => 数据保存成功...' %(file))
    return df

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
    except Exception as e: input('读取错误 %s 请关闭或检查docx文档 '%e);sys.exit()
    if tables: codes = read_jdocs_table(tables)
    paras = Document(doc).paragraphs
    if not paras: return codes,adrs
    if len(paras) > 20:lines = int(len(paras)/2)
    parass = paras[:lines]
    for i,para in enumerate(parass):
        x = para.text.strip()
        if len(x) > 150: continue
        if re.search(path_code_ix,x) and len(x) < 25:
            codes = case_codes_fix(x);continue # codes
        if re.search(r'诉讼|代理|律师|请求|证据|辩|不服',x): continue # filter agent
        if re.search(r'(?<=[：:]).*?(?=[,，。])',x) and re.search(r'(户[籍口]|居住|所在地?|住所地?|住址?|原住|现住?).*?[省市州县区乡镇村]',x):
            try:
                name = re.search(r'(?<=[：:]).*?(?=[,，。])|$',x).group(0).strip()
                name = name if re.search(r'[(（]曾用名.*?[）)]',x) else re.sub(r'[(（].*?[）)]','',name)
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
            # print_log('>>> %s 重命名判决书1 => %s'%(doc,jdoc_name))
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
        if codes and flag_rename_jdocs:
            rename_jdoc_x(doc,codes)
        numlist.append(codes)
        nadr.append(adrs)
    numlist = list(map(case_codes_fix,numlist))
    x = DataFrame({'原一审案号':numlist,'new_adr':nadr})
    return x

def rename_jdocs_codes_x(d,r):
    '''add jdoc now case codes for reference'''
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
            # print_log('>>> %s 重命名判决书2 => %s'%(d,nd))
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

def save_jdocs_infos(x):
    try:
        global address_tmp
        address_tmp = x.copy()
        x.to_excel(address_tmp_xlsx,index=False)
        print_log('保存地址文件到 => %s...'%address_tmp_xlsx)
    except Exception as e:
        print_log('%s <= 保存失败,请检查... %s'%(address_tmp_xlsx,e))

def copy_rows_adr(x):
    ''' copy jdocs address to address column'''
    x[:3] = x[:3].astype(str)
    user = x[0];agent = x[1];adr = x[2];n_adr = x[3]
    if not isinstance(n_adr,dict):return adr
    y = re.split(r'[,，]',adr)
    for i,k in enumerate(n_adr):
        by_agent = any([k in ag for ag in re.findall(r'[\w+、]*\/[\w+]*',agent)])
        if by_agent and k in adr: # remove user's address when user with agent
            y = list(filter(lambda x:not k in x,y))
        if type(n_adr) == dict and not k in adr and k in user and not by_agent:
            y += [k+adr_tag+n_adr.get(k)] # append address by rules
    adr =  '，'.join(list(filter(None, y)))
    return adr

def copy_rows_agent(x):
    '''copy example phone-numbers to agent column if empty'''
    x = x.astype(str)
    user = x[0];agent = x[1];adr = x[2]
    users = user_to_list(user)
    y = [agent]
    for i,u in enumerate(users):
        if not u in agent and u in adr:
            y += [u+'_12345678910'] # fake tel
    agent = ','.join(list(filter(None, y)))
    return agent

def copy_users_compare(jrow,df):
    x = Counter(list(jrow['new_adr'].keys()))
    y = Counter(user_to_list(df['当事人']))
    dx = list((x-y).elements())
    dy = list((y-x).elements())
    rxy = len(list((x&y).elements()))/len(list((x|y).elements()))
    # if rxy > 0:print('rxy:',rxy)
    if rxy == 0:return ''
    if max(x.values()) > 1 or max(y.values()) > 1:
        xdu = [k for k,v in x.items() if v > 1]
        ydu = [k for k,v in y.items() if v > 1]
        print_log('>>> 判决书和OA用户重复【%s】【判决书:%s:%s】-【OA或Data:%s:%s】-【【判断出错！！[不]拷贝地址！！请检查修改OA或Data！！】'
                  %("{0:.0%}".format(1-rxy),jrow['原一审案号'],xdu,df['原一审案号'],ydu))
        return ''
    if rxy == 1:return df['当事人']
    if 0 < rxy < 1:
        xx = Counter(''.join(dx))
        yy = Counter(''.join(dy))
        rxxyy = len(list((xx&yy).elements()))/len(list((xx|yy).elements()))
        if rxxyy >= .6:
            print_log('>>> 判决书和OA用户差异【%s】【判决书:%s:%s】-【OA或Data:%s:%s】-【判断出错！！[不]拷贝地址！！请检查修改OA或Data！！】'
                      %("{0:.0%}".format(1-rxy),jrow['原一审案号'],dx,df['原一审案号'],dy))
            return ''
        if flag_check_jdocs:
            print_log('>>> 判决书和OA用户疑似差异【%s】【判决书:%s:%s】-【OA或Data:%s:%s】-【判断案件[相同]拷贝地址】'
                      %("{0:.0%}".format(1-rxy),jrow['原一审案号'],dx,df['原一审案号'],dy))
        return df['当事人']
    return ''

def copy_rows_users(jrow,df):
    '''copy users line regard adr user'''
    urow = ''
    if isinstance(jrow['new_adr'],dict):
        dfjr = df[df['原一审案号']==jrow['原一审案号']]
        if len(dfjr) > 0:
            urow = copy_users_compare(jrow,dfjr.iloc[0])
        else:#[::-1]
            for (i,r) in df[['原一审案号','当事人']][::-1].iterrows():
                urow = copy_users_compare(jrow,r)
                if urow:break
    return urow

def func_infos_fill(df_jdocs,df_orgin):
    '''combine address between data and judgment docs and delete duplicate'''
    df_jdocs['当事人'] = df_jdocs.apply(lambda x:copy_rows_users(x,df_orgin), axis=1)
    # save infos
    save_jdocs_infos(df_jdocs)
    x = merge(df_orgin,df_jdocs,how='left',on=['当事人'])
    x['原一审案号_y'] = x['原一审案号_y'].fillna(x['原一审案号_x'])
    x = x.drop_duplicates(['案号','当事人']).fillna('')
    # rename with new codes
    docs = glob(parse_subpath(jdocs_path,'判决书_*.docx'))
    if flag_fill_jdocs_adr and docs:rename_jdocs_codes(docs,x)
    # after rename then clean df codes
    x['原一审案号'] = x['原一审案号_x']
    x.drop(['原一审案号_x','原一审案号_y'],axis=1,inplace=True)
    xx = x.loc[:,['当事人','诉讼代理人','地址','new_adr']]
    xx['地址'] = xx.apply(lambda x:copy_rows_adr(x), axis=1)
    if flag_fill_phone: xx['诉讼代理人'] = xx.apply(lambda x:copy_rows_agent(x), axis=1)
    x[['地址','诉讼代理人']] =  xx[['地址','诉讼代理人']]
    x = x.drop(['new_adr'],axis=1).fillna('')
    return x

#%% df process steps

def df_read_fix(df):
    '''fix codes remove error format'''
    df[['立案日期','案号','主审法官','当事人']] = df[['立案日期','案号','主审法官','当事人']].replace('',float('nan'))
    df.dropna(how='any',subset=['立案日期','案号','主审法官','当事人'],inplace=True)
    df['原一审案号'] = df['原一审案号'].fillna('')
    df[['案号','原一审案号']] = df[['案号','原一审案号']].applymap(case_codes_fix)
    return df

def df_infos_fill(df_orgin):
    '''main fill jdocs infos'''
    # docs = glob(parse_subpath(jdocs_path,'判决书_*.docx'))
    docs = glob(parse_subpath(jdocs_path,'*.docx'))
    if not docs:return df_orgin
    df_jdocs = get_all_jdocs(docs)
    if len(df_jdocs) == 0:print_log('>>> 没有找到判决书...不处理！！');return df_orgin
    if len(df_orgin) == 0:return df_orgin
    df_new = func_infos_fill(df_jdocs,df_orgin)
    df_new = titles_resort(df_new,titles_cn)
    try:
        df_new = save_adjust_xlsx(df_new,data_xlsx)
        df_orgin = df_new
    except PermissionError:
        print_log('>>> %s 文件已打开...填充判决书地址失败！！...请关闭并重新执行'%data_xlsx)
    return df_orgin

def df_oa_append(df_orgin):
    '''main fill OA data into df data and mark new add'''
    if flag_append_oa:
        if not os.path.exists(data_oa_xlsx):
            print_log('>>> 没有找到OA模板 %s...不处理！！'%data_oa_xlsx);return df_orgin
        df_oa = read_excel(data_oa_xlsx,sort=False,na_values='').tail(last_lines_oa)[titles_oa].fillna('') # only oa columns
        df_oa.rename(columns={'承办人':'主审法官'},inplace=True)
        df_oa = df_read_fix(df_oa) # fix empty data columns
        df1 = df_orgin.copy()
        df2 = df_oa
        df2['add_index'] = 'new'
        df1['add_index'] = 'old'
        df3 = df1.append(df2,sort=False).drop_duplicates(['立案日期','案号'],keep='first')
        df3.sort_values(by=['立案日期','案号'],inplace=True)
        if any(df3['add_index'] == 'new'):
            df3 = titles_resort(df3,titles_cn)
            try:
                df_orgin = save_adjust_xlsx(df3,data_xlsx)
            except PermissionError:print_log('>>> %s 文件已打开...填充OA数据失败！！。。。请关闭并重新执行'%data_xlsx)
        print_log('>>> 预计添加【%s条】OA记录...'%last_lines_oa)
        print_log('>>> 实际添加【%s条】新OA记录...'%len(df3[df3['add_index'] == 'new']))
    return df_orgin

def df_make_subset(df_orgin):
    '''cut orgin data into subset df data'''
    df = titles_combine_en(df_orgin,titles_cn)
    dfs = df.copy()
    if ':' in date_range:
        print_log('>>> 找到日期范围【%s】，优先读取日期'%date_range)
        df['datetime'] = to_datetime(df['datetime'])
        df.sort_values(by=['datetime'],inplace=True)
        dats = date_range.split(':')
        dats[0] = parse_datetime(dats[0]);dats[1] = parse_datetime(dats[1]);
        date_start = dats[0] if dats[0] else df['datetime'][0]
        date_end = dats[1] if dats[1] else df['datetime'].iloc[-1]
        df = df[(df['datetime']>date_start)&(df['datetime']<date_end)]
        df['datetime'] = df['datetime'].astype(str)
        if len(df) == 0:df = dfs
    elif last_lines_data:
        print_log('>>> 预定读取【%s条】记录...'%last_lines_data)
        df = df.tail(last_lines_data) if cut_tail_lines else df.head(last_lines_data)
    print_log('>>> 实际读取【%s条】记录...'%len(df))
    return df

def df_check_format(x):
    '''check data address and agent format with check flag'''
    if x['aname']!='' and not re.search(r'[\/_]',x['aname']):
        print_log('>>> 记录\'%s\'---- 【诉讼代理人】格式 \'%s\' 不正确,如无请留空,请自行修改...'%(x['number'],x['aname']))
    if x['address']!='' and not re.search(r'\/地址[:：]',x['address']):
        print_log('>>> 记录\'%s\'---- 【地址】格式 \'%s\' 不正确,如无请留空,请自行修改...'%(x['number'],x['address']))
    return x

#%% main processing stream
try:
    if not os.path.exists(cfgfile):
        write_config(cfgfile)
        print_log('>>> 生成配置文件...')
    # cfg.read(cfgfile,'utf-8-sig')
    cfg = configparser.ConfigParser(allow_no_value=True,
                                    inline_comment_prefixes=('#', ';'))
    cfg.read(cfgfile,encoding='utf-8-sig')
    data_xlsx = cfg['config']['data_xlsx']
    data_oa_xlsx = cfg['config']['data_oa_xlsx']
    sheet_docx = cfg['config']['sheet_docx']
    last_lines_oa = cfg['config'].getint('last_lines_oa')
    last_lines_data = cfg.getint('config', 'last_lines_data')
    date_range = cfg.get('config', 'date_range')
    flag_rename_jdocs = cfg.getboolean('config', 'flag_rename_jdocs')
    flag_fill_jdocs_adr = cfg.getboolean('config', 'flag_fill_jdocs_adr')
    flag_fill_phone  = cfg.getboolean('config', 'flag_fill_phone')
    flag_append_oa = cfg.getboolean('config', 'flag_append_oa')
    flag_to_postal = cfg.getboolean('config', 'flag_to_postal')
    flag_check_jdocs = cfg.getboolean('config', 'flag_check_jdocs')
    flag_check_postal = cfg.getboolean('config', 'flag_check_postal')
except Exception as e:
    # raise e
    print_log('>>> 配置文件出错 %s ,已删除...'%e)
    if os.path.exists(cfgfile):
        os.remove(cfgfile)
        
print_log('''>>> 正在处理...
    主表路径 = %s
    指定日期 = %s
    输出判决书提示 = %s
    输出邮单提示 = %s
    '''%(os.path.abspath(data_xlsx),str(date_range),str(flag_check_jdocs),str(flag_check_postal)))

if os.path.exists(data_xlsx):
    df_orgin = read_excel(data_xlsx,sort=False).fillna('')
    df_orgin = df_read_fix(df_orgin) # fix empty data columns
    df_orgin = df_oa_append(df_orgin)
    df_orgin = df_infos_fill(df_orgin)
    df = df_make_subset(df_orgin)
    if flag_check_postal:df.apply(lambda x:df_check_format(x), axis=1)
    print_log('>>> 读取记录成功...')
else: input('>>> %s 记录文件不存在...任意键退出'%(data_xlsx));sys.exit()

#%% df tramsfrom functions
def clean_rows_aname(x,names):
    '''Clean agent name for agent to match address's agent name'''
    if names:
        for name in names:
            if not check_contain_chinese(name):continue
            if name in x:x = x.replace(x,name);break
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

def clean_rows_agent(agent):
    '''clean agent format'''
    y = re.split(r'[,，]',agent)
    if y:
        def clean_x(x):
            x = x.replace('/','') + '/' + x.replace('/','') if re.search(r'^/|/$',x) else x
            x = x + '/' + x if not re.search(r'/',x) else x
            x = re.sub(r'(?<=[\w()（）])_.*?(?=/\w)','',x)
            return x
        y = list(map(clean_x,y))
        agent = '，'.join(list(filter(None, y)))
    return agent

def make_adr(adr,clean_aname=[]):
    adr = adr[adr != '']
    adr = adr.str.strip().str.split(r'[,，。]',expand=True).stack()
    adr = adr.str.strip().apply(lambda x:clean_rows_adr(x))
    adr = adr.str.strip().str.split(r'\/地址[:：]',expand=True).fillna('')
    adr.columns = ['aname','address']
    adr['clean_aname'] = adr['aname'].str.strip().apply(lambda x:clean_rows_aname(x,clean_aname)) # clean adr
    adr = adr.reset_index().drop(['level_1','aname'],axis=1)
    return adr

def make_agent(agent,clean_aname=[]):
    agent = agent[agent != '']
    agent = agent.str.strip().str.split(r'[,，。]',expand=True).stack() #Series
    agent = agent.str.strip().apply(lambda x:clean_rows_agent(x))
    agent = agent.str.strip().str.split(r'\/',expand=True).fillna('') #DataFrame
    agent.columns = ['uname','aname']
    agent['clean_aname'] = agent['aname'].str.strip().apply(lambda x: clean_rows_aname(x,clean_aname))
    dd_l = agent['uname'].str.strip().str.split(r'、',expand=True).stack().to_frame(name = 'uname').reset_index()
    dd_r = agent[agent.columns.difference(['uname'])].reset_index()
    agent = merge(dd_l,dd_r,how='outer',on=['level_0','level_1']).drop(['level_1','level_2'],axis=1).fillna('')
    return agent

def merge_user(user,agent):
    return merge(user,agent,how='left',on=['level_0','uname']).fillna('')

def merge_agent_adr(agent,adr):
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

def sort_data(x):
    x = x[['level_0','uname','aname','address']].sort_values(by=['level_0'])
    x = merge(number,x,how='right',on=['level_0']).drop(['level_0'],axis=1).fillna('')
    return x

#%% df tramsfrom stream

if flag_to_postal:
    try:
        print_log('>>> 开始生成新数据...')
        number = df[titles_en[:4]]
        number = number.reset_index()
        number.columns.values[0] = 'level_0'
        # user
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
            adr = make_adr(adr)
            agent = make_agent(agent,adr['clean_aname'].tolist())
            agent = merge_user(user,agent)
            df_x = reclean_data(merge_agent_adr(agent,adr))
            df_x = sort_data(df_x)
        elif opt.address:
            print_log('>>> 只有【地址】...正在处理...')
            adr = make_adr(adr,user['uname'].tolist())
            adr['uname'] = adr['clean_aname']
            adr = merge_user(user,adr)
            adr = adr.assign(aname='')
            df_x = reclean_data(adr)
            df_x = sort_data(df_x)
        elif opt.aname:
            print_log('>>> 只有【代理人】...正在处理...')
            agent = make_agent(agent)
            agent = merge_user(user,agent)
            agent = agent.assign(address='')
            df_x = reclean_data(agent)
            df_x = sort_data(df_x)
        else:
            print_log('>>> 缺失【代理人】和【地址】...正在处理...')
            agent_adr.index.name = 'level_0'
            agent_adr.reset_index(inplace=True)
            df_x = merge(user,agent_adr,how='left',on=['level_0']).fillna('')
            df_x = sort_data(df_x)

        del user,agent,number,adr,df,agent_adr

        if len(df_x):
            data_tmp = os.path.splitext(data_xlsx)[0]+"_tmp.xlsx"
            df_save = df_x.copy()
            df_save.columns = titles_trans(df_save.columns.tolist())
            try:df_save = save_adjust_xlsx(df_save,data_tmp,width=40)
            except PermissionError: print_log('>>> %s 文件已打开...请手动关闭并重新执行...保存失败'%data_tmp)

    except Exception as e:
        raise e
        print_log('>>> 错误 \'%s\' 生成数据失败,请检查源 \'%s\' 文件...退出...'%(e,data_xlsx));sys.exit()
#%% generate postal sheets

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
    # if not search_names_phone(agent_text):
    #     if flag_check_postal:
    #         print_log('>!> %s => 手机格式不对,请自行修改'% agent_text)
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
    print_log('>>> 生成邮单【%s条】范围: 【%s:%s】日期:【%s:%s】'%(count,codes.iloc[0],codes.iloc[-1],dates.iloc[0],dates.iloc[-1]))
    # re_write_text(df_x.iloc[1])

#%% main finish
# if __name__ == "__main__":
print_log('>>> 全部完成,可以回顾记录...任意键退出')


