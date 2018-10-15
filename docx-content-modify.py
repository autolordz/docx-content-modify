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
print("""
Auto generate word:docx from excel:xslx file.

Updated on Thu Oct 11 2018

Depends on: python-docx,pandas.

@author: Autoz
""")

#%%
import os,re,sys,datetime,codecs,configparser
from pandas import DataFrame, read_excel, merge, concat, set_option, to_datetime
from docx import Document
from glob import glob
from StyleFrame import StyleFrame, Styler
set_option('max_colwidth',500)
set_option('max_rows', 50)
set_option('max_columns',50)

#%% print_log log
logname = 'history.txt'
if os.path.exists(logname):
    os.remove(logname)
def print_log(*args, **kwargs):
    print(*args, **kwargs)
    with codecs.open(logname, "a", "utf-8-sig") as file:
        print(*args, **kwargs, file=file)

#%% config and default values
cfgfile = 'config.txt'
data_xlsx = 'data_main.xlsx'
data_oa_xlsx = 'data_oa.xlsx'
sheet_docx = 'sheet.docx'
postal_path = os.path.join('.','postal')
jdocs_path = os.path.join('.','jdocs')
OA_last_lines = 100
data_last_lines = 100
date_range = '2018-06-01:2018-08-01'
flag_rename_jdocs = True
flag_fill_jdocs_adr = True
flag_fill_phone = False
flag_append_oa = True
flag_to_postal = True
flag_check_data = True
cut_tail_lines = True

def write_config(cfgfile):
    with codecs.open(cfgfile, "w", "utf-8-sig") as configfile:
        cfg = configparser.RawConfigParser(allow_no_value=True)
        cfg.add_section('config')
        cfg.set('config', '# 数据模板地址')
        cfg.set('config', 'data_xlsx', data_xlsx)
        cfg.set('config', '# OA数据地址')
        cfg.set('config', 'data_oa_xlsx', data_oa_xlsx)
        cfg.set('config', '# 邮单模板地址')
        cfg.set('config', 'sheet_docx', sheet_docx)
        cfg.set('config', '# 是否重命名判决书')
        cfg.set('config', 'flag_rename_jdocs', flag_rename_jdocs)
        cfg.set('config', '# 是否填充判决书地址')
        cfg.set('config', 'flag_fill_jdocs_adr', flag_fill_jdocs_adr)
        cfg.set('config', '# 是否填充伪手机')
        cfg.set('config', 'flag_fill_phone', flag_fill_phone)
        cfg.set('config', '# 是否导入OA数据')
        cfg.set('config', 'flag_append_oa', flag_append_oa)
        cfg.set('config', '# 导入OA数据的最后几行')
        cfg.set('config', 'OA_last_lines', OA_last_lines)
        cfg.set('config', '# 是否打印邮单')
        cfg.set('config', 'flag_to_postal', flag_to_postal)
        cfg.set('config', '# 打印数据模板的最后几行')
        cfg.set('config', 'data_last_lines', data_last_lines)
        cfg.set('config', '# 打印数据模板的日期范围')
        cfg.set('config', 'date_range', date_range)
        cfg.set('config', '# 检查数据模板的内容格式')
        cfg.set('config', 'flag_check_data', flag_check_data)
        cfg.write(configfile)

#%% global variable
titles_cn = ['立案日期','案号','原一审案号','主审法官','当事人','诉讼代理人','地址']
titles_en = ['datetime','number','prenumber','judge','uname','aname','address']
titles_oa = ['立案日期','案号','原一审案号','承办人','当事人']
titles_cn2en = dict(zip(titles_cn, titles_en))
titles_en2cn = dict(zip(titles_en, titles_cn))

path_names_clean = re.compile(r'[^A-Za-z\u4e00-\u9fa5（）()：]') # remain only name including old name 包括括号冒号
search_names_phone = lambda x: re.search(r'[\w（）()：:]+\_\d+',x)  # phone numbers
path_code_ix = re.compile(r'[(（][0-9]+[)）].*?号') # case numbers
path_adr_ix = re.compile(r'住|(市(.*)[0-9])') # chinese address
adr_tag = '/地址：'

#%% read func
def user_to_list(u):
    '''get name from user list'''
    u = re.split(r'[:、,，]',u)
    return [x for x in u if not re.search(r'申请人|被申请人|原告|被告|原审被告|上诉人|被上诉人|第三人',x)]

def case_codes_fix(x):
    x = str(x)
    '''fix string with chinese codes format'''
    y = re.search(path_code_ix,x)
    if re.search(path_code_ix,x): x = y.group()
    x = x.replace('(','（').replace(')','）').replace(' ','')
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
                                  # cols_to_style=['add_index'],
                                  overwrite_default_style=False)
        sf.apply_column_style(cols_to_style=['当事人', '诉讼代理人', '地址'],
                              width = width,
                              styler_obj=Styler(wrap_text=False,shrink_to_fit=True))
        # sf.data_df.drop(columns=['add_index'],inplace=True)
    else:
        sf.set_column_width_dict(col_width_dict={('当事人', '诉讼代理人', '地址'): width})
    sf.to_excel(file,best_fit=sf.data_df.columns.difference(['当事人', '诉讼代理人', '地址']).tolist()).save()
    print_log('>>> 保存文件 => 文件名:%s => 数据保存成功...' %(file))
    return df

def rename_doc_by_infos(file):
    '''rename only judgment doc files'''
    try:doc = Document(file)
    except Exception as e:
        os.remove(file)
        print_log('读取判决书失败,格式不正确 => %s => 删除文件！！ %s '%(e,file))
        return
    for i,para in enumerate(doc.paragraphs[:10]):
        x = para.text
        if re.search(path_code_ix,x):
            x = case_codes_fix(x)
            jdoc_name = os.path.join(os.path.split(file)[0],'判决书_'+x+'.docx')
            if os.path.exists(jdoc_name):
                os.remove(jdoc_name)
            os.rename(file,jdoc_name)
            print_log('>>> 找到文件名,重命名判决书 => ',jdoc_name);break

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
        # print_log('-bool-isdict-%s-nokinadr-%s-kinuser-%s-nobyagent-%s--'%(type(n_adr) == dict,not k in adr,k in user,not by_agent))
        if type(n_adr) == dict and not k in adr and k in user and not by_agent:
            y += [k+adr_tag+n_adr.get(k)] # append address
    adr =  '，'.join(list(filter(None, y)))
    #  print_log('==new address==当事人 => %s 代理人 => %s 地址 => %s'%(user,agent,adr))
    return adr

def copy_rows_agent(x):
    '''copy example phone-numbers to agent column if empty'''
    x = x.astype(str)
    user = x[0];agent = x[1];adr = x[2]
    users = user_to_list(user)
    y = [agent]
    for i,u in enumerate(users):
        if not u in agent and u in adr:
            y += [u+'_123456']
    agent = ','.join(list(filter(None, y)))
    return agent

def copy_rows_users(new_adr,urows):
    '''copy users line regard adr user'''
    urow = ''
    if not isinstance(new_adr,dict):return urow
    user = list(new_adr.keys())
    urows = list(map(str, urows))
    for r in urows:
        if all(u in r for u in user):
            urow = r;break
    return urow

def get_jdocs_address(doc,lines = 20):# range from 20 lines
    '''get pre address from judgment docs, return docs pre code and address'''
    doc = Document(doc).paragraphs[:lines]
    adrs = {};number = ''
    for i,para in enumerate(doc):
        x = para.text.strip()
        if len(x) > 150: continue
        if re.search(path_code_ix,x) and len(x) < 25:
            number = x.strip().replace(' ','');continue # number
        if re.search(r'诉讼|代理|律师|请求|证据|辩|不服',x): continue # filter agent
        if re.search(r'(?<=[：:]).*?(?=[,，。])',x) and re.search(path_adr_ix,x):
            try:
                name = re.search(r'(?<=[：:]).*?(?=[,，。])|$',x).group(0).strip()
                name = name if re.search(r'[(（]曾用名.*?[）)]',x) else re.sub(r'[(（].*?[）)]','',name)
                z = re.split(r'[,，:：.。]',x)
                z = [re.sub(r'户籍|所在地|身份证|住所地|住址|现住|住','',y) for y in z if re.search(path_adr_ix,y)][-1] # 几个地址选最后一个 # remain only address
                adr = {name:''.join(z)}
                adrs.update(adr)
            except Exception as e:
                print_log('获取人名地址格式不正确 =>',e)
    return number,adrs

def get_all_docs(docs,df_orgin):
    numlist=[]; nadr = []
    for doc in docs:
        number,adrs = get_jdocs_address(doc)
        numlist.append(number)
        nadr.append(adrs)
    numlist = list(map(case_codes_fix,numlist))
    x = DataFrame({'原一审案号':numlist,'new_adr':nadr})
    return x

def add_jdoc_codes(d,r):
    '''add jdoc now case codes for reference'''
    if str(r['原一审案号_y']) in str(d):
        nd = os.path.join(os.path.split(d)[0],'判决书_'+str(r['案号']) +'_原_'+ str(r['原一审案号_y']) + '.docx')
        if(d == nd):return d
        if os.path.exists(nd):
            os.remove(nd)
        os.rename(d,nd)
        return nd
    return d

def rename_jdoc_codes(docs,df):
    '''rename jdoc now case codes for reference'''
    if not len(df):return docs 
    if not docs:return docs 
    for (i,r) in df[df!=''].iterrows():
        docs = list(map(lambda x:add_jdoc_codes(x,r),docs))
    return docs

def fill_infos(docs,df_orgin):
    '''combine address between data and judgment docs and delete duplicate'''
    global address_tmp
    x = get_all_docs(docs,df_orgin)
    x['当事人'] = x['new_adr'].apply(lambda x:copy_rows_users(x,df_orgin['当事人'].tolist()))
    try:
        address_tmp = x.copy()
        x.to_excel('address_tmp.xlsx',index=False)
        print_log('保存地址文件到 => address_tmp.xlsx...')
    except Exception as e:
        print_log('address_tmp.xlsx <= 保存失败,请检查...',e)
    x = merge(df_orgin,x,how='left',on=['当事人']).fillna('')
    x = x.drop_duplicates(['案号','当事人'])
    x['原一审案号_y'].replace('',float('nan'),inplace=True)
    x['原一审案号_y'] = x['原一审案号_y'].fillna(x['原一审案号_x'])
    rename_jdoc_codes(docs,x) # after rename then clean df codes
    
    x['原一审案号'] = x['原一审案号_x']
    x.drop(['原一审案号_x','原一审案号_y'],axis=1,inplace=True)
    xx = x.loc[:,['当事人','诉讼代理人','地址','new_adr']]
    xx['地址'] = xx.apply(lambda x:copy_rows_adr(x),axis=1)
    if flag_fill_phone: xx['诉讼代理人'] = xx.apply(lambda x:copy_rows_agent(x), axis=1)
    x[['地址','诉讼代理人']] =  xx[['地址','诉讼代理人']]
    x = x.drop(['new_adr'],axis=1)
    return x

#%% df process steps

def df_rename_jdocs():
    '''rename jdocs with contents'''
    if flag_rename_jdocs:
        docs = glob(parse_subpath(jdocs_path,'*.docx'))
        if docs:
            print_log('>>> 正在重命名判决书...')
            for file in docs:
                if not re.search(r'判决书_（[0-9]+）.*?号',file):
                    rename_doc_by_infos(file)
            print_log('>>> 重命名完毕...')

def df_fix_data(df_orgin):
    '''fix codes and remove nan'''
    df_orgin[['立案日期','案号','主审法官','当事人']] = df_orgin[['立案日期','案号','主审法官','当事人']].replace('',float('nan'))
    df_orgin.dropna(how='any',subset=['立案日期','案号','主审法官','当事人'],inplace=True)
    df_orgin['原一审案号'] = df_orgin['原一审案号'].fillna('')
    df_orgin[['案号','原一审案号']] = df_orgin[['案号','原一审案号']].applymap(case_codes_fix)
    return df_orgin

def df_oa_append(df_orgin):
    '''main fill OA data into df data and mark new add'''
    if flag_append_oa:
        if os.path.exists(data_oa_xlsx):
            print_log('找到OA数据,开始追加数据到主表 data.xlsx ...')
            df_orgin = read_excel(data_xlsx,sort=False)
            df_oa = read_excel(data_oa_xlsx,sort=False).tail(OA_last_lines)[titles_oa]
            df_oa.rename(columns={'承办人':'主审法官'},inplace=True)
            df1 = df_orgin.copy()
            df2 = df_oa
            df2['add_index'] = 'new'
            df1['add_index'] = 'old'
            df3 = df1.append(df2,sort=False).drop_duplicates(['立案日期','案号'],keep='first')
            if any(df3['add_index'] == 'new'):
                for (i,r) in df3[df3['add_index'] == 'new'].iterrows():
                    print_log('>>> 添加新OA记录 => ',r.tolist()[:3])
                df3 = titles_resort(df3,titles_cn)
                try:df_orgin = save_adjust_xlsx(df3,data_xlsx)
                except PermissionError: input('>>> data.xlsx 在其他地方打开...请手动关闭并重新执行...任意键退出');sys.exit()
    return df_orgin

def df_infos_fill(df_orgin):
    '''main fill jdocs infos'''
    docs = glob(parse_subpath(jdocs_path,'判决书_*.docx'))
    if flag_fill_jdocs_adr:
        if docs:
            # print_log('>>> 找到判决书 => %s' % docs)
            df_new = fill_infos(docs,df_orgin)
            df_new = titles_resort(df_new,titles_cn)
            try:df_new = save_adjust_xlsx(df_new,data_xlsx)
            except PermissionError: input('>>> 填充判决书 >>> data.xlsx 在其他地方打开...请手动关闭并重新执行...任意键退出');sys.exit()
            return df_new
    print_log('>>> 不填充判决书地址 ,继续下一步...')
    return df_orgin

def df_make_subset(df_orgin):
    '''cut orgin data into subset df data'''
    df = titles_combine_en(df_orgin,titles_cn)
    if data_last_lines:
        df = df.tail(data_last_lines) if cut_tail_lines else df.head(data_last_lines)
    elif ':' in date_range:
        df['datetime'] = to_datetime(df['datetime'])
        df.sort_values(by=['datetime'],inplace=True)
        dats = date_range.split(':')
        dats[0] = parse_datetime(dats[0]);dats[1] = parse_datetime(dats[1]);
        date_start = dats[0] if dats[0] else df['datetime'][0]
        date_end = dats[1] if dats[1] else df['datetime'].iloc[-1]
        df = df[(df['datetime']>date_start)&(df['datetime']<date_end)]
        df['datetime'] = df['datetime'].astype(str)
    return df

#%% main processing stream
print_log('>>> 正在读取记录...')
try:
    if not os.path.exists(cfgfile):
        write_config(cfgfile)
    else:
        cfg = configparser.RawConfigParser(allow_no_value=True)
        cfg.read_file(codecs.open(cfgfile, "r", "utf-8-sig"))
        data_xlsx = cfg.get('config', 'data_xlsx')
        data_oa_xlsx = cfg.get('config', 'data_oa_xlsx')
        sheet_docx = cfg.get('config', 'sheet_docx')
        OA_last_lines = cfg.getint('config', 'OA_last_lines')
        data_last_lines = cfg.getint('config', 'data_last_lines')
        date_range = cfg.get('config', 'date_range')
        flag_rename_jdocs = cfg.getboolean('config', 'flag_rename_jdocs')
        flag_fill_jdocs_adr = cfg.getboolean('config', 'flag_fill_jdocs_adr')
        flag_fill_phone  = cfg.getboolean('config', 'flag_fill_phone')
        flag_to_postal = cfg.getboolean('config', 'flag_to_postal')
        flag_check_data = cfg.getboolean('config', 'flag_check_data')
        flag_append_oa = cfg.getboolean('config', 'flag_append_oa')
except:
    print_log('配置文件出错,已重新生成...')
    os.remove(cfgfile)
    write_config(cfgfile)

if os.path.exists(data_xlsx):
    df_rename_jdocs()
    df_orgin = read_excel(data_xlsx,sort=False)
    df_orgin = df_fix_data(df_orgin)
    df_orgin = df_oa_append(df_orgin)
    df_orgin = df_infos_fill(df_orgin)
    df = df_make_subset(df_orgin)
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


def check_format(column,check=False):
    '''check data address and agent format with check flag'''
    if column.any():
        if check:
            if column.name == 'address':
                err = column[(column.str.strip().str.len()>2) & (column.str.contains(r'\/地址(\:|：)')==False)]
                if err.size > 0:
                    for i,item in enumerate(err): print_log('>>> excel line: %s val: \'%s\'  => 地址格式有误 => 缺少\'xxx/地址:xxx\' 如无请留空' %(i+2,item))
                    input('...请再次手动填充...任意键退出');sys.exit()
            if column.name == 'aname':
                err = column[(column.str.strip().str.len()>2) & (column.str.contains(r'\/|\_')==False)]
                if err.size > 0:
                    for i,item in enumerate(err): print_log('>>> excel line: %s val: \'%s\'  => 代理人格式有误 => 必须\'人名_xxx\' 如无请留空' %(i+2,item))
                    input('...请再次手动填充...任意键退出');sys.exit()
        return True
    return False

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
    dd_l = agent['uname'].str.strip().str.split(r'\、',expand=True).stack().to_frame(name = 'uname').reset_index()
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
    glist = tg['uname'].str.split(r'\、',expand=True).stack().values.tolist()
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
        print_log('>>> 生成新数据...start')
        number = df[titles_en[:4]]
        number = number.reset_index()
        number.columns.values[0] = 'level_0'
        # user
        user = df['uname']
        user = user[user != '']
        user = user.str.strip().str.split(r'[,，。]',expand=True).stack() # divide user
        user = user.str.strip().str.split(r'[:]',expand=True)# divide character
        user = user[1].str.strip().str.split(r'[、]',expand=True).stack().to_frame(name = 'uname');user #DataFrame
        user = user.reset_index().drop(['level_1','level_2'],axis=1)
        # agent and address
        agent_adr = df[['aname','address']]
        agent = df['aname']
        adr = df['address']

        if agent_adr.apply(lambda x: check_format(x,flag_check_data), axis=0).all():
            print_log('有 代理人和 有 地址...正在处理...')
            adr = make_adr(adr)
            agent = make_agent(agent,adr['clean_aname'].tolist())
            agent = merge_user(user,agent)
            df_x = reclean_data(merge_agent_adr(agent,adr))
            df_x = sort_data(df_x)
        elif check_format(adr,flag_check_data):
            print_log('无 代理人和 有 地址...正在处理...')
            adr = make_adr(adr,user['uname'].tolist())
            adr['uname'] = adr['clean_aname']
            adr = merge_user(user,adr)
            adr = adr.assign(aname='')
            df_x = reclean_data(adr)
            df_x = sort_data(df_x)
        elif check_format(agent,flag_check_data):
            print_log('有 代理人和 无 地址...正在处理...')
            agent = make_agent(agent)
            agent = merge_user(user,agent)
            agent = agent.assign(address='')
            df_x = reclean_data(agent)
            df_x = sort_data(df_x)
        else:
            print_log('无 代理人和 无 地址...正在处理...')
            agent_adr.index.name = 'level_0'
            agent_adr.reset_index(inplace=True)
            df_x = merge(user,agent_adr,how='left',on=['level_0']).fillna('')
            df_x = sort_data(df_x)

        del user,agent,number,adr,df,agent_adr

        if len(df_x):
            data_tmp = os.path.splitext(data_xlsx)[0]+"_temp.xlsx"
            df_save = df_x.copy()
            df_save.columns = titles_trans(df_save.columns.tolist())
            try:df_save = save_adjust_xlsx(df_save,data_tmp,width=40)
            except PermissionError: input('>>> data_tmp.xlsx 在其他地方打开...请手动关闭并重新执行...任意键退出');sys.exit()
            print_log('>>> 完成生成新数据...',df_save.iloc[:3,:5])

    except Exception as e:
        print_log('>>> 错误 %s 生成数据失败,请检查源 %s 文件...退出...'%(e,data_xlsx));sys.exit()
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
        text = re.sub(r'(?<=\s)\w+',address_text,para.text)
        para.clear().add_run(text)
    except Exception as e:
        print_log('错误 %s 替换文本 => %s 失败！！！' %(e,para.text))

    sheet_file = number_text+'_'+agent_text+'_'+user_text+'_'+address_text+'.docx'
    sheet_file = re.sub(r'[\/\\\:\*\?\"\<\>]',' ',sheet_file) # keep rename legal

    if os.path.exists(parse_subpath(postal_path,sheet_file)):
        print_log('>>> 邮单已存在！！！ <=',sheet_file)
        return sheet_file

    if not address_text:
        if flag_check_data:
            print_log('>>> %s 地址暂缺！！生成失败！！ <= %s'%(address_text,sheet_file))
        return sheet_file
    # if not search_names_phone(agent_text):
    #     if flag_check_data:
    #         print_log('>!> %s => 手机格式不对,请自行修改'% agent_text)
    try:
        doc.save(parse_subpath(postal_path,sheet_file))
        print_log('>>> 已生成邮单 =>',sheet_file)
    except Exception as e:
        print_log('>>> 生成失败！！！ =>',e)
    return sheet_file

if flag_to_postal:
    print_log('>>> 正在生成邮单...')
    if not os.path.exists(sheet_docx):
        input('>>> 没有找到邮单模板 %s...任意键退出' %sheet_docx);sys.exit()
    df_x.apply(re_write_text,axis = 1)
    print_log('>>> 邮单数据范围 => %s-------%s'%(df_x['number'].astype(str).iloc[0],df_x['number'].astype(str).iloc[-1]))
    # re_write_text(df_x.iloc[1])

#%% main finish
# if __name__ == "__main__":
input('>>> 全部完成,可以回顾记录...任意键退出')


