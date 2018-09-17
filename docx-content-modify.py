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

Created on Wed Aug 15 2018

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
#%% config and default values
cfgfile = 'config.txt'
data_xlsx = 'data.xlsx'
data_oa_xlsx = 'data_oa.xlsx'
sheet_docx = 'sheet.docx'
postal_path = os.path.join('.','postal')
jdocs_path = os.path.join('.','jdocs')
OA_last_lines = 10
data_last_lines = 10
date_range = '2018-06-01:2018-08-01'
rename_jdocs = False
fill_jdocs_adr = True
fill_phone_flag = True
append_data_flag = True
to_postal = False
tmp_file = True
cut_tail_lines = True
check_data_flag = True

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
        cfg.set('config', 'rename_jdocs', rename_jdocs)
        cfg.set('config', '# 是否填充判决书地址')
        cfg.set('config', 'fill_jdocs_adr', fill_jdocs_adr)
        cfg.set('config', '# 是否填充伪手机')
        cfg.set('config', 'fill_phone_flag', fill_phone_flag)
        cfg.set('config', '# 是否导入OA数据')
        cfg.set('config', 'append_data_flag', append_data_flag)
        cfg.set('config', '# 导入OA数据的最后几行')
        cfg.set('config', 'OA_last_lines', OA_last_lines)
        cfg.set('config', '# 是否打印邮单')
        cfg.set('config', 'to_postal', to_postal)
        cfg.set('config', '# 打印数据模板的最后几行')
        cfg.set('config', 'data_last_lines', data_last_lines)
        cfg.set('config', '# 打印数据模板的日期范围')
        cfg.set('config', 'date_range', date_range)
        cfg.set('config', '# 检查数据模板的内容格式')
        cfg.set('config', 'check_data_flag', check_data_flag)
        cfg.write(configfile)

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
        rename_jdocs = cfg.getboolean('config', 'rename_jdocs')
        fill_jdocs_adr = cfg.getboolean('config', 'fill_jdocs_adr')
        fill_phone_flag  = cfg.getboolean('config', 'fill_phone_flag')
        to_postal = cfg.getboolean('config', 'to_postal')
        check_data_flag = cfg.getboolean('config', 'check_data_flag')
        append_data_flag = cfg.getboolean('config', 'append_data_flag')
except:
    print('配置文件出错,已重新生成...')
    os.remove(cfgfile)
    write_config(cfgfile)
    
titles_cn = ['立案日期','案号','原一审案号','主审法官','当事人','诉讼代理人','地址']
titles_en = ['datetime','number','prenumber','judge','uname','aname','address']
titles_oa = ['立案日期','案号','原一审案号','承办人','当事人']
titles_cn2en = dict(zip(titles_cn, titles_en))
titles_en2cn = dict(zip(titles_en, titles_cn))

#%% read util
path_names_ix = re.compile(r'[^A-Za-z\u4e00-\u9fff（）()：:]') # remain only name including old name 包括括号冒号
path_code_ix = re.compile(r'(\(|\（)([0-9]+)(\)|\）).*?号', re.UNICODE) # case numbers
path_adr_ix = re.compile(r'住|(市(.*)[0-9])') # chinese address
path_adr_clean = re.compile(r'户籍|所在地|身份证|住所地|住址|现住|住') # remain only address
path_adr_cut = re.compile(r'\,|\，|\:|\：|\.|\。')
titles_adr = ['当事人','诉讼代理人','地址','new_adr']

def user_get_name(y):
    y = re.split(r'\:|、|,|，',y)
    return [x for x in y if not re.search(r'申请人|被申请人|原告|被告|原审被告|上诉人|被上诉人|第三人',x)]

def case_codes_fix(x):
    y = re.search(path_code_ix,x)
    if y: x = y.group()
    x = x.replace('(','（').replace(')','）').replace(' ','')
    return x

def parse_subpath(path,file):
    if not os.path.exists(path):
        os.mkdir(path)
    return os.path.join(path,file)

def check_contain_chinese(check_str):
    return any((u'\u4e00' <= char <= u'\u9fff') for char in check_str)

def parse_datetime(date):
    try:date = datetime.datetime.strptime(date,'%Y-%m-%d')
    except ValueError:print('时间范围格式有误,默认选取全部日期');date = ''
    return date

def titles_trans(df_list):
    '''Change titles between Chinese and English'''
    flag = check_contain_chinese(df_list[0])
    tdict = titles_cn2en if flag else titles_en2cn
    print('Change titles to:', 'Chinese' if flag else 'English')
    return [tdict.get(x) for x in df_list if x in tdict.keys()]

def titles_resort(df,titles):
    '''resort titles with orders'''
    titles_rest = df.drop(titles,axis=1).columns.tolist()
    return df[titles + titles_rest].fillna('')

def titles_combine_en(df,titles):
    '''refer to titles to sub-replace with columns titles'''
    titles_rest = df.drop(titles,axis=1).columns.tolist()
    df = df[titles + titles_rest]
    df.columns = titles_trans(titles) + titles_rest
    return df
    
def save_adjust_xlsx(df,file):
    '''save and re-adjust excel format'''
    df = df.astype(str)
    ew = StyleFrame.ExcelWriter(file)
    StyleFrame.A_FACTOR = 5
    StyleFrame.P_FACTOR = 1.2
    sf = StyleFrame(df,Styler(**{'wrap_text': False, 'shrink_to_fit':True, 'font_size': 12}))
    sf.set_column_width_dict(col_width_dict={('当事人', '诉讼代理人', '地址'): 80})
    sf.to_excel(excel_writer=ew,best_fit=df.columns.difference(['当事人', '诉讼代理人', '地址']).tolist()).save()
    print('>>> 保存文件 => 文件名:%s...列名:%s => 数据保存成功...' %(file,df.columns.tolist()))

def df_oa_append(df_orgin):
    '''fill OA data into df data'''
    if os.path.exists(data_oa_xlsx):
        print('找到OA数据,开始追加数据到主表 data.xlsx ...')
        df_oa = read_excel(data_oa_xlsx,sort=False).tail(OA_last_lines)[titles_oa]
        df_oa.rename(columns={"承办人": '主审法官'},inplace=True)
        df_orgin[['案号','原一审案号']] = df_orgin[['案号','原一审案号']].applymap(case_codes_fix)
        x = df_orgin.append(df_oa).drop_duplicates(['立案日期','案号'])
        x = x.sort_values('立案日期').groupby('案号').backfill().drop_duplicates(['立案日期','案号','当事人'])
        df_orgin = titles_resort(x,titles_cn)
    if append_data_flag:
        try:save_adjust_xlsx(df_orgin,data_xlsx)
        except PermissionError: input('>>> data.xlsx 在其他地方打开...请手动关闭并重新执行...任意键退出');sys.exit()
    return df_orgin

def df_make_subset(df_orgin):
    '''from orgin data to df data'''
    df = titles_combine_en(df_orgin,titles_cn)
    df['datetime'] = to_datetime(df['datetime'])
    if data_last_lines:
        df = df.tail(data_last_lines) if cut_tail_lines else df.head(data_last_lines)
    elif ':' in date_range:
        df.sort_values(by=['datetime'],inplace=True)
        dats = date_range.split(':')
        dats[0] = parse_datetime(dats[0]);dats[1] = parse_datetime(dats[1]);
        date_start = dats[0] if dats[0] else df['datetime'][0]
        date_end = dats[1] if dats[1] else df['datetime'].iloc[-1]
        df = df[(df['datetime']>date_start)&(df['datetime']<date_end)]
    return df

def rename_doc_by_infos(file):
    '''rename only judgment doc files'''
    # doc = Document(file)
    try:doc = Document(file)
    except Exception as e:
        print('读取判决书失败,可能格式不正确 =>',e);return
    for i,para in enumerate(doc.paragraphs[:10]):
        x = para.text
        if re.search(path_code_ix,x):
            x = case_codes_fix(x)
            jdoc_name = os.path.join(os.path.split(file)[0],'判决书_'+x+'.docx')
            if os.path.exists(jdoc_name):
                os.remove(jdoc_name)
            os.rename(file,jdoc_name)
            print('重命名判决书 => %s',jdoc_name);break
    print('%s <= 文件不是判决书，不重命名',file)

def renamef(x,y):
    '''Clean agent name for agent to match address's agent name'''
    if bool(y):
        for name in y:
            if not check_contain_chinese(name):continue
            if name in x: x = x.replace(x,name);break
    return x

def check_format(column,check=False):
    if column.any(): # column.map(type).eq(str).any()|and column.fillna('').apply(check_contain_chinese).any(): # 有中文信息才合格
        if check:
            if column.name == 'address':
                err = column[(column.str.strip().str.len()>2) & (column.str.contains(r'\/地址(\:|：)')==False)]
                if err.size > 0:
                    for i,item in enumerate(err): print('>>> excel line: %s val: \'%s\'  => 地址格式有误 => 缺少\'xxx/地址:xxx\' 如无请留空' %(i+2,item))
                    input('...请再次手动填充...任意键退出');sys.exit()
            if column.name == 'aname':
                err = column[(column.str.strip().str.len()>2) & (column.str.contains(r'\/|\_')==False)]
                if err.size > 0:
                    for i,item in enumerate(err): print('>>> excel line: %s val: \'%s\'  => 代理人格式有误 => 必须\'人名_xxx\' 如无请留空' %(i+2,item))
                    input('...请再次手动填充...任意键退出');sys.exit()
        return True
    return False

#%% rename judgment docs
def df_rename_jdocs():
    if rename_jdocs:
        docs = glob(parse_subpath(jdocs_path,'*.docx'))
        if len(docs)>0:
            print('>>> 正在重命名判决书...')
            for file in docs:
                if not '判决书' in file:
                    rename_doc_by_infos(file)
            print('>>> 重命名完毕...')

#%% merge address from judgment docs (Chinese Format)
def get_pre_address(doc,userlist,lines = 20):
    '''get pre address from judgment docs, return docs pre code and address'''
    doc = Document(doc).paragraphs[:lines] # range from 30 lines
    adrs = {};number = '';users = ''
    for i,para in enumerate(doc):
        x = para.text.strip()
        if len(x) > 150: continue
        if re.search(path_code_ix,x) and len(x) < 25:
            number = x.strip().replace(' ','');continue # number
        if re.search('诉讼|代理|律师',x): continue # filter agent
        if re.search(path_adr_ix,x) and not re.search(path_code_ix,x):
            try:
                name = re.split(r'\,|\，|。',re.split(r'\:|\：',x)[1])[0].strip() # 冒号后第二个是人名
                alist = re.split(path_adr_cut,x)
                alist = [re.sub(path_adr_clean,'',x) for x in alist if re.search(path_adr_ix,x)][-1] # 几个地址选最后一个
                adr = {name:''.join(alist)}
                adrs.update(adr)
                # print('获取=>',adrs)
            except Exception as e:
                print('获取人名地址格式不正确 =>',e)
    users = ''.join([yy for yy in userlist if all(str(l) in yy for l in list(adrs.keys()))])
    print('找到地址===%s===%s===%s'%(number,users,adrs))
    return number,adrs,users

def copy_rows_adr(x):
        user = x[0];agent = x[1];adr = x[2];n_adr = x[3]
        if n_adr: 
            y = [adr]
            print('==find address=',n_adr)
            for i,k in enumerate(n_adr):
                # check records from user,agent and address
                by_agent = ((k in agent) if re.search(r'(.+\、)*.+\/[\u4e00-\u9fff]+',agent) else False) 
                # print('------bool--%s-%s-%s--%s-'%(type(n_adr) == dict,k in user,not by_agent,not k in adr))
                if type(n_adr) == dict and k in user and not by_agent and not k in adr:
                    y += [k+'/地址：'+n_adr.get(k)]
            adr = ','.join(list(filter(None, y)))
            # print('==find address=',y)
            print('==new address==当事人 => %s 代理人 => %s 地址 => %s'%(user,agent,adr))
        return adr
    
def copy_rows_agent(x):
    user = x[0];agent = x[1];adr = x[2]
    users = user_get_name(user)
    y = [agent]
    for i,u in enumerate(users):
        if not u in agent and u in adr:
            y += [u+'_123456']
    agent = ','.join(list(filter(None, y)))
    return agent

def fill_infos(docs,df_orgin):
    '''combine address between data and judgment docs and delete duplicate'''
    numlist=[]; nadr = []; userlist = []
    for doc in docs:
        number,adrs,users = get_pre_address(doc,df_orgin['当事人'].tolist())
        numlist.append(number)
        nadr.append(adrs)
        userlist.append(users)
    numlist = list(map(case_codes_fix,numlist))
    x = DataFrame({'原一审案号':numlist,'当事人':userlist,'new_adr':nadr})
    x.to_excel('address_tmp.xlsx',index=False)
    x = merge(df_orgin,x,how='left',on=['当事人'])
    x = x.drop_duplicates(['案号','当事人'])
    x['原一审案号'] = x['原一审案号_x']
    x.drop(['原一审案号_x','原一审案号_y'],axis=1,inplace=True)
    # x = x.sort_values('立案日期').groupby('当事人').backfill()
    xx = x[titles_adr].fillna('')
    xx['地址'] = xx.apply(lambda x:copy_rows_adr(x), axis=1)
    if fill_phone_flag: xx['诉讼代理人'] = xx.apply(lambda x:copy_rows_agent(x), axis=1)
    x[['地址','诉讼代理人']] =  xx[['地址','诉讼代理人']]
    x = x.drop(['new_adr'],axis=1)
    return x

def df_infos_fill(df_orgin):
    docs = glob(parse_subpath(jdocs_path,'判决书_*.docx'))
    if len(docs)>0:
        print('>>> 找到判决书 => %s' % docs)
        df_new = fill_infos(docs,df_orgin)
        df_new = titles_resort(df_new,titles_cn)
    else: print('>>> 没有找到判决书docx,可复制判决书到jdocs目录....继续下一步')
    
    if fill_jdocs_adr:
        try:save_adjust_xlsx(df_new,data_xlsx)
        except PermissionError: input('>>> data.xlsx 在其他地方打开...请手动关闭并重新执行...任意键退出');sys.exit()
        print('>>> 填充地址到原文件完毕...请手动填充%s的代理人...自动继续...' % data_xlsx)
        return df_new
    else: print('设定不填充地址 ,继续下一步...'); return df_orgin
    
#%% read xlsx
print('>>> 正在读取记录...')
if os.path.exists(data_xlsx):
    df_rename_jdocs()
    df_orgin = read_excel(data_xlsx,sort=False)
    df_orgin[['案号','原一审案号']] = df_orgin[['案号','原一审案号']].applymap(case_codes_fix)
    df_orgin = df_oa_append(df_orgin)
    df_orgin = df_infos_fill(df_orgin)
    df = df_make_subset(df_orgin)
    print('>>> 读取记录成功...')
else: input('>>> data.xlsx 记录文件不存在...任意键退出');sys.exit()
#%% No.
print('>>> 生成新数据...start')
number = df[titles_en[:4]]
number = number.reset_index()
number.columns.values[0] = 'level_0'
# user
user = df['uname'];user
user = user.str.strip().str.split(r'\,|\，|\n|\t',expand=True).stack();user # divide user    
user = user.str.strip().str.split(r'\:',expand=True);user # divide character
user = user[1].str.strip().str.split(r'\、',expand=True).stack().to_frame(name = 'uname');user #DataFrame
user = user.reset_index().drop(['level_1','level_2'],axis=1)
# agent and address
def make_adr(adr,clean_aname=[]):
    adr = df['address'].fillna('').replace(re.compile(r'nan.+'),'')
    adr = adr[adr != '']
    adr = adr.str.strip().str.split(r'\,|\，|\n|\t',expand=True).stack()
    adr = adr.str.strip().apply(lambda x: x if '/' in x else ( (x.split('_')[0]+'/地址:'+x) if '_' in x else x+'/地址:'))
    adr = adr.str.strip().str.split(r'\/',expand=True).fillna('')  
    adr.columns = ['aname','address']
    adr['address'] = adr['address'].apply(lambda x:re.sub(r'地址(\:|\：)','',x))# drop '地址:'
    adr['clean_aname'] = adr['aname'].str.strip().apply(lambda x: renamef(x,clean_aname)).apply(lambda x: re.sub(r'\_(.*)','',x)).replace(path_names_ix,'') # clean adr
    adr = adr.reset_index().drop(['level_1','aname'],axis=1)
    return adr

def make_agent(agent,clean_aname=[]):
    agent = df['aname'].fillna('').replace(re.compile(r'nan.+'),'')
    agent = agent[agent != '']
    agent = agent.str.strip().str.split(r'\,|\，|\n|\t',expand=True).stack() #Series
    agent = agent.str.strip().apply(lambda x: x if '/' in x else ( x.split('_')[0]+'/'+x if '_' in x else x+'/'))
    agent = agent.str.strip().str.split(r'\/',expand=True).fillna('') #DataFrame
    agent.columns = ['uname','aname']
    agent['clean_aname'] = agent['aname'].str.strip().apply(lambda x: renamef(x,clean_aname)).apply(lambda x:re.sub(r'\_(.*)','',x)).replace(path_names_ix,'')
    dd_l = agent['uname'].str.strip().str.split(r'\、',expand=True).stack().to_frame(name = 'uname').reset_index();dd_l
    dd_r = agent[agent.columns.difference(['uname'])].reset_index();dd_r
    agent = merge(dd_l,dd_r,how='outer',on=['level_0','level_1']).drop(['level_1','level_2'],axis=1)
    return agent

def merge_user(user,agent):
    return merge(user,agent,how='left',on=['level_0','uname'])

def merge_agent_adr(agent,adr):
    agent['clean_aname'].replace('',float("nan"),inplace=True)
    agent['clean_aname'] = agent['clean_aname'].fillna(agent['uname']).replace(path_names_ix,'')
    adr['clean_aname'] = adr['clean_aname'].apply(lambda x: renamef(x,agent['clean_aname'].tolist()))
    tb = merge(agent,adr,how='outer',on=['level_0','clean_aname'])
    return tb

def reclean_data(tb):
    tg = tb.fillna('').groupby(['level_0','clean_aname','aname','address'])['uname'].apply(lambda x: '、'.join(x.astype(str))).reset_index()
    glist = tg['uname'].str.split(r'\、',expand=True).stack().values.tolist()
    rest = tb[tb['uname'].isin(glist) == False]
    x = concat([rest,tg],axis=0,sort=True)
    return x

def sort_data(x):
    x = x[['level_0','uname','aname','address']].sort_values(by=['level_0']).fillna('***').replace('','***')
    x = merge(number,x,how='right',on=['level_0']).drop(['level_0'],axis=1)
    return x

agent_adr = df[['aname','address']]
agent = df['aname']
adr = df['address']

if agent_adr.apply(lambda x: check_format(x,check_data_flag), axis=0).all():
    print('有 代理人和 有 地址...正在处理...')
    adr = make_adr(adr)
    agent = make_agent(agent,adr['clean_aname'].tolist())
    agent = merge_user(user,agent)
    df_x = reclean_data(merge_agent_adr(agent,adr))
    df_x = sort_data(df_x)
elif check_format(adr,check_data_flag):
    print('无 代理人和 有 地址...正在处理...')
    adr = make_adr(adr,user['uname'].tolist())
    adr['uname'] = adr['clean_aname']
    adr = merge_user(user,adr)
    adr = adr.assign(aname='')
    df_x = reclean_data(adr)
    df_x = sort_data(df_x)
elif check_format(agent,check_data_flag):
    print('有 代理人和 无 地址...正在处理...')
    agent = make_agent(agent)
    agent = merge_user(user,agent)
    agent = agent.assign(address='')
    df_x = reclean_data(agent)
    df_x = sort_data(df_x)
else:
    print('无 代理人和 无 地址...正在处理...')
    agent_adr.index.name = 'level_0'
    agent_adr.reset_index(inplace=True)
    df_x = merge(user,agent_adr,how='left',on=['level_0'])
    df_x = sort_data(df_x)

print('>>> 完成生成新数据...end => data')
#%% check_data_flag and save data
if check_data_flag:
    if '***' in df_x[['aname','address']].values:
        print('>>> 提示:代理人和地址还有空缺...不影响继续生成')
        
if tmp_file and len(df_x):
    data_tmp = os.path.splitext(data_xlsx)[0]+"_temp.xlsx"
    df_save = df_x.copy()
    df_save.columns = titles_trans(df_save.columns.tolist())
    #data.to_excel(data_tmp,index=False)
    try:save_adjust_xlsx(df_save,data_tmp)
    except PermissionError: input('>>> data.xlsx 在其他地方打开...请手动关闭并重新执行...任意键退出');sys.exit()
else: input('>>> 生成数据失败,请检查源data.xlsx文件...退出');sys.exit()

#%% generate postal sheets

def re_writ_text(x):
    doc = Document(sheet_docx)
    doc.styles['Normal'].font.bold = True
    
    is_star = lambda x: x == '***'
    clean_text = lambda x,y: (y+'#'+'暂缺') if is_star(x) else x
    agent_text = clean_text(x['aname'],'代理') if not is_star(x['aname']) else clean_text(x['uname'],'当事人')
    user_text = '代 '+ clean_text(x['uname'],'当事人') if not is_star(x['uname']) and not x['uname'] in agent_text else ''
    address_text = clean_text(x['address'],'地址')
    number_text = clean_text(x['number'],'案号')
    
    try:
        paragraph = doc.paragraphs[9]  # No.9 line is agent name
        text = re.sub('([\u4e00-\u9fff]).*',agent_text,paragraph.text)
        paragraph.clear().add_run(text)
        
        paragraph = doc.paragraphs[11]  # No.11 line is user name
        text = re.sub('([\u4e00-\u9fff]).*',user_text,paragraph.text)
        paragraph.clear().add_run(text)
        
        paragraph = doc.paragraphs[13]  # No.13 line is number and address
        text = re.sub(path_code_ix,number_text,paragraph.text)
        paragraph.clear().add_run(text)
        
        path_code_adr_ix = re.compile(r'\s([\u4e00-\u9fff]+.*)', re.UNICODE)
        text = re.sub(re.findall(path_code_adr_ix,paragraph.text)[0],address_text,paragraph.text)
        paragraph.clear().add_run(text)
    except:
        print('替换文本 => %s 失败',paragraph.text)

    sheet_file = number_text+'_'+agent_text+'_'+user_text+'_'+address_text+'.docx'
    if '暂缺' in address_text:
        # print('%s 暂缺不生成 <= %s'%(address_text,sheet_file))
        pass
    elif '暂缺' in agent_text:
        # print('%s 暂缺不生成 <= %s'%(agent_text,sheet_file))
        pass
    elif not re.search(r'[A-z\u4e00-\u9fff]+\_\d+',agent_text):
        print('%s 手机格式不对,不生成 <= %s'%(agent_text,sheet_file))
        pass
    else:
        doc.save(parse_subpath(postal_path,sheet_file))
        print('>>> 已生成邮单 =>',sheet_file)
    return doc

if to_postal:
    print('>>> 正在生成邮单...')
    if not os.path.exists(sheet_docx):
        input('>>> 没有找到邮单模板%s...任意键退出' % sheet_docx);sys.exit()
    df_x.apply(re_writ_text,axis = 1)
    print('>>> 邮单数据范围 => %s-------%s'%(df['number'].astype(str).iloc[0],df['number'].astype(str).iloc[-1]))
    #re_writ_text(df_x.iloc[1])

#%% main
# if __name__ == "__main__":
input('>>> 全部完成,可以回顾记录...任意键退出')
    


