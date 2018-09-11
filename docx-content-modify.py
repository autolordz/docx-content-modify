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
from docx import Document  # doc = Document('*.docx')
from glob import glob
from StyleFrame import StyleFrame

set_option('max_colwidth',500)
set_option('max_rows', 50)
set_option('max_columns',50)
#%% config and default values
cfgfile = 'config.txt'
data_xlsx = 'data.xlsx'
data_orgin = 'data_orgin.xlsx'
sheet_docx = 'sheet.docx'
postal_path = os.path.join('.','postal')
jdocs_path = os.path.join('.','jdocs')
OA_last_lines = 10
data_last_lines = 5
date_range = '2018-06-01:2018-08-01'
rename_jdocs = True
fill_jdocs_adr = True
append_data_flag = False
to_postal = True
check_data_flag = True
tmp_file = True
cut_tail_lines = False

cfg = configparser.RawConfigParser(allow_no_value=True)
cfg.add_section('config')
cfg.set('config', '# 数据模板地址')
cfg.set('config', 'data_xlsx', data_xlsx)
cfg.set('config', '# OA数据地址')
cfg.set('config', 'data_orgin', data_orgin)
cfg.set('config', '# 邮单模板地址')
cfg.set('config', 'sheet_docx', sheet_docx)
cfg.set('config', '# 是否重命名判决书')
cfg.set('config', 'rename_jdocs', rename_jdocs)
cfg.set('config', '# 是否填充判决书地址')
cfg.set('config', 'fill_jdocs_adr', fill_jdocs_adr)
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


def write_config(cfg,cfgfile):
     # with open(cfgfile, 'w') as configfile:
        # cfg.write(configfile)
     with codecs.open(cfgfile, "w", "utf-8-sig") as configfile:
        cfg.write(configfile)

try:
    if not os.path.exists(cfgfile):
        write_config(cfg,cfgfile)
    else:
        cfg.read_file(codecs.open(cfgfile, "r", "utf-8-sig"))
        data_xlsx = cfg.get('config', 'data_xlsx')
        data_orgin = cfg.get('config', 'data_orgin')
        sheet_docx = cfg.get('config', 'sheet_docx')
        OA_last_lines = cfg.getint('config', 'OA_last_lines')
        data_last_lines = cfg.getint('config', 'data_last_lines')
        date_range = cfg.get('config', 'date_range')
        rename_jdocs = cfg.getboolean('config', 'rename_jdocs')
        fill_jdocs_adr = cfg.getboolean('config', 'fill_jdocs_adr')
        to_postal = cfg.getboolean('config', 'to_postal')
        check_data_flag = cfg.getboolean('config', 'check_data_flag')
        append_data_flag = cfg.getboolean('config', 'append_data_flag')
except:
    print('配置文件出错,已重新生成...')
    write_config(cfg,cfgfile)
    
titles_cn = ['立案日期','案号','原一审案号','主审法官','当事人','诉讼代理人','地址']
titles_en = ['datetime','number','prenumber','judge','uname','aname','address']
titles_oa = ['立案日期','案号','原一审案号','承办人','当事人']
titles_cn2en = dict(zip(titles_cn, titles_en))
titles_en2cn = dict(zip(titles_en, titles_cn))

#%% read util
pat_chinese = lambda :re.compile(r'\?|\.|\。|\!|\/|\;|\:|\*|\>|\<|\~|\(|\)|\[|\]|[0-9]|') # remain only Chinese

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

def titles_combine_en(df,titles):
    '''refer to titles to sub-replace with columns titles'''
    titles_rest = df.drop(titles,axis=1).columns.tolist()
    df = df[titles + titles_rest]
    df.columns = titles_trans(titles) + titles_rest
    return df
    
def save_adjust_xlsx(df,file,titles):
    '''save and re-adjust excel format'''
    df = df.astype(str)
    ew = StyleFrame.ExcelWriter(file)
    StyleFrame.A_FACTOR = 4
    StyleFrame.P_FACTOR = 1.2
    sf = StyleFrame(df)
    sf.to_excel(excel_writer=ew,best_fit=titles).save()
    print('>>> 保存文件 => 文件名:%s...列名:%s => 数据保存成功...' %(file,titles))

def df_append(df):
    '''fill OA data into df data'''
    if append_data_flag:
        if os.path.exists(data_orgin):
            print('找到OA数据,开始追加数据到主表 data.xlsx ...')
            df_oa = read_excel(data_orgin,sort=False).tail(OA_last_lines)[titles_oa]
            df_oa.rename(columns={"承办人": '主审法官'},inplace=True)
            df = df.append(df_oa).drop_duplicates(['立案日期','案号']).sort_values(by=['案号'])[df.columns]
            try:save_adjust_xlsx(df,data_xlsx,titles_cn)
            except PermissionError: input('>>> data.xlsx 在其他地方打开...请手动关闭并重新执行...任意键退出');sys.exit()
    return df

def make_df(df_orgin):
    '''from orgin data to df data'''
    df_orgin['datetime'] = to_datetime(df_orgin['datetime'])
    if data_last_lines:
        df = df_orgin.tail(data_last_lines) if cut_tail_lines else df_orgin.head(data_last_lines)
    elif ':' in date_range:
        df_orgin.sort_values(by=['datetime'],inplace=True)
        dats = date_range.split(':')
        dats[0] = parse_datetime(dats[0]);dats[1] = parse_datetime(dats[1]);
        date_start = dats[0] if dats[0] else df_orgin['datetime'][0]
        date_end = dats[1] if dats[1] else df_orgin['datetime'].iloc[-1]
        df = df_orgin[(df_orgin['datetime']>date_start)&(df_orgin['datetime']<date_end)]
    return df

def rename_doc_by_infos(file):
    '''rename only judgment doc files'''
    doc = Document(file)
    for i,paragraph in enumerate(doc.paragraphs[:10]):
        if re.search(r'民(申|终)\d+号',paragraph.text):
            os.rename(file,os.path.join(os.path.split(file)[0],'判决书_'+paragraph.text+'.docx'))


def renamef(x,y):
    '''Clean agent name for agent to match address's agent name'''
    if bool(y):
        for name in y:
            if not check_contain_chinese(name):continue
            if name in x: x = x.replace(x,name);break
    return x

def check_format(column,check=False):
    if check:
        if column.map(type).eq(str).any(): # and column.fillna('').apply(check_contain_chinese).any(): # 有中文信息才合格
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
if rename_jdocs:
    docs = glob(parse_subpath(jdocs_path,'*.docx'))
    if len(docs)>0:
        print('>>> 正在重命名判决书...')
        for file in docs:
            if not '判决书' in file:
                rename_doc_by_infos(file)
        print('>>> 重命名完毕...')
#%% read xlsx
print('>>> 正在读取记录...')
if os.path.exists(data_xlsx):
    df_orgin = read_excel(data_xlsx,sort=False)
    df_orgin = df_append(df_orgin)
    df = titles_combine_en(df_orgin,titles_cn)
    df = make_df(df)
    print('>>> 读取记录成功...',df.tail(5))
else: input('>>> data.xlsx 记录文件不存在...任意键退出');sys.exit()
#%% merge address from judgment docs (Chinese Format)

path_code_ix = re.compile(r'(\(|\（)([0-9]+)(\)|\）)(.*)号')
path_adr_ix = re.compile(r'住|(市(.*)[0-9])')
path_adr_clean = re.compile(r'户籍|所在地|身份证|住所地|住址|住') # remain only address
path_adr_cut = re.compile(r'\,|\，|\:|\：|\.|\。')
titles_adr = ['当事人','诉讼代理人','地址','new_adr']

def get_pre_address(doc,lines = 30):
    '''get pre address from judgment docs, return docs pre code and address'''
    doc = Document(doc).paragraphs[:lines]
    adrs = {};number = ''
    for i,para in enumerate(doc): # range from 3:10 lines
        if len(para.text) > 100: continue
        # if not '肖汉城' in para.text:continue
        # print('===docline=%s=%s='%(i,para.text))
        if re.search(path_code_ix,para.text) and len(para.text) < 25:number = para.text.strip();continue # number
        if re.search('诉讼|代理|律师',para.text): continue # filter agent
        # print('cccccccccc', re.search(path_code_ix,para.text))
        if re.search(path_adr_ix,para.text) and not re.search(path_code_ix,para.text):
            # cc = para.text;  print('ccccccc', cc)
            alist = re.split(path_adr_cut,para.text)
            adr = {alist[1]:''.join(re.sub(path_adr_clean,'',x) for x in alist if re.search(path_adr_ix,x))}
            # if adr: print('找到地址==%s==%s='%(number,adr))
            adrs.update(adr)
    return number,adrs
        
# numlist=[]; nadr = []
# for doc in docs:
#     number,adrs = get_pre_address(doc)
#     numlist.append(number)
#     nadr.append(adrs)
dfs = DataFrame()
def fill_duplicate_adr(docs,df_orgin):
    '''combine address between data and judgment docs and delete duplicate'''
    numlist=[]; nadr = []
    for doc in docs:
        number,adrs = get_pre_address(doc)
        numlist.append(number)
        nadr.append(adrs)
    dfs = DataFrame({'原一审案号':numlist,'new_adr':nadr})
    dfs.to_excel('address_tmp.xlsx',index=False)
    dfn = merge(df_orgin,dfs,how='left',on=['原一审案号'])
    dfn_adr = dfn[titles_adr].fillna('')
    def copy_rows_adr(x):
        user = x[0];agent = x[1];adr = x[2];n_adr = x[3]
        if n_adr: 
            xx = [adr]
            for i,k in enumerate(n_adr):
                # check records from user,agent and address
                if type(n_adr) == dict and k in user and not k+'/' in agent and not k in adr:
                    xx += [k+'/地址：'+n_adr.get(k)]
            xx = list(filter(None, xx))
            print('==find address=',xx)
            adr = ','.join(xx)
            print('==new address==当事人 => %s 代理人 => %s 地址 => %s====='%(user,agent,adr))
        return adr
    dfn['地址'] = dfn_adr.apply(lambda x:copy_rows_adr(x), axis=1)
    dfn.drop(['new_adr'],axis=1,inplace=True)
    return dfn

# docs = glob(parse_subpath(jdocs_path,'判决书_*.docx'))
# dfn = fill_duplicate_adr(docs,df_orgin)
        
if fill_jdocs_adr:
    docs = glob(parse_subpath(jdocs_path,'判决书_*.docx'))
    if len(docs)>0:
        print('>>> 找到判决书 => %s...正在填充判决书上地址...' % docs)
        dfn = fill_duplicate_adr(docs,df_orgin)
        try:save_adjust_xlsx(dfn,data_xlsx,titles_cn)
        except PermissionError: input('>>> data.xlsx 在其他地方打开...请手动关闭并重新执行...任意键退出');sys.exit()
        input('>>> 填充地址到原文件完毕...请手动填充%s的代理人...任意键继续' % data_xlsx)
    else: input('>>> 没有找到判决书docx,可复制判决书到jdocs目录...任意键继续')
else: print('config 选择不填充地址 ,继续下一步...')

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
    adr = adr.str.strip().str.split(r'\/',expand=True).fillna('')  
    adr.columns = ['aname','address']
    adr['address'] = adr['address'].apply(lambda x:re.sub(r'地址(\:|\：)','',x))# drop '地址:'
    adr['clean_aname'] = adr['aname'].str.strip().apply(lambda x: renamef(x,clean_aname)).apply(lambda x: re.sub(r'\_(.*)','',x)).replace(pat_chinese(),'') # clean adr
    adr = adr.reset_index().drop(['level_1','aname'],axis=1)
    return adr

def make_agent(agent,clean_aname=[]):
    agent = df['aname'].fillna('').replace(re.compile(r'nan.+'),'')
    agent = agent[agent != '']
    agent = agent.str.strip().str.split(r'\,|\，|\n|\t',expand=True).stack() #Series    
    agent = agent.str.strip().str.split(r'\/',expand=True).fillna('') #DataFrame
    agent.columns = ['uname','aname']
    agent['clean_aname'] = agent['aname'].str.strip().apply(lambda x: renamef(x,clean_aname)).apply(lambda x:re.sub(r'\_(.*)','',x)).replace(pat_chinese(),'')
    dd_l = agent['uname'].str.strip().str.split(r'\、',expand=True).stack().to_frame(name = 'uname').reset_index();dd_l
    dd_r = agent[agent.columns.difference(['uname'])].reset_index();dd_r
    agent = merge(dd_l,dd_r,how='outer',on=['level_0','level_1']).drop(['level_1','level_2'],axis=1)
    return agent

def merge_user(user,agent):
    return merge(user,agent,how='left',on=['level_0','uname'])

def merge_agent_adr(agent,adr):
    agent['clean_aname'].replace('',float("nan"),inplace=True)
    agent['clean_aname'] = agent['clean_aname'].fillna(agent['uname']).replace(pat_chinese(),'')
    adr['clean_aname'] = adr['clean_aname'].apply(lambda x: renamef(x,agent['clean_aname'].tolist()))
    tb = merge(agent,adr,how='outer',on=['level_0','clean_aname'])
    return tb

def reclean_data(tb):
    tg = tb.fillna('').groupby(['level_0','clean_aname','aname','address'])['uname'].apply(lambda x: '、'.join(x.astype(str))).reset_index()
    glist = tg['uname'].str.split(r'\、',expand=True).stack().values.tolist()
    rest = tb[tb['uname'].isin(glist) == False]
    data = concat([rest,tg],axis=0,sort=True)
    return data

def sort_data(data):
    data = data[['level_0','uname','aname','address']].sort_values(by=['level_0']).fillna('***').replace('','***');data
    data = merge(number,data,how='right',on=['level_0']).drop(['level_0'],axis=1)
    return data

agent_adr = df[['aname','address']]
agent = df['aname']
adr = df['address']

if agent_adr.apply(lambda x: check_format(x,check_data_flag), axis=0).all():
    print('找到代理人和地址...正在处理...')
    adr = make_adr(adr)
    agent = make_agent(agent,adr['clean_aname'].tolist())
    agent = merge_user(user,agent)
    agent = merge_agent_adr(agent,adr)
    data = reclean_data(agent)
    data = sort_data(data)
elif check_format(adr,check_data_flag):
    adr = make_adr(adr,user['uname'].tolist())
    adr['uname'] = adr['clean_aname']
    adr = merge_user(user,adr)
    adr = adr.assign(aname='')
    data = reclean_data(adr)
    data = sort_data(data)
elif check_format(agent,check_data_flag):
    agent = make_agent(agent)
    agent = merge_user(user,agent)
    agent = agent.assign(address='')
    data = reclean_data(agent)
    data = sort_data(data)
else:
    agent_adr.index.name = 'level_0'
    agent_adr.reset_index(inplace=True)
    data = merge(user,agent_adr,how='left',on=['level_0']);data
    data = sort_data(data)

print('>>> 完成生成新数据...end => data')
#%% check_data_flag and save data
if check_data_flag:
    if '***' in data[['aname','address']].values:
        print('>>> 提示:代理人和地址还有空缺...不影响继续生成')
        
if tmp_file and len(data):
    data_tmp = os.path.splitext(data_xlsx)[0]+"_temp.xlsx"
    data.columns = titles_en
    #data.to_excel(data_tmp,index=False)
    try:save_adjust_xlsx(data,data_tmp,titles_en)
    except PermissionError: input('>>> data.xlsx 在其他地方打开...请手动关闭并重新执行...任意键退出');sys.exit()
else: input('>>> 生成数据失败,请检查源data.xlsx文件...退出');sys.exit()

#%% generate postal sheets

def reg_text_group(reg,repl,text):
    z = re.search(reg,text)
    z = z.group(1) if bool(z) else ''
    text = re.sub(z,repl,text) if bool(z) else 'replace failed'
    return text

def re_writ_text(data):
    doc = Document(sheet_docx)
    doc.styles['Normal'].font.bold = True
    
    is_star = lambda x: bool(x == '***')
    clean_text = lambda x,y: (y+'#'+'暂缺') if is_star(x) else x
    agent_text = clean_text(data['aname'],'代理') if not is_star(data['aname']) else clean_text(data['uname'],'当事人')
    user_text = '代 '+ clean_text(data['uname'],'当事人') if not is_star(data['uname']) and not data['uname'] in agent_text else ''
    address_text = clean_text(data['address'],'地址')
    number_text = clean_text(data['number'],'案号')
    
    paragraph = doc.paragraphs[9]  # No.9 line is agent name
    text = re.sub('([\u4e00-\u9fff]).*',agent_text,paragraph.text)
    #print('===cutline=9==%s==%s'%(paragraph.text,text))
    paragraph.clear().add_run(text)
    
    paragraph = doc.paragraphs[11]  # No.11 line is user name
    text = re.sub('([\u4e00-\u9fff]).*',user_text,paragraph.text)
    #print('===cutline==11=%s==%s'%(paragraph.text,text))
    paragraph.clear().add_run(text)
    
    
    paragraph = doc.paragraphs[13]  # No.13 line is number and address
    text = reg_text_group(r'\s{2,}([\d\（\(].*?号)',number_text,paragraph.text)
    paragraph.clear().add_run(text)
    text = reg_text_group(r'\s{2,}([\u4e00-\u9fff].+)',address_text,paragraph.text)
    #print('===cutline=%s==%s==%s'%(i,paragraph.text,text))
    paragraph.clear().add_run(text)

#    for i,run in enumerate(paragraph.runs):
#        print('===cutline=%s=%s='%(i,run.text))

    sheet_file = number_text+'_'+agent_text+'_'+user_text+'_'+address_text+'.docx'
    if '暂缺' in address_text:
        print('%s 暂缺不生成 <= %s'%(address_text,sheet_file))
    else:
        doc.save(parse_subpath(postal_path,sheet_file))
        print('已生成邮单 =>',sheet_file)

    return doc

if to_postal:
    print('>>> 正在生成邮单...')
    if not os.path.exists(sheet_docx):
        input('>>> 没有找到邮单模板%s...任意键退出' % sheet_docx);sys.exit()
    data.apply(lambda x:re_writ_text(x),axis = 1)
    #re_writ_text(data.iloc[1])

#%% main
# if __name__ == "__main__":
input('>>> 全部完成,可以回顾记录...任意键退出')
    


