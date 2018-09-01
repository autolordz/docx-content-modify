# -*- coding: utf-8 -*-
print("""
Auto generate word:docx from excel:xslx file.

Created on Wed Aug 15 11:06:04 2018

Depends on: python-docx,pandas.

@author: Autoz.George
""")

#%%
import os
import re
import sys
import configparser 
import datetime
from pandas import DataFrame, read_excel, merge, concat, set_option, to_datetime
from docx import Document  # doc = Document('*.docx')
from glob import glob
from StyleFrame import StyleFrame

#os.chdir("F:/pdf-title-rename/tmp")
set_option('max_colwidth',500)
set_option('max_rows', 50)
set_option('max_columns',50)

#%% config and default values
cfgfile = 'config.txt'
data_xlsx = 'data.xlsx'
sheet_docx = 'sheet.docx'
postal_path = os.path.join('.','postal')
jdocs_path = os.path.join('.','jdocs')
last_pages = 10
date_range = '2018-06-01:2018-08-01'
rename_jdocs = True
fill_jdocs_adr = True
to_postal = True
check_format = True
tmp_file = True
cfg = configparser.RawConfigParser()

def write_config(cfg,cfgfile):
     with open(cfgfile, 'w') as configfile:
        cfg.write(configfile)

if not os.path.exists(cfgfile):
    cfg.add_section('config')
    cfg.set('config', 'data_xlsx', data_xlsx)
    cfg.set('config', 'sheet_docx', sheet_docx)
    cfg.set('config', 'last_pages', last_pages)
    cfg.set('config', 'date_range', date_range)
    cfg.set('config', 'rename_jdocs', rename_jdocs)
    cfg.set('config', 'fill_jdocs_adr', fill_jdocs_adr)
    cfg.set('config', 'to_postal', to_postal)
    cfg.set('config', 'check_format', check_format)
    write_config(cfg,cfgfile)
else:
    cfg.read(cfgfile)
    data_xlsx = cfg.get('config', 'data_xlsx')
    sheet_docx = cfg.get('config', 'sheet_docx')
    last_pages = cfg.getint('config', 'last_pages')
    date_range = cfg.get('config', 'date_range')
    rename_jdocs = cfg.getboolean('config', 'rename_jdocs')
    fill_jdocs_adr = cfg.getboolean('config', 'fill_jdocs_adr')
    to_postal = cfg.getboolean('config', 'to_postal')
    check_format = cfg.getboolean('config', 'check_format')

titles_cn = ['立案日期','案号','原一审案号','主审法官','当事人','诉讼代理人','地址']
titles_en = ['datetime','number','prenumber','judge','uname','aname','address']

def parse_subpath(path,file):
    if not os.path.exists(path):
        os.mkdir(path)
    return os.path.join(path,file)
#%% read data

def check_contain_chinese(check_str):
    return any((u'\u4e00' <= char <= u'\u9fff') for char in check_str)

def parse_datetime(date):
    try:date = datetime.datetime.strptime(date,'%Y-%m-%d')
    except ValueError:print('时间范围格式有误,默认选取全部日期');date = ''
    return date

def titles_trans(df):
    '''Change titles between Chinese and English'''
    tag1,tag2 =  (titles_en,titles_cn) if titles_en[0] in df.columns.tolist() else (titles_cn,titles_en)
    titles_rest = df.drop(tag1,axis=1).columns.tolist()
    df = df[tag1 + titles_rest]
    df.columns = tag2 + titles_rest
    print('Change titles to:', 'Chinese' if not check_contain_chinese(tag1[0]) else 'English')
    return df

if os.path.exists(data_xlsx):
    df_orgin = read_excel(data_xlsx)
    df_orgin = titles_trans(df_orgin)
    df_orgin['datetime'] = to_datetime(df_orgin['datetime'])
    if last_pages:
        df = df_orgin.iloc[-last_pages:-1];df
    elif ':' in date_range:
        df_orgin.sort_values(by=['datetime'],inplace=True)
        dats = date_range.split(':')
        dats[0] = parse_datetime(dats[0]);dats[1] = parse_datetime(dats[1]);
        date_start = dats[0] if dats[0] else df_orgin['datetime'][0]
        date_end = dats[1] if dats[1] else df_orgin['datetime'].iloc[-1]
        df = df_orgin[(df_orgin['datetime']>date_start)&(df_orgin['datetime']<date_end)]
    print('读取记录成功...',df.columns.values)
else: print('data.xlsx记录文件不存在...退出');sys.exit()
    
#%% rename judgment docs
def rename_doc_by_infos(file):
    '''rename only judgment doc files'''
    doc = Document(file)
    for i,paragraph in enumerate(doc.paragraphs[:10]):
        if re.search('民(申|终)\d+号',paragraph.text):
            os.rename(file,os.path.join(os.path.split(file)[0],'判决书_'+paragraph.text+'.docx'))
            
if rename_jdocs:
    docs = glob(parse_subpath(jdocs_path,'*.docx'))
    print('开始重命名判决书...')
    for file in docs:
        if not '判决书' in file:
            rename_doc_by_infos(file)
    print('重命名完毕...')

#%% merge address from judgment docs (Chinese Format)
    
def save_adjust_xlsx(df,file,titles):
    '''Save and adjust excel'''
    df = df.astype(str)
    excel_writer = StyleFrame.ExcelWriter(file)
    sf = StyleFrame(df)
    sf.to_excel(excel_writer=excel_writer,best_fit=titles)
    excel_writer.save()
    print('文件名:%s...列名:%s...校对数据保存成功...' %(file,titles))
    
def get_pre_address(doc):
    '''get pre address from judgment docs, return docs pre code and address'''
    adrs = []
    number = doc.paragraphs[2].text.strip() # number 
    for i,paragraph in enumerate(doc.paragraphs[3:10]): # range from 3:10 lines
        #print('===docline=%s=%s='%(i,paragraph.text))
        if re.search('诉讼|代理|律师',paragraph.text): continue # filter agent
        if '市' in paragraph.text: # find user with address
            try:
                name = re.search('\：([\u4e00-\u9fff]+)\，',paragraph.text).group(1)
                adr = ''.join([x for x in re.findall('([\u4e00-\u9fff]+)+',paragraph.text) if ('市' in x)])
                adrs.append(name+'/'+'地址：'+adr.replace('住',''))
            except AttributeError: pass
    address = '，'.join(adrs)
    print('=找到地址==%s==%s='%(number,address))
    return number,address

def fill_duplicate_adr(dfn):
    '''combine address between data and judgment docs'''
    dfn_adr = dfn[['address', 'new_adr']].fillna('***')
    for i,item in enumerate(dfn_adr.values):
        if str(item[1]) in str(item[0]):item[1] = ''
        #print('==%s=%s==%s'%(i,item[0],item[1]))
    dfn_adr.replace('***','',inplace=True)
    dfn['address'] = dfn_adr.apply(lambda x: '，'.join(x.fillna('').map(str)).strip(',|，'), axis=1) 
    dfn.drop(['new_adr'],axis=1,inplace=True);
    return dfn

if fill_jdocs_adr:
    print('开始填充判决书地址...')
    docs = glob(parse_subpath(jdocs_path,'判决书_*.docx'))
    if len(docs)>0:
        print('判决书:',docs)
        numlist=[]; adrlist = []
        for file in docs:
            number,address = get_pre_address(Document(file))
            numlist.append(number);adrlist.append(address);
        dfp = DataFrame({'prenumber':numlist,'new_adr':adrlist});dfp
        dfn = merge(df_orgin,dfp,how='left',on=['prenumber'])
        dfn = fill_duplicate_adr(dfn);dfn['address']
        try:
            dfn = titles_trans(dfn)
            save_adjust_xlsx(dfn,data_xlsx,dfn.columns.tolist())
            #dfn.to_excel(data_xlsx,index=False)
            print('填充地址到原文件完毕...请手动填充%s的代理人' % data_xlsx)
        except PermissionError:
            print('data.xlsx在其他地方打开...请手动关闭')
    else: print('没有找到判决书docx,先复制判决书到jdocs目录...退出');sys.exit()
else: print('fill_jdocs_adr = False控制不填充地址 ,继续下一步...')
#%% No.
def renamef(x,y):
    '''Clean agent name for agent to match address's agent name'''
    for name in y.tolist():
        if name in x:
            x = x.replace(x,name)
            break
    return x

print('生成新数据...')
number = df[titles_en[:4]]
number = number.reset_index()
number.columns.values[0] = 'level_0';number
# user
user = df['uname'];user
user = user.str.strip().str.split(r'\,|\，|\n|\t',expand=True).stack();user # divide user    
user = user.str.strip().str.split(r'\:',expand=True);user # divide character
user = user[1].str.strip().str.split(r'\、',expand=True).stack().to_frame(name = 'uname');user #DataFrame
user = user.reset_index().drop(['level_1','level_2'],axis=1);user
# address
adr = df['address'];adr
adr = adr.str.strip().str.split(r'\,|\，|\n|\t',expand=True).stack();adr #Series    
adr = adr.str.strip().str.split(r'\/',expand=True);adr #DataFrame
adr.columns = ['aname','address']
adr['address'] = adr['address'].apply(lambda x:re.sub(r'地址(\:|\：)','',x))# drop 地址:
adr['clean_aname'] = adr['aname'].str.strip().apply(lambda x: re.sub(r'\_(.*)','',x))
adr = adr.reset_index().drop(['level_1','aname'],axis=1);adr
# agent
agent = df['aname'];agent
agent = agent.str.strip().str.split(r'\,|\，|\n|\t',expand=True).stack();agent #Series    
agent = agent.str.strip().str.split(r'\/',expand=True);agent #DataFrame
agent.columns = ['uname','aname'];agent
agent['clean_aname'] = agent['aname'].apply(lambda x: renamef(x,adr['clean_aname'])).apply(lambda x:re.sub(r'\_(.*)','',x));agent
agent = agent.reset_index().drop(['level_1'],axis=1);agent
# combine
dd = merge(agent,adr,how='left',on=['level_0','clean_aname']).drop(['clean_aname'],axis=1)
cc = dd.rename(columns={"level_0": 'index'})
cc.set_index('index',append=True,inplace=True)
dd_l = cc['uname'].str.strip().str.split(r'\、',expand=True).stack().to_frame(name = 'uname').reset_index()
dd_l = dd_l.rename(columns={"level_0": 'index',"index": 'level_0'})
dd_r = dd[dd.columns.difference(['uname'])].reset_index()
dd = merge(dd_l,dd_r,how='outer',on=['index','level_0']).drop(['index','level_2'],axis=1)
tt = merge(user,dd,how='left',on=['level_0','uname']);tt
gg = tt.groupby(['level_0','aname','address'])['uname'].apply(lambda x: '、'.join(x.astype(str))).reset_index()
glist = gg['uname'].str.split(r'\、',expand=True).stack().values.tolist()
rest = tt[tt['uname'].isin(glist) == False]
data = concat([rest,gg],axis=0,sort=True)
data = data[['level_0','uname','aname','address']].sort_values(by=['level_0']).fillna('***').replace('','***')
data = merge(number,data,how='outer',on=['level_0']) # merge number and data
data.drop(['level_0'],axis=1,inplace=True);data
if check_format:
    if '***' in data[['aname','address']].values:
        print('提示:代理人和地址还有空缺...不影响继续生成')
        
if tmp_file and len(data):
    data_tmp = os.path.splitext(data_xlsx)[0]+"_temp.xlsx"
    data.columns = titles_en
    #data.to_excel(data_tmp,index=False)
    save_adjust_xlsx(data,data_tmp,titles_en)
else: print('生成数据失败,请检查源data.xlsx文件...');sys.exit()

#%% replace postal_sheet.docx

def reg_text_group1(reg,repl,text):
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
    text = reg_text_group1(r'\s{2,}([\d\（\(].*?号)',number_text,paragraph.text)
    paragraph.clear().add_run(text)
    text = reg_text_group1(r'\s{2,}([\u4e00-\u9fff].+)',address_text,paragraph.text)
    #print('===cutline=%s==%s==%s'%(i,paragraph.text,text))
    paragraph.clear().add_run(text)

#    for i,run in enumerate(paragraph.runs):
#        print('===cutline=%s=%s='%(i,run.text))

    sheet_file = number_text+'_'+agent_text+'_'+user_text+'_'+address_text+'.docx'
    doc.save(parse_subpath(postal_path,sheet_file))
    print('已生成邮单...',sheet_file)
    return doc

if to_postal:
    print('生成邮单...')
    if not os.path.exists(sheet_docx):
        print('没有找到邮单模板%s...退出' % sheet_docx);sys.exit()
    data.apply(lambda x:re_writ_text(x),axis = 1)
    #re_writ_text(data.iloc[1])

# if __name__ == "__main__":
input('全部完成,可以回顾记录,或者按任意键退出...')
    


