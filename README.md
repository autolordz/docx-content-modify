## 邮单自动批量生成器

> * 法院法务自动化批量生成邮寄单据-Legal agency postal notes automatically generate app
> * 给予法务邮递人员从法务OA数据表(excel)和公开的判决书(docx)提取当事人地址内容，批量直接生成邮单。 减轻相关员负担，尤其系列案，人员多地址多，手工输入地址重复性劳动太多，信息容易错漏

> * 编译环境:  
> * conda version : 4.6.14
> * conda-build version : 3.17.8
> * python version : 3.7.3.final.0
> * Win10+Spyder3.3.4(自上向下运行,或者自己添加main来py运行)
> * 使用编程技术:python-docx,pandas,StyleFrame,configparser  
> * 打包程序:pyinstaller

[![](https://img.shields.io/github/release/autolordz/docx-content-modify.svg?style=popout&logo=github&colorB=ff69b4)](https://github.com/autolordz/docx-content-modify/releases)
[![](https://img.shields.io/badge/github-source-orange.svg?style=popout&logo=github)](https://github.com/autolordz/docx-content-modify)
[![](https://img.shields.io/github/license/autolordz/docx-content-modify.svg?style=popout&logo=github)](https://github.com/autolordz/docx-content-modify/blob/master/LICENSE)

## 目录

<!-- MarkdownTOC autoanchor="true" autolink="true" uri_encoding="false" -->

- [更新](#更新)
- [内容](#内容)
- [规则](#规则)
- [具体使用方法](#具体使用方法)
- [Licence](#licence)

<!-- /MarkdownTOC -->

<a id="更新"></a>
## 更新

【2019-6-19】  
> * 添加合并系列案功能，节省打印资源

【2019-6-12】

> * 更新判决书过滤词汇


<a id="内容"></a>
## 内容

- [x] 按格式重命名判决书
	- [x]提取判决书人员和地址信息
	- [x]自动重命名为 **判决书_AAA号_原_BBB号.docx**

- [x] 拷贝OA表记录到Data表
	- [x]按数量提取，按日期提取，按指定案号提取
	- [x]整理Data表格式，对表中数据的变形，清洗，符合打印邮单的字段格式
	- [x]填充判决书信息到Data表

- [x] 按照Data表输出寄送邮单
	- [x]填充好所有信息，再次运行就能输出Data表指定邮单

<a id="规则"></a>
## 规则

1. 当事人收信规则，没代理律师的每个当事人一份，有委托律师的只要寄给律师一份，多个律师寄给第一个律师，同一律所也是一份 

1.1 判决书过滤词汇，判决书如果每行包含如下就不提取信息  

词汇1：法定代表|诉讼|代理人|判决|律师|请求|证据|辩称|辩论|不服  
词汇2：上市|省略|区别|借款|保证|签订  

2. 字段解析:  

OA表【data_oa.xlsx】必须字段:  
| 【立案日期】 | 【适用程序】 | 【案号】 | 【原一审案号】 | 【承办人】 | 【当事人】 | ... |  
Data表【data_main.xlsx】必须字段(包括程序生成):  
| 【立案日期】 | 【适用程序】 | 【案号】 | 【原一审案号】 | 【判决书源号】 | 【主审法官】 | 【当事人】 | 【诉讼代理人】 | 【地址】 | ... |  
注意： 数据表处理后【承办人】会更换为【主审法官】  

3. 【诉讼代理人】规则:  

**注意：姓名和曾用名如例子所示，'/'前面是当事人，后面是律师，'_'连接电话，逗号'，'表示分隔，顿号表示一起，'/地址：'不能缺**     

Data表部分字段演示：  

| 【当事人】 | 【诉讼代理人】 | 【地址】 |
| --- | --- | --- |
| 申请人:姓名AAA，被申请人:姓名BBB| 姓名AAA/律师姓名CCC_电话，姓名BBB_电话 | 姓名BBB/地址：XXX市XXX，姓名CCC/地址：XXXX市XXX |
| 申请人:张三(曾用名张五)、李四、王五 | 张三(曾用名张五)/律师张二三_123123_李三四_123123 | 张二三/地址：XXXX市XXX |
| 申请人:赵六(曾用名:赵五)、孙七、周八 | 赵六(曾用名:赵五)，孙七、周八/代理人吴九_123123，郑十/委托人张三_123123| 赵六(曾用名:赵五)/地址：XXX市XXX，吴九/地址：XXX市XXX，张三/地址：XXX市XXX |

4. 【适用程序】规则(系列案用):  

此处在OA表的【适用程序】填写,人为判断几个案是同一系列案的请在该字段中标注len(str)>3的唯一记号,系列案会自动合并

len(str)>3 = 记号多于三个字符 

例如：  

| 【适用程序】 | 【案号】 |
| --- | --- |
| AAA       | 2773 |
| 2774-2776 | 2774 |
| 2774-2776 | 2775 |
| 2774-2776 | 2776 |


5. config.txt:  
```python
[config]
data_xlsx = data_main.xlsx    # 数据模板地址
data_oa_xlsx = data_oa.xlsx    # OA数据地址
sheet_docx = sheet.docx    # 邮单模板地址
flag_fill_jdocs_infos = 1    # 是否填充判决书地址
flag_append_oa = 1    # 是否导入OA数据
flag_to_postal = 1    # 是否打印邮单
flag_check_jdocs = 0    # 是否检查用户格式,输出提示信息
flag_check_postal = 0    # 是否检查邮单格式,输出提示信息
flag_output_log = 1    # 是否保存打印
data_case_codes =    # 指定打印案号,可接多个,示例:AAA,BBB,优先级1
data_date_range =   # 指定打印数据日期范围示例:2018-09-01:2018-12-01,优先级2
data_last_lines = 10    # 指定打印最后行数,优先级3
```

<a id="具体使用方法"></a>
## 具体使用方法

简称：  
- [A表: data_oa.xlsx,OA表自己下载,这个只是参考](./demo_docs/data_oa.xlsx)
- [B表: data_main.xlsx,会自动生成,也要修改](./demo_docs/data_main.xlsx)
- [C目录: jdocs/,判决书目录,要放下载的判决书](./demo_docs/jdocs/)
- [D文档: sheet.docx,邮单模板,按照背景生成邮单](./demo_docs/sheet.docx)
- [E目录: postal/,邮单目录](./demo_docs/postal/)  

1. 根据 **A表** 格式,整理自己的OA表(没数据是没用的),先在OA表中修改【适用程序】(系列案),修改conf.txt文件,参考[规则](#规则),如文件丢失再次运行会生成  

2. 手动下载在线判决书[中国裁判文书网](http://wenshu.court.gov.cn/),判决书docx文件放在**C目录**   

3. 第一次运行(不带【诉讼代理人】)  

3.2. 运行会自动重命名判决书 **C目录** ,提取判决书内容 **address_tmp.xlsx**     
3.3. 自动从 **A表** 添加数据到 **B表**  
3.4. 自动通过 **D文档** 批量输出寄送邮单到 **E目录**  
3.5. 自动生成临时文件 ***data_tmp.xlsx*** 用于校对,是邮单信息来源   
3.6. 运行记录在log.txt  

4. 手动填充 **完整** 的律师(代理人)及当事人信息到 **B表**，具体是填写【诉讼代理人】信息(电话地址)，参考[规则](#规则)  

5. 第二次运行(带【诉讼代理人】)  
会重复 3.4.  3.5. 3.6.  

6. 看不懂以上说明的可以直接下载最新的exe版本[win7/win10(32/64))](https://github.com/autolordz/docx-content-modify/releases/download/1.0.1/exe-win7win10-8962f68c.zip)

<a id="licence"></a>
## Licence

[See Licence](https://github.com/autolordz/docx-content-modify/blob/master/LICENSE)

THE END
Enjoy
