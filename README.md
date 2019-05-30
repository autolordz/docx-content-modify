## 邮单自动批量生成器

> * 法院法务自动化批量生成邮寄单据-Legal agency postal notes automatically generate app
> * 给予法务邮递人员从法务OA数据表(excel)结合网上公开的判决书(docx)提取当事人地址内容，批量直接生成邮单。 减轻相关员负担，尤其系列案，人员多地址多，手工输入地址重复性劳动太多，信息容易错漏
> * 使用编程技术:python-docx,pandas,StyleFrame,configparser  
> * 打包程序:pyinstaller

> 编译环境:
> conda version : 4.6.14
> conda-build version : 3.17.8
> python version : 3.7.3.final.0
> Win10+Spyder3.3.4(自上向下运行,或者自己添加main来py运行)

[![](https://img.shields.io/github/release/autolordz/docx-content-modify.svg?style=popout&logo=github&colorB=ff69b4)](https://github.com/autolordz/docx-content-modify/releases)
[![](https://img.shields.io/badge/github-source-orange.svg?style=popout&logo=github)](https://github.com/autolordz/docx-content-modify)
[![](https://img.shields.io/github/license/autolordz/docx-content-modify.svg?style=popout&logo=github)](https://github.com/autolordz/docx-content-modify/blob/master/LICENSE)

## 目录

<!-- MarkdownTOC autoanchor="true" autolink="true" uri_encoding="false" -->

- [更新](#更新)
- [内容](#内容)
- [规则](#规则)
- [使用方法](#使用方法)
- [Licence](#licence)

<!-- /MarkdownTOC -->

<a id="更新"></a>
## 更新

【2019-5-30】

> * 更改配置文件选项,添加打印指定条目
> * 优化代码,代码添加中文注释

<a id="内容"></a>
## 内容

- [A表: data_oa.xlsx](./demo_docs/data_oa.xlsx)
- [B表: data_main.xlsx](./demo_docs/data_main.xlsx)
- [C目录: /jdocs/](./demo_docs/jdocs/)
- [D文档: 判决书_XXX.docx](./demo_docs/jdocs/jdocs.docx)
- [E文档: 邮单模板](./demo_docs/sheet.docx)

- [x] 重命名判决书
	- 手动下载公开的判决书[中国裁判文书网](http://wenshu.court.gov.cn/)到**C目录**
	- 自动重命名格式 **判决书_XXX.docx**

- [x] 批量填充判决书地址到数据模板
	- 手动从 **法务OA系统** (非公开)下载 **A表**, 格式请参考Demo来调整列数据
	- 自动从 **A表** 添加数据到 **B表** 
 	- 自动填充的 **D文档** 的 **非完整** 的地址等信息到 **B表**
 	- 手动填充 **完整** 的律师(代理人)及当事人信息到 **B表**

- [x] 批量生成寄送邮单
	- 自动通过 **E文档** 批量生成寄送邮单
	- 自动生成临时文件 ***data_temp.xlsx*** 用于校对,是邮单信息来源 

<a id="规则"></a>
## 规则

1. 当事人收信规则，没代理律师的每个当事人一份，有委托律师的只要寄给律师一份，多个律师寄给第一个律师，同一律所也是一份  

2. 字段解析:  

法务OA源文件的【data_oa.xlsx】字段:  
| 【立案日期】 | 【案号】 | 【原一审案号】 | 【承办人】 | 【当事人】 | 【其他】... |  
数据表的【data_main.xlsx】字段:  
| 【立案日期】 | 【案号】 | 【原一审案号】 | 【主审法官】 | 【当事人】 | 【诉讼代理人】 | 【地址】 | 【其他】... |  
注意： 数据表处理后【承办人】会更换为【主审法官】  

<a id="使用方法"></a>
## 使用方法

1. 判决书docx文件放在 /jdocs  
2. 首次运行会生成配置文件conf.txt:
```python
[config]
data_xlsx = data_main.xlsx    # 数据模板地址
data_oa_xlsx = data_oa.xlsx    # OA数据地址
sheet_docx = sheet.docx    # 邮单模板地址
flag_rename_jdocs = 1    # 是否重命名判决书
flag_fill_jdocs_infos = 1    # 是否填充判决书地址
flag_append_oa = 1    # 是否导入OA数据
flag_to_postal = 1    # 是否打印邮单
flag_check_jdocs = 0    # 是否检查用户格式,输出提示信息
flag_check_postal = 0    # 是否检查邮单格式,输出提示信息
flag_output_log = 1    # 是否保存打印
data_case_codes =  # 指定打印案号,可接多个,示例:AAA,BBB,优先级1
data_date_range =   # 指定打印数据日期范围示例:2018-09-01:2018-12-01,优先级2
data_last_lines = 10    # 指定打印最后行数,优先级3
```
3. 首次运行会自动填写诉讼人员电话地址  
4. 修改conf.txt  
5. 除了部分当事人地址自动填充外,填充律师规则如下:  

**注意：姓名和曾用名如例子所示，'/'前面是当事人，后面是律师，'_'连接电话，逗号'，'表示分隔，顿号表示一起，'/地址：'不能缺**     

例子：  

| 【当事人】 | 【诉讼代理人】 | 【地址】 |
| --- | --- | --- |
| 申请人:姓名AAA，被申请人:姓名BBB| 姓名AAA/律师姓名CCC_电话，姓名BBB_电话 | 姓名BBB/地址：XXX市XXX，姓名CCC/地址：XXXX市XXX |
| 申请人:张三(曾用名张五)、李四、王五 | 张三(曾用名张五)/律师张二三_123123_李三四_123123 | 张二三/地址：XXXX市XXX |
| 申请人:赵六(曾用名:赵五)、孙七、周八 | 赵六(曾用名:赵五)，孙七、周八/代理人吴九_123123，郑十/委托人张三_123123| 赵六(曾用名:赵五)/地址：XXX市XXX，吴九/地址：XXX市XXX，张三/地址：XXX市XXX |

6. 再次运行exe
7. 生成的邮单在 **postal/** ,当事人没有律师的单独一份,有律师的几个当事人合一份
8. 看不懂以上说明的可以直接下载最新的exe版本[win7/win10](https://github.com/autolordz/docx-content-modify/releases/download/1.0.1/exe-win7win10-8962f68c.zip)

<a id="licence"></a>
## Licence

[See Licence](https://github.com/autolordz/docx-content-modify/blob/master/LICENSE)

THE END
Enjoy