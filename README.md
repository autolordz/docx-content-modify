<a id="docx-content-modify"></a>
## docx-content-modify

> * 中国法院人员批量邮寄脚本(公开法律文书)court staffs postal receipt generate
> * 给予法院邮政人员从OA数据表(excel)和公开判决书(docx)提取当事人地址内容，批量直接生成邮单
> * 减轻相关员负担，尤其系列案，人员多地址多，手打地址重复性劳动太多，信息容易错漏

> 技术:python-docx,pandas,StyleFrame,configparser
> 打包程序:pyinstaller

[![](https://img.shields.io/github/release/autolordz/docx-content-modify.svg?style=popout&logo=github&colorB=ff69b4)](https://github.com/autolordz/docx-content-modify/releases)
[![](https://img.shields.io/badge/github-source-orange.svg?style=popout&logo=github)](https://github.com/autolordz/docx-content-modify)
[![](https://img.shields.io/github/license/autolordz/docx-content-modify.svg?style=popout&logo=github)](https://github.com/autolordz/docx-content-modify/blob/master/LICENSE)

## TOC

<!-- MarkdownTOC autoanchor="true" autolink="true" uri_encoding="false" -->

- [Updated](#updated)
- [Features](#features)
- [Rules](#rules)
- [Usage](#usage)
- [Licence](#licence)

<!-- /MarkdownTOC -->

<a id="updated"></a>
## Updated

【2018-11-13】

> * 优化配置文件日期范围和打印记录
> * 优化拷贝判决书上地址兼容性

<a id="features"></a>
## Features

- [x] 重命名判决书
	- 手动下载公开的判决书[中国裁判文书网](http://wenshu.court.gov.cn/)
	- 自动重命名格式 **判决书_XXX.docx**

- [x] 批量填充判决书地址到数据模板
	- 手动从 **法院人员OA系统**(非公开)下载信息表
	- 自动添加OA数据[data_oa.xlsx](./demo_docs/data_oa.xlsx)到数据模板
 	- 自动填充判决书[判决书_XXX.docx](./demo_docs/jdocs)的**非精确**的地址到数据模板[data_main.xlsx](./demo_docs/data_main.xlsx)
 	- 手动填充**精确**的律师(代理人)及当事人信息

- [x] 批量生成寄送邮单
	- 自动通过[邮单模板](./demo_docs/sheet.docx),批量生成寄送邮单
	- 自动生成临时文件 ***data_temp.xlsx*** 用于校对,是邮单信息来源 

<a id="rules"></a>
## Rules

1. 当事人收信规则，没代理律师的每个当事人一份，有委托律师的只要寄给律师一份，多个律师寄给第一个律师，同一律所也是一份
2. 法院OA系统表格【data_oa.xlsx】必须包含如下字段:  

【OA.xlsx】字段:
| 【立案日期】 | 【案号】 | 【原一审案号】 | 【承办人】 | 【当事人】 | 【其他】... |
注意：**【承办人】转换为【主审法官】**

【data_main.xlsx】字段:  
| 【立案日期】 | 【案号】 | 【原一审案号】 | 【主审法官】 | 【当事人】 | 【诉讼代理人】 | 【地址】 | 【其他】... |


<a id="usage"></a>
## Usage

1. 判决书docx文件放在 /jdocs  
2. 首次运行会生成配置文件conf.txt:
```python
[config]
data_xlsx = data_main.xlsx    # 数据模板地址
data_oa_xlsx = data_oa.xlsx    # OA数据地址
sheet_docx = sheet.docx    # 邮单模板地址
flag_rename_jdocs = 1    # 是否重命名判决书
flag_fill_jdocs_adr = 1    # 是否填充判决书地址
flag_fill_phone = 0    # 是否填充伪手机
flag_append_oa = 1    # 是否导入OA数据
flag_to_postal = 1    # 是否打印邮单
flag_check_jdocs = 1    # 是否检查用户格式,输出提示信息
flag_check_postal = 0    # 是否检查邮单格式,输出提示信息
date_range_oa_data = # 2018-01-01:2018-12-01    # 导入OA和打印数据日期范围,比行数优先,去掉注释后读取,井号注释掉
last_lines_oa_data = 200    # 导入OA和打印数据的最后几行
```
3. 首次运行会自动填写诉讼人员电话地址  
4. 修改conf.txt  
5. 除了部分当事人地址自动填充外,填充律师规则如下:  

**注意：姓名要保持一致包括曾用名,姓名/姓名_电话,逗号表示分隔,顿号表示一起,'/地址：'不能缺**
| 【当事人】 | 【诉讼代理人】 | 【地址】 |
| --- | --- | --- |
| 申请人:姓名A,被申请人:姓名B | 姓名A/律师姓名C_电话,姓名B_电话 | 姓名B**/地址：**XXX市XXX,姓名C/地址：XXX市XXX |
| 申请人:张三(曾用名张五)、李四、王五 | 张三(曾用名张五)/律师张二三_123123_李三四_123123 | 张二三**/地址：**XXX市XXX |
| 申请人:赵六(曾用名:赵五)、孙七、周八 | 赵六(曾用名:赵五),孙七、周八/代理人吴九_123123,郑十/委托人张三_123123| 赵六(曾用名:赵五)/地址：XXX市XXX,吴九/地址：XXX市XXX,张三/地址：XXX市XXX |

6. 再次运行exe
7. 生成的邮单在**postal/**,当事人没有律师的单独一份,有律师的几个当事人合一份
8. 看不懂说明的可以直接下载最新的exe版本[win7/win10](https://github.com/autolordz/docx-content-modify/releases/download/1.0.1/exe-win7win10-8962f68c.zip)

<a id="licence"></a>
## Licence

[See Licence](#docx-content-modify)

THE END
Enjoy