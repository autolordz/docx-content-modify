
# docx-content-modify

## Newest

【2018-10-26】

> * 修改复制判决书内人员逻辑，判断与OA人员差异，聪明复制
> * 优化配置文件，尤其日期范围和条数
> * 添加判决书和邮单日志记录FLAG

## Guide

法院人员批量生成邮单程序，减轻书记员负担

> * 灵感来自于法院人员发送案件生成寄件邮单麻烦，完成功能如下
> * 技术:python-docx,pandas,StyleFrame,configparser
> * 打包程序:pyinstaller(单文件没压缩)

- [x] 重命名判决书
	- 判决书来自于[中国裁判文书网](http://wenshu.court.gov.cn/)
	- 重命名后格式 ***判决书_XXX.docx***

- [x] 批量填充判决书地址到数据模板
	- 先从 ***法院人员OA系统***(法院工作的都有)下载信息表
	- 添加OA数据[data_oa.xlsx](./demo_docs/data_oa.xlsx)到数据模板
 	- 选择自动填充判决书[判决书_XXX.docx](./demo_docs/jdocs)的地址到数据模板[data_main.xlsx](./demo_docs/data_main.xlsx),其他缺失的律师(代理人)及地址需要***手动***填充

- [x] 最后批量生成寄送邮单
	- [邮单模板](./demo_docs/sheet.docx)
	- 生成临时文件 ***data_temp.xlsx*** 用于校对,是邮单信息来源 

法院OA系统表格【data_oa.xlsx】必须包含如下字段:

**【承办人】会转换为【主审法官】**

| 【立案日期】 | 【案号】 | 【原一审案号】 | 【承办人】 | 【当事人】 | 【其他】... |


【data_main.xlsx】包括字段：

| 【立案日期】 | 【案号】 | 【原一审案号】 | 【主审法官】 | 【当事人】 | 【诉讼代理人】 | 【地址】 | 【其他】... |


第一次运行会生成配置文件:

```python
[config]
data_xlsx = data_main.xlsx # 数据模板地址
data_oa_xlsx = data_oa.xlsx # OA数据地址
sheet_docx = sheet.docx # 邮单模板地址
flag_rename_jdocs = True # 是否重命名判决书
flag_fill_jdocs_adr = True # 是否填充判决书地址
flag_fill_phone = False # 是否填充伪手机
flag_append_oa = True # 是否导入OA数据
flag_to_postal = True # 是否打印邮单
flag_check_jdocs = False # 是否检查用户格式,输出提示信息
flag_check_postal = False # 是否检查邮单格式,输出提示信息
date_range =  #2018-09-01:2018-12-01 # 打印数据日期范围,比行数优先,去掉注释后读取,井号注释掉
last_lines_oa = 50 # 导入OA数据的最后几行,当flag_append_oa开启才有效
last_lines_data = 50 # 打印数据的最后几行
```

除了部分当事人地址自动填充外,填充律师规则如下:

***姓名要保持一致包括曾用名,姓名/姓名_电话,逗号表示分隔,顿号表示一起,'/地址：'不能缺***

| 【当事人】 | 【诉讼代理人】 | 【地址】 |
| --- | --- | --- |
| 申请人:姓名A,被申请人:姓名B | 姓名A/律师姓名C_电话,姓名B_电话 | 姓名B/地址：XXX市XXX,姓名C/地址：XXX市XXX |
| 申请人:张三(曾用名张五)、李四、王五 | 张三(曾用名张五)/律师张二三_123123_李三四_123123 | 张二三/地址：XXX市XXX |
| 申请人:赵六(曾用名:赵五)、孙七、周八 | 赵六(曾用名:赵五),孙七、周八/代理人吴九_123123,郑十/委托人张三_123123| 赵六(曾用名:赵五)/地址：XXX市XXX,吴九/地址：XXX市XXX,张三/地址：XXX市XXX |

生成的邮单是没有代理人的单独一份,有代理人的几个当事人合一份,法院书记员的都懂

看不懂说明的可以直接下载最新的exe版本(win7/win10)

THE END
Enjoy

## Licence

- 版权见代码,MIT Licence
- 打包好exe即源码,没有后门放心使用