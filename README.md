
# docx-content-modify

## Newest

【2018-10-09】

> * 高亮导入OA数据
> * 优化导入OA数据逻辑
> * 优化设置FLAG逻辑
> * 优化保存文件逻辑

## Guide

法院书记员批量生成邮单脚本程序，减轻书记员负担

> * 灵感来自于法院书记员发送案件复制邮单麻烦，完成功能如下
> * 技术:python-docx,pandas
> * 打包程序:pyinstaller(单文件没压缩)

- [x] 重命名判决书
	- 判决书来自于[中国裁判文书网](http://wenshu.court.gov.cn/)
	- 重命名后格式 ***判决书_XXX.docx***

- [x] 批量填充判决书地址到数据模板
	- 先从 ***法院书记员OA系统*** 下载信息表
	- 编辑数据模板[data.xlsx](./demo_docs/data.xlsx)
 	- 选择自动填充判决书的地址到数据模板[data.xlsx](./demo_docs/data.xlsx),其他缺失的律师和地址需要***自行***填充

- [x] 最后批量生成寄送邮单
	- [邮单模板](./demo_docs/sheet.docx)
	- 生成临时文件 ***data_temp.xlsx*** 用于校对,是邮单信息来源 

法院OA系统表格[data_oa.xlsx]必须包含如下字段:
**其实【承办人】就是【主审法官】,法院工作人员都懂**

| 【立案日期】 | 【案号】 | 【原一审案号】 | 【承办人】 | 【当事人】 |


[data.xlsx]包括字段：

| 【立案日期】 | 【案号】 | 【原一审案号】 | 【主审法官】 | 【当事人】 | 【诉讼代理人】 | 【地址】 | 【其他】... |


第一次运行会生成配置文件:

```python
[config]
# 数据模板地址
data_xlsx = data.xlsx
# oa数据地址
data_oa_xlsx = data_oa.xlsx
# 邮单模板地址
sheet_docx = sheet.docx
# 是否重命名判决书
flag_rename_jdocs = True
# 是否填充判决书地址
flag_fill_jdocs_adr = True
# 是否填充伪手机
flag_fill_phone = True
# 是否导入oa数据
flag_append_oa = True
# 导入oa数据的最后几行
oa_last_lines = 30
# 是否打印邮单
flag_to_postal = True
# 打印数据模板的最后几行
data_last_lines = 100
# 打印数据模板的日期范围
date_range = 2018-06-01:2018-08-01
# 检查数据模板的内容格式
flag_check_data = False
```
除了部分当事人地址自动填充外,填充律师规则如下:

***姓名要保持一致包括曾用名,姓名/姓名_电话,逗号表示分隔,顿号表示一起,'/地址：'不能缺***

| 【当事人】 | 【诉讼代理人】 | 【地址】 |
| 申请人:姓名A,被申请人:姓名B | 姓名A/律师姓名C_电话,姓名B_电话 | 姓名B/地址：XXX市XXX,姓名C/地址：XXX市XXX |
| :------| ------: | :------: |
| 申请人:张三(曾用名张五)、李四、王五 | 张三(曾用名张五)/律师张二三_123123_李三四_123123 | 张二三/地址：XXX市XXX |
| 申请人:赵六(旧名:赵五)、孙七、周八 | 赵六(旧名:赵五),孙七、周八/代理人吴九_123123,郑十/委托人张三_123123| 赵六(旧名:赵五)/地址：XXX市XXX,吴九/地址：XXX市XXX,张三/地址：XXX市XXX |

生成的邮单是没有代理人的单独一份,有代理人的几个当事人合一份,法院书记员的都懂

看不懂说明的可以直接下载最新的exe版本(win7/win10)

THE END
Enjoy

## Licence

- 版权见代码
- MIT Licence
