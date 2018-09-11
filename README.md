
# docx-content-modify

法院书记员批量生成邮单脚本程序，减轻书记员负担

> * 灵感来自于法院书记员发送案件复制邮单麻烦，完成功能如下
> * 技术:python,docx,pandas
> * 打包程序:pyinstaller(单文件没压缩)

- [x] 重命名判决书
	- 判决书来自于[中国裁判文书网](http://wenshu.court.gov.cn/)
	- 命名格式***判决书_XXX.docx***

- [x] 批量填充判决书地址到数据模板
	- 先从***法院书记员OA系统***下载信息表,编辑data.xlsx包括字段 ['立案日期','案号','原一审案号','主审法官','当事人','诉讼代理人','地址','备注',...]

 	- 再自动填充判决书中有地址的角色到[数据模板](./demo_docs/data.xlsx),其他缺失的律师和地址需要***自行***填充

- [x] 最后批量生成寄送邮单
	- [邮单模板](./demo_docs/sheet.docx)
	- 生成临时文件***data_temp.xlsx***用于校对,是邮单信息来源 

第一次运行会生成配置文件:

```python
[config]
data_xlsx = data.xlsx #数据模板
sheet_docx = sheet.docx #邮单模板
last_pages = 40  #生成邮单记录条数,从最后数起
date_range = 2018-06-01:2018-08-01 #生成日期范围,last_pages不填或0时才有用
rename_jdocs = True #是否重命名
fill_jdocs_adr = True #是否填充地址
to_postal = True #是否生成邮单
check_format = True #是否校对格式
```
data.xlsx 填充格式,重点是['诉讼代理人','地址']

除了部分当事人地址自动填充外,填充律师规则如下:

| 【当事人】 | 【诉讼代理人】 | 【地址】 |
| :------| ------: | :------: |
| 申请人:张三、李四、王五 | 张三/律师张三三_xxx_李四四_xxx | 张三三/地址：XXX市XXX |
| 申请人:赵六、孙七、周八 | 赵六、孙七、周八/代理人吴九_xxx,郑十/代理人张三_xxx| 吴九/地址：XXX市XXX,张三/地址：XXX市XXX |

备注: 律师是用一律所用'_'连接, 不同律所用','连接,大小写符号都可以,其余斜杠顿号逗号都一定严格按照标准

生成的邮单是没有代理人的单独一份,有代理人的几个当事人一份,法院书记员的都懂

看不懂说明的可以直接下载exe版本(win7/win10)

THE END
Enjoy

# 版权见licence

- MIT licence
