# *Modify docx. Without tears.*
快速、准确、批量替换文件夹内全部docx文档特定字段，并高亮显示存在修改的段落。
## 使用场景
需要根据项目属性更改多个制式word文档（.docx格式）中的多个相同字段，word上的常规流程为打开每一个word文档然后“查找并替换”每一个字段。

这个过程费时且重复性高，且偶尔会遗漏相关字段，因此我在`python-docx`库的基础上，写了这样一个python脚本，思路是将制式文档分为静态内容和动态内容，通过程序定位并替换动态内容，高亮显示存在修改的段落以便校对，同时实现批量修改特定文件夹下的全部word文档。
## 环境依赖
1. Python 3.0 及以上版本
2. python-docx库，在电脑终端输入并运行`pip install python-docx`即可安装。
## 文件结构及说明
```
modifydocx
├─ README.md
├─ modify.py
├─ text
│  ├─ text.txt
├─ template
│  ├─ 合伙协议.docx
└─ output
   └─ ABCD（有限合伙）_合伙协议[0313-104933].docx
```
- `text` 文件夹下存放了`text.txt`，可以在这个txt文件中创建和修改需要查找替换的字段和被替换的值，格式为`字段：值`（注意"："为中文全角），例如`基金名称：欢乐堡（有限合伙）`
- `template`文件夹下可存放多个docx的模版文件。将text中定义且需要被替换的字段按照`@字段@`的格式插入至相应位置，例如`blabla@基金名称@blabla`，
- `modify.py` 为程序运行主文件，在终端进入程序目录时，输入`python3 modify.py` 运行。
  
## 图例
![Modify docx, without tears.](https://s3.ax1x.com/2021/03/16/6s0K1I.jpg)