本来只是想用PDMS直接输出分类汇总好的Mto
然后想要直接输出Excel多好，看了一下PDMS自带的excel输出方式，感觉好复杂，最终放弃。

最终在范大师的指导和帮助下，调用Python完美解决了这个问题。

步骤：
1、先用PML输出文本文件，格式如下：

序号@物料编码@物料描述@数量@单位
3@AAFABRD0JJAI@法兰（SO） B16.5 2” RF CL300, A105@4@EA
4@AAFABRD0LLAI@法兰（SO） B16.5 3” RF CL300, A105@4@EA
29@AAGBJRB0RRBM@缠绕垫（带内环和对中环） B16.20 6” RF CL150, CS/304/Graphite@13@EA
30@AAGBJRB0TTBM@缠绕垫（带内环中环） B16.20 8” RF CL150, CS/304/Graphite@20@EA
44@AAPCBR0KEEAR@无缝钢管 B36.10 3/4” PE  Sch80, A106-B@6.9@M
45@AAPCBR0KFFAR@无缝钢管 B36.10 1” PE  Sch80, A106-B@1.7@M

注：分类汇总使用PML语言自带的功能实现的：sort 和 subtotal   其实也在python中完成
    var !Index sort !tMatCode lastingroup !Group
    var !TotalQty subtotal !tQty !Index !Group



2、通过PDMS的 "SYSCOM" 调用Python通过.txt生成.xlsx文件

 问题1：python；公司其他人的电脑上并没有安装python，所以必须保证服务器上有python程序，而服务器上不能随便安装python。

 解决：只能用自行编译的Python程序。这里范大师想到了Goagent的实现，可以用它编译好的python程序，这个不错。

 问题2：库；即便有了以上的python程序，但是服务器上并没有相关的库文件（openpyxl）。

 解决：将以上的python拷贝到服务器之后，查询它的sys.path，发现其值就是python.exe的路径。所以可以将库文件直接放到该文件夹中。
于是将openpyxl文件夹拷贝到该文件夹中，import的时候出现报错信息，貌似是因为这个库还有其他的很多依赖库，这个方法不行。
然后想到用virtualenv来实现，但是服务器还是没有python，没法在服务器上直接部署该环境，这也不行。
最后范大师采用了先在本地电脑部署安装virtualenv，只给他安装Openpyxl库，然后再将所有库文件copy到服务器上，这样就不需要很多的库文件放在服务上　　　了，解决了所有openpyxl的依赖问题，而且库文件包也达到最小化，不至于将python很大的库文件包都放到服务器上




                           # -*- coding: utf-8 -*-
                           from openpyxl import Workbook
                           import sys
                           wb = Workbook()
                           ws = wb.active
                           ws.title = "Pipe_Mto"
                           inputfile = sys.argv[1]
                           #inputfile = 'a.txt'
                           outputfile = sys.argv[2]
                           with open(inputfile) as infile:
                               for line in infile:
                                       ws.append(line.decode('cp936').split("@"))
                                       wb.save(outputfile)       


                                       运行：python xxxxx.py  inputfile(%path%\filename.txt)  outputfile(%path%\filename.xlsx )

