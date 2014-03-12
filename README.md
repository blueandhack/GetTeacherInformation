Get Teacher Information
===================================

此脚本基于Python，可以拉取基于青果软件的教育网络管理系统的教师信息，并生成多个Excel表格，进行教师信息的统计

如何使用？
---------


你需要安装[Python](http://python.org/download/)
确定Python命令可以使用
在终端输入

    python GetInfo.py

若是Windows用户在项目目录文件夹下输入以上命令，例如我的GetInfo.py文件在F盘GetTeacherInformation目录下

    F:
    cd GetTeacherInformation
    python GetInfo.py
    
在运行时按照程序提示输入，即可使用

常见问题
--------


1. 此版基于Python3.3，请使用Python3.3+运行
2. 若出现错误，请按照软件提示输入
2. 若无法获得教师信息，请检查Cookie和Url值是否正确
3. 若Excel文件创建不成功，请确认目录是否存在，或者是否拥有该目录权限
4. **注意：**因为使用了基于Python3+的[xlwt](#依赖的模块)模块，此模块无法追加写入Excel文件，若一次写入超过50条信息就会出现乱码情况，目前我使用的保存方式是，每收集一条信息都放入内存，先不写入文件，然后记录50条信息后一次写入一个Excel工作薄中，最后未达到50条的最后写入最后一个Excel工作薄中（命名方式是Teacher{0-N}.xls）
4. 若以上问题仍然无法解决你的问题请查看[联系方式](#如何联系我)，让我们一起共同完善
5. 增加了对mongodb的支持，如果你需要将数据写入到mongodb中，请首先启动mongodb后再运行本程序
6. **注意：**在这个版本里增加了读取配置文件的功能，请在本项目文件夹下建立一个config.txt的文本文件，在文件内写入：

```
 http://********.com/XXX/
 ASP.NET_SessionId=*****************************
 1
 localhost
```
第一行表示教务系统登陆界面的url
第二行登陆后的Cookie
第三行选择存储模式
第四行 如果第三行填写的是1，那么填写mongodb的配置链接，如果选择的是2那么填写xls的存放路径
**注意！！！**配置文件中每行结尾或者行中不可以有空格，严格按照以上范例来写！

依赖的模块
----------


- xlwt  [这里使用的是支持Python3.3的模块 xlwt-future 0.8.0](https://pypi.python.org/pypi/xlwt-future)
- urllib
- time
- re
- pymongo


如何联系我
----------


[我的博客](http://blueandhack.com)
我的邮箱 blueandhack # gmail.com


协议
----


遵守MIT协议
