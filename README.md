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


依赖的模块
----------


- xlwt  [这里使用的是支持Python3.3的模块 xlwt-future 0.8.0](https://pypi.python.org/pypi/xlwt-future)
- urllib
- time
- re


如何联系我
----------


[我的博客](http://blueandhack.com)


协议
----


遵守MIT协议
