用户手册


1.标准文档的维护及使用
         此Excel文件包含源数据的文件夹地址工作簿、Issue列的标准值工作簿、Department列的标准值工作簿、Carline列的标准值工作簿、Lifecycle列的标准值工作簿、Working/NonWorking列的标准值工作簿。
 

1.1文件夹标准
工作簿：【Folder Path】
	第一次使用该标准时，需在“Path”列下的第一行输入待检测文件的文件夹Raw Data 所在的磁盘路径；如Raw Data文件夹不存在，需新建文件夹并命名为“Raw Data“；
	维护时，如待检测文件的磁盘路径与标准中的路径相同，则无需更改；否则需用新的磁盘路径覆盖标准中的路径；
 
注意：
①红色框内为源数据文件夹在公盘存放位置的全路径；
②路径以【磁盘号】开始，以【 / 】结尾；

1.2 Issue列标准
工作簿：【Issue Standard】
	第一次使用该标准时，需在“Issue FCST”列下输入待检测文件中Issue列所包含的所有与FCST相关的值；在“Issue Actual”列下输入待检测文件中Issue列所包含的所有与Actual相关的值；
	维护时，如待检测文件的Issue列的值与标准中的值相同，则无需更改；否则需用新值覆盖；
 

1.3 Department列标准
工作簿：【Department Standard】
列名为部门名称，每一列为该部门下所属的团队名称；
 
	某个部门新增或减少一个团队，即在对应的列下增加或减少对应的团队名称；如该部门下没有任何团队，则同时删除该部门；
	如新增一个部门，则在第I列添加部门名称以及该部门所属的团队；

1.4 Carline列标准
工作簿：【Carline Standard】
	如需添加新车型，则在“Carline”下添加新车型的型号；
	如需删除旧车型，则在“Carline”下删除旧车型的型号；
 


1.5 Lifecycle列标准
工作簿：【Lifecycle Standard】
	如需添加新记录，则在“Lifecycle”下添加新记录；
	如需删除旧记录，则在“Lifecycle”下删除旧记录；
 
1.6 Working/NonWorking列标准
工作簿：【Working Standard】、【NonWorking Standard】
1.6.1 Working Standard
	Sales Funnel列下为与Working对应的相关值；
	Category列下为与Sales Funnel对应的相关值；
	Category之后的列为与列名对应的相关值；
 
注意：
①如Category列下增加新值，则对应的后面应增加以此值命名的新列名，并在该新列下增加与其对应的关联关系值；
②如需删除Category下某值时，对应的删除后面对应的列及其关联关系值；
1.6.2 NonWorking Standard
同1.6.1；



2.第一次检查
2.1添加数据源
       将待检测的文件放到“Raw Data/Checking”文件夹下；
 
2.2启动检查
       运行检查程序【Checking Step 1.exe】，程序窗口即弹出，此时程序会执行数据检查；窗口出现下图所示时，则检查结束，按任意键退出；
 
3.第一次修改
3.1修改列名错误
       打开“Result/Checking/Header Error/Files”文件夹，查看是否有header错误的文件；如果存在，打开该文件并仅修改错误的列名，修改结束后保存该文件；
 
注意：
①修改绿色矩形内的Header有误的单元格，如红色矩形内的内容；
②橙色矩形内的内容无需修改；
3.2修改Actual错误
       打开“Result/Checking/Actual/Data&Log”文件夹，查看“Act_Error.xlsx”文件，如果存在错误日志，根据错误类型定位到相关列并修改错误数据，待所有错误修改完毕，删除File Name, Exception, Index三列并保存文件；
 

3.3修改FCST错误
       打开“Result/Checking/FCST/Data&Log”文件夹，步骤同3.2；
3.4修改Budget错误
       打开“Result/Checking/Budget/Data&Log”文件夹，步骤同3.2；
4.重复检查
4.1启动重复检查
       运行重复检查程序【Checking Step 2 Or More.exe】，程序窗口即弹出，此时程序会执行第二次数据检查；窗口出现下图所示时，则第二次检查结束时，按任意键退出；
 
4.2修改错误
       重复步骤3，然后重复步骤4.1，直至“Result/Checking/Header Error/Header Error.xlsx”,“ Result/Checking/Actual/Data&Log/Act_Error.xlsx”和“Result/Checking/FCST/Data&Log/FCST_Error.xlsx”都不存在错误日志记录；如果没有错误文件及错误日志行，则数据检查结束，该月份的Actual数据保存于“Act_Summary.xlsx”文件，FCST数据保存于“FCST_Summary.xlsx”文件；
       如果上一步执行了Budget数据的检查，则重复步骤3，然后重复步骤4.1，直至“Result/Checking/Budget/Data&Log/Budget_Error.xlsx” 不存在错误日志记录；
5.按部门团队拆分
5.1添加数据源
       将待检测的文件放到“Raw Data/Parsing”文件夹下；
 
5.2启动拆分
       运行检查程序【Parsing.exe】，程序窗口即弹出，此时程序会执行数据拆分；窗口出现下图所示时，则检查结束，按任意键退出；
 

6. 附录
6.1 几种不易发现的异常
	如果抛出Header Exception，但是检查一遍header之后并未发现有明显异常时，请检查是否有隐藏的sheet,如果有的话将待检测sheet放于所有sheet的最前面；如果上述两项都正确的话, 再重点检查【SC NO.】和【 SC Name】列：前者的“O“是否为大写；后者”SC“前是否有空格(正常为有)；
	如果抛出Data Exception，但是在检查一遍数据之后并未发现明显异常时，请重点检查单词拼写，如“production”和“proudction”；

