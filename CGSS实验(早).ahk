#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.


;——————————————————————
;——————————————————————
;**一、备份CGSS数据命令**
;——————————————————————
;——————————————————————

:*:bfcgsssjz::

;——————————————————————
;0.8:00运行
;——————————————————————

;——————————————————————
;1.获取系统当前日期，日期格式为yyyy-MM-dd和yyyyMMdd
;——————————————————————

	FormatTime, now_date, %A_Now%, yyyy-MM-dd
	Sleep,100
	FormatTime, now_date_wg, %A_Now%, yyyyMMdd
	Sleep,100

;——————————————————————
;2.在CGSSdata目录下新建以系统日期命名的文件夹
;——————————————————————

	;2.1 打开D:\CGSSdata目录
	Run, D:\CGSSdata\
	WinWait, CGSSdata, 
	IfWinNotActive, CGSSdata, , WinActivate, CGSSdata, 
	WinWaitActive, CGSSdata, 
	Sleep,500
	
	;2.2 在CGSSdata目录下新建以系统日期命名的文件夹
	Send, {APPSKEY}
	Sleep,500
	Send,w
	Sleep,100
	Send,f
	Sleep,2000  ;白天可以缩短时间
	Send,%now_date_wg%
	Sleep,100
	Send,{ENTER}
	Sleep,500
	
;——————————————————————
;3.在该文件夹下由bat程序创建三个子目录：访员、问卷、电核
;程序位于D:\CGSSdata目录下，通过搜索、复制、粘贴、运行、删除五步实现上述目标
;——————————————————————
	
	;3.1 ctrl+F命令搜索以bat结尾的文件
	Send,{CTRLDOWN}f{CTRLUP}
	Sleep,1000
	Send, bat
	Sleep,500
	Send,{SHIFT}
	WinWait, bat - “CGSSdata”中的搜索结果, 
	IfWinNotActive, bat - “CGSSdata”中的搜索结果, , WinActivate, bat - “CGSSdata”中的搜索结果, 
	WinWaitActive, bat - “CGSSdata”中的搜索结果, 
	Sleep,500


	;3.2 复制该bat文件
	Send, {TAB}
	Sleep,100
	Send,{TAB}
	Sleep,100
	Send,{RIGHT}
	Sleep,100
	Send,{CTRLDOWN}c{CTRLUP}
	Sleep,500
	Send,{BACKSPACE}
	Sleep,2000
	Send,{ENTER}
	WinWait, %now_date_wg%, 
	IfWinNotActive, %now_date_wg%, , WinActivate, %now_date_wg%, 
	WinWaitActive, %now_date_wg%, 
	Sleep,2000
	
	;3.3 在CGSSdata/系统日期/目录下粘贴该bat文件
	Send,{CTRLDOWN}v{CTRLUP}
	Sleep,2000
	
	;3.4 执行该bat文件，生成电核、访员、问卷三个目录
	Send, {ENTER}
	Sleep,10000
	
	;3.5 删除CGSSdata/系统日期/目录下的bat文件，然后关闭窗口
	Send,{DELETE}
	Sleep,2000
	Send,{ALTDOWN}{F4}{ALTUP}
	Sleep,2000

;——————————————————————
;4.依次打开登录页面和14个下载页面
;——————————————————————

	;4.1 打开limesurvey登录页面并登录limesurvey系统
	Run,http://101.200.178.132/limesurvey_2017/index.php/admin
	WinWait, LimeSurvey_2017 - Google Chrome, 
	IfWinNotActive, LimeSurvey_2017 - Google Chrome, , 	WinActivate, LimeSurvey_2017 - Google Chrome, 
	WinWaitActive, LimeSurvey_2017 - Google Chrome, 
	Sleep,1000
	Send,{LWINDOWN}{UP}{LWINUP}
	Sleep,2000
	;以上两行命令确保窗口最大化
	Send, longzhengfan
	Sleep,500
	Send,{SHIFT}
	Sleep,100
	Send,{TAB}
	Sleep,100
	Send,123456
	Sleep,100
	Send,{ENTER}
	Sleep,100

	;/*选定像素点的颜色及位置信息
	;color[8 of 63]: 0x0BEDCB
	;ErrorLevel[1 of 3]: 0
	;MouseX[3 of 3]: 535
	;MouseY[3 of 3]: 241
	;*/
	
	;4.2 判断是否登录了limesurvey系统
	;方法是判断屏幕上(501,216)和(553,271)两个坐标点所确定的矩形区域内是否存在颜色为0x0BEDCB的像素块，如果存在，循环终止，如果不存在，循环继续。当登录了limesurvey系统时，该矩形区域内是limesurvey系统的大logo图标的一小部分，其中有颜色为0x0BEDCB的像素块。因此当登录了limesurvey系统时，PixelSearch函数就会找到颜色为0x0BEDCB的像素块，循环就会终止，继而执行下面的命令。关于PixelSearch函数的用法见AutoHotkey帮助文件。
	;颜色0x0BEDCB的获取来自于AutoHotkey的一个示例脚本，现置于linshi.ahk中，由鼠标所在位置以及ctrl+shift+Z热键联合触发
	Loop
	{
		PixelSearch, LogoX, LogoY, 501, 216, 553, 271, 0x0BEDCB, 2, Fast
		Sleep,1000
		If ErrorLevel
			continue
		else
			break	
	}
	
	;4.3 变量 下载地址头部1、头部2
	sjdz1 = http://101.200.178.132/limesurvey_2017/index.php/admin/export/sa/exportspss/sid/
	sjdz2 =	http://101.200.178.132/limesurvey_2017/index.php/admin/export/sa/exportresults/surveyid/
	
	;4.4 打开电核数据下载页面（spss,dta,csv,xlsx,大表）
	Loop 2
	{
		Run,%sjdz1%637954
		Sleep, 1000
	}
	Loop 2
	{
		Run,%sjdz2%637954
		Sleep, 1000
	}
	Run,https://docs.google.com/spreadsheets/d/1tfRQ_n3-4bTEemeqWZlhvLNlV-eAjKCv9dpNVs7ETSc/edit?usp=drive_web
	Sleep,1000
	
	;4.5 打开访员数据下载页面（spss,dta,csv,xlsx）
	Loop 2
	{
		Run,%sjdz1%252672
		Sleep, 1000
	}
	Loop 2
	{
		Run,%sjdz2%252672
		Sleep, 1000
	}

	
	;4.6 打开问卷数据下载页面（spss,dta,csv,大表,xlsx）
	Loop 2
	{
		Run,%sjdz1%963159
		Sleep, 1000
	}
	Run,%sjdz2%963159
	Sleep, 1000
	Run,   https://docs.google.com/spreadsheets/d/1-pBKYvO8Fd32Rmk76ghDrIcKPC4R5renbWGtE5-DYLg/edit#gid=0
	Sleep, 1000
	Run,%sjdz2%963159
	Sleep, 10000

;——————————————————————
;5.下载第1个文件
;——————————————————————	

	;5.1 切换到第1个下载页面
	Send,{CAPSLOCK}
	Sleep,1000
	Send,{K}
	Sleep,2000
	Send,{K}
	Sleep,2000
	
	;5.2 下载第1个文件637954.spss到D:\CGSSdata\系统日期\电核
	Send,{f}
	Sleep,500
	Send,ss
	WinWait, 另存为, 
	IfWinNotActive, 另存为, , WinActivate, 另存为, 
	WinWaitActive, 另存为,
	Sleep,1000
	Loop 6
	{
		Send,{SHIFTDOWN}{TAB}{SHIFTUP}
		Sleep,500
	}
	Sleep,500
	Send,{ENTER}
	Sleep,1000
	Send,D:\CGSSdata\{SHIFTDOWN}{SHIFTUP}
	Sleep,500
	Send,%now_date_wg%
	Sleep,500
	Send,\电核
	Sleep,500
	Send,{ENTER}
	Loop 9
	{
		Send,{TAB}
		Sleep,500
	}
	Send,{ENTER}
	Sleep,2000
	
;——————————————————————
;6.下载第2个文件
;——————————————————————	

	;6.1 切换到第2个下载页面
	Send,{CAPSLOCK}
	Sleep,1000
	Send,{K}
	Sleep,2000
	
	;6.2 下载第2个文件637954.dta到D:\CGSSdata\系统日期\电核
	Send,{f}
	Sleep,500
	Send,w
	WinWait, 另存为, 
	IfWinNotActive, 另存为, , WinActivate, 另存为, 
	WinWaitActive, 另存为,
	Sleep,1000
	Send,{ENTER}
	Sleep,2000
	
;——————————————————————
;7.下载第3个文件
;——————————————————————	

	;7.1 切换到第3个下载页面
	Send,{CAPSLOCK}
	Sleep,1000
	Send,{K}
	Sleep,2000
	
	;7.2 下载第3个文件637954.csv到D:\CGSSdata\系统日期\电核
	Loop 19
	{
		Send,{DOWN}
		Sleep,500
	}
	Send,f
	Sleep,500
	Send,e
	Sleep,500
	Send,f
	Sleep,500
	Send,sk
	Sleep,500
	Send,f
	Sleep,500
	Send,sl
	Sleep,500
	Send,f
	Sleep,500
	Send,p
	Sleep,500
	Send,f
	Sleep,500
	Send,sw
	Sleep,500
	Send,{CTRLDOWN}a{CTRLUP}
	Sleep,500
	Send,{TAB}
	Sleep,500
	Send,f
	Sleep,500
	Send,a
	Sleep,500
	WinWait, 另存为, 
	IfWinNotActive, 另存为, , WinActivate, 另存为, 
	WinWaitActive, 另存为,
	Sleep,1000
	Send,{ENTER}
	Sleep,2000

;——————————————————————
;8.下载第4个文件
;——————————————————————	

	;8.1 切换到第4个下载页面
	Send,{CAPSLOCK}
	Sleep,1000
	Send,{K}
	Sleep,2000
	
	;下载第4个文件637954.xlsx到D:\CGSSdata\系统日期\电核
	Send,f
	Sleep,500
	Send,sk
	Sleep,500
	MouseClick, left,  585,  460
	Sleep,500
	Loop 19
	{
		Send,{DOWN}
		Sleep,500
	}
	Send,f
	Sleep,500
	Send,e
	Sleep,500
	Send,f
	Sleep,500
	Send,sk
	Sleep,500
	Send,f
	Sleep,500
	Send,sl
	Sleep,500
	Send,f
	Sleep,500
	Send,p
	Sleep,500
	Send,f
	Sleep,500
	Send,sw
	Sleep,500
	Send,{CTRLDOWN}a{CTRLUP}
	Sleep,500
	Send,{TAB}
	Sleep,500
	Send,f
	Sleep,1000
	Send,a
	Sleep,500
	WinWait, 另存为, 
	IfWinNotActive, 另存为, , WinActivate, 另存为, 
	WinWaitActive, 另存为,
	Sleep,1000
	Send,{ENTER}
	Sleep,2000

;——————————————————————
;9.下载第5个文件
;——————————————————————	

	;9.1 切换到第5个下载页面
	Send,{CAPSLOCK}
	Sleep,1000
	Send,{K}
	Sleep,8000
	
	;9.2 下载第5个文件2017电话核查大表.xlsx到D:\CGSSdata\系统日期\电核
	MouseClick, left,  981,  126
	Sleep, 1000
	MouseClick, left, 80, 144
	Sleep, 1000
	Loop 7
	{
		Send,{UP}
		Sleep,500
	}
	Send,{RIGHT}
	Sleep,500
	Send,{ENTER}
	WinWait, 另存为, 
	IfWinNotActive, 另存为, , WinActivate, 另存为, 
	WinWaitActive, 另存为,
	Sleep,1000
	Send,{ENTER}	
	Sleep,2000
	MouseClick, left,  740,  122
	Sleep,2000
	
;——————————————————————
;10.下载第6个文件
;——————————————————————	

	;10.1 切换到第6个下载页面
	Send,{CAPSLOCK}
	Sleep,1000
	Send,{K}
	Sleep,2000
	
	;10.2 下载第6个文件252672.spss到D:\CGSSdata\系统日期\访员
	Send,{f}bfcgsss
	Sleep,500
	Send,ss
	WinWait, 另存为, 
	IfWinNotActive, 另存为, , WinActivate, 另存为, 
	WinWaitActive, 另存为,
	Sleep,1000
	Loop 6
	{
		Send,{SHIFTDOWN}{TAB}{SHIFTUP}
		Sleep,500
	}
	Sleep,500
	Send,{ENTER}
	Sleep,1000
	Send,D:\CGSSdata\{SHIFTDOWN}{SHIFTUP}
	Sleep,500
	Send,%now_date_wg%
	Sleep,500
	Send,\访员
	Sleep,500
	Send,{ENTER}
	Loop 9
	{
		Send,{TAB}
		Sleep,500
	}
	Send,{ENTER}
	Sleep,2000
	
;——————————————————————
;11.下载第7个文件
;——————————————————————	

	;11.1 切换到第7个下载页面
	Send,{CAPSLOCK}
	Sleep,1000
	Send,{K}
	Sleep,2000
	
	;11.2 下载第7个文件252672.dta到D:\CGSSdata\系统日期\访员
	Send,{f}
	Sleep,500
	Send,w
	WinWait, 另存为, 
	IfWinNotActive, 另存为, , WinActivate, 另存为, 
	WinWaitActive, 另存为,
	Sleep,1000
	Send,{ENTER}
	Sleep,2000
	
;——————————————————————
;12.下载第8个文件
;——————————————————————	

	;12.1 切换到第8个下载页面
	Send,{CAPSLOCK}
	Sleep,1000
	Send,{K}
	Sleep,2000
	
	;12.2 下载第8个文件252672.csv到D:\CGSSdata\系统日期\访员
	Loop 19
	{
		Send,{DOWN}
		Sleep,500
	}
	Send,f
	Sleep,500
	Send,e
	Sleep,500
	Send,f
	Sleep,500
	Send,l
	Sleep,500
	Send,f
	Sleep,500
	Send,m
	Sleep,500
	Send,f
	Sleep,500
	Send,ss
	Sleep,500
	Send,f
	Sleep,500
	Send,a
	Sleep,500
	WinWait, 另存为, 
	IfWinNotActive, 另存为, , WinActivate, 另存为, 
	WinWaitActive, 另存为,
	Sleep,1000
	Send,{ENTER}
	Sleep,2000
	
;——————————————————————
;13.下载第9个文件
;——————————————————————	

	;13.1 切换到第9个下载页面
	Send,{CAPSLOCK}
	Sleep,1000
	Send,{K}
	Sleep,2000
	
	;13.2 下载第9个文件252672.xlsx到D:\CGSSdata\系统日期\
	Send,f
	Sleep,500
	Send,sk
	Sleep,500
	MouseClick, left,  585,  460
	Sleep,500
	Loop 19
	{
		Send,{DOWN}
		Sleep,500
	}
	Send,f
	Sleep,500
	Send,se
	Sleep,500
	Send,f
	Sleep,500
	Send,l
	Sleep,500
	Send,f
	Sleep,500
	Send,m
	Sleep,500
	Send,f
	Sleep,500
	Send,ss
	Sleep,500
	Send,f
	Sleep,500
	Send,sa
	Sleep,500
	WinWait, 另存为, 
	IfWinNotActive, 另存为, , WinActivate, 另存为, 
	WinWaitActive, 另存为,
	Sleep,1000
	Send,{ENTER}
	Sleep,2000
	
;——————————————————————
;14.下载第10个文件
;——————————————————————	

	;14.1 切换到第10个下载页面
	Send,{CAPSLOCK}
	Sleep,1000
	Send,{K}
	Sleep,2000
	
	;14.2 下载第10个文件963159.spss到D:\CGSSdata\系统日期\问卷
	Send,{f}
	Sleep,500
	Send,ss
	WinWait, 另存为, 
	IfWinNotActive, 另存为, , WinActivate, 另存为, 
	WinWaitActive, 另存为,
	Sleep,1000
	Loop 6
	{
		Send,{SHIFTDOWN}{TAB}{SHIFTUP}
		Sleep,500
	}
	Sleep,500
	Send,{ENTER}
	Sleep,1000
	Send,D:\CGSSdata\{SHIFTDOWN}{SHIFTUP}
	Sleep,500
	Send,%now_date_wg%
	Sleep,500
	Send,\问卷
	Sleep,500
	Send,{ENTER}
	Loop 9
	{
		Send,{TAB}
		Sleep,500
	}
	Send,{ENTER}
	Sleep,2000

;——————————————————————
;15.下载第11个文件
;——————————————————————	

	;15.1 切换到第11个下载页面
	Send,{CAPSLOCK}
	Sleep,1000
	Send,{K}
	Sleep,2000
	
	;15.2 下载第11个文件963159.dta到D:\CGSSdata\系统日期\问卷
	Send,{f}
	Sleep,500
	Send,w
	WinWait, 另存为, 
	IfWinNotActive, 另存为, , WinActivate, 另存为, 
	WinWaitActive, 另存为,
	Sleep,1000
	Send,{ENTER}
	Sleep,2000	
	
;——————————————————————
;16.下载第12个文件
;——————————————————————	

	;16.1 切换到第12个下载页面
	Send,{CAPSLOCK}
	Sleep,1000
	Send,{K}
	Sleep,2000
	
	;16.2 下载第12个文件963159.csv到D:\CGSSdata\系统日期\问卷
	Loop 19
	{
		Send,{DOWN}
		Sleep,500
	}
	Send,f
	Sleep,500
	Send,e
	Sleep,500
	Send,f
	Sleep,500
	Send,sk
	Sleep,500
	Send,f
	Sleep,500
	Send,sl
	Sleep,500
	Send,f
	Sleep,500
	Send,p
	Sleep,500
	Send,f
	Sleep,500
	Send,sw
	Sleep,500
	Send,{CTRLDOWN}a{CTRLUP}
	Sleep,500
	Send,{TAB}
	Sleep,500
	Send,f
	Sleep,500
	Send,a
	Sleep,500
	WinWait, 另存为, 
	IfWinNotActive, 另存为, , WinActivate, 另存为, 
	WinWaitActive, 另存为,
	Sleep,1000
	Send,{ENTER}
	Sleep,2000
	
;——————————————————————
;17.下载第13个文件
;——————————————————————	

	;17.1 切换到第13个下载页面
	Send,{CAPSLOCK}
	Sleep,1000
	Send,{K}
	Sleep,8000
	
	;17.2 下载第13个文件2017进度大表.xlsx到D:\CGSSdata\系统日期\问卷
	MouseClick, left,  981,  126
	Sleep, 1000
	MouseClick, left, 80, 144
	Sleep, 1000
	Loop 7
	{
		Send,{UP}
		Sleep,500
	}
	Send,{RIGHT}
	Sleep,500
	Send,{ENTER}
	WinWait, 另存为, 
	IfWinNotActive, 另存为, , WinActivate, 另存为, 
	WinWaitActive, 另存为,
	Sleep,1000
	Send,{ENTER}
	Sleep,2000
	MouseClick, left,  424,  162
	Sleep,1000
	
;——————————————————————
;18.下载第14个文件
;——————————————————————	

	;18.1 切换到第14个下载页面
	Send,{CAPSLOCK}
	Sleep,1000
	Send,{K}
	Sleep,2000
	
	;18.2 下载第14个文件963159.xlsx到D:\CGSSdata\系统日期\问卷
	Send,f
	Sleep,500
	Send,sk
	Sleep,500
	MouseClick, left,  585,  460
	Sleep,500
	Loop 19
	{
		Send,{DOWN}
		Sleep,500
	}
	Send,f
	Sleep,500
	Send,e
	Sleep,500
	Send,f
	Sleep,500
	Send,sk
	Sleep,500
	Send,f
	Sleep,500
	Send,sl
	Sleep,500
	Send,f
	Sleep,500
	Send,p
	Sleep,500
	Send,f
	Sleep,500
	Send,sw
	Sleep,500
	Send,{CTRLDOWN}a{CTRLUP}
	Sleep,500
	Send,{TAB}
	Sleep,500
	Send,f
	Sleep,500
	Send,a
	Sleep,500
	WinWait, 另存为, 
	IfWinNotActive, 另存为, , WinActivate, 另存为, 
	WinWaitActive, 另存为,
	Sleep,1000
	Send,{ENTER}

	Sleep,10000
	MouseMove, 152, 721
	
	;下载完成后像素信息
	;	color[8 of 63]: 0xDFDFDF
	;	ErrorLevel[1 of 3]: 0
	;	MouseX[2 of 3]: 47
	;	MouseY[3 of 3]: 719
	
	;18.3 判断文件963159.xlsx是否下载完毕
	;方法是判断屏幕上(47,719)坐标点上的像素块颜色是否为0xDFDFDF，如果存在，循环终止，如果不存在，循环继续。当文件正在下载时，该坐标点上像素块为浅蓝色，而当文件下载结束时，该坐标点上像素块为灰色0xDFDFDF。因此当文件下载完毕时，循环便会终止，程序将紧接着执行下面的命令。
	;颜色0xDFDFDF的获取来自于AutoHotkey的一个示例脚本，现置于linshi.ahk中，由鼠标所在位置以及ctrl+shift+Z热键联合触发
	Loop
	{
		PixelSearch, DowlX, DowlY, 47, 719, 47, 719, 0xDFDFDF, 0, Fast
		Sleep,1000
		If ErrorLevel
			continue
		else
			break
	}
	
	;18.4 退出系统并关闭浏览器
	MouseClick, left, 1058, 535
	Sleep,500
	Loop 22
	{
		Send,{UP}
		Sleep,500
	}
	MouseClick, left, 1236, 123 
	Sleep,500
	Loop 2
	{
		Send,{DOWN}
		Sleep,500
	}
	Send,{ENTER}
	Sleep,4000
	Send,{ALTDOWN}{F4}{ALTUP}
	Sleep,1000


;——————————————————————
;7.更改电核大表属性，将sheet1另存为dhdb.csv，将sheet2另存为fjbh.csv
;——————————————————————	
		
	;7.1 打开 D:\CGSSdata\系统日期\电核\ 目录
	Run,D:\CGSSdata\%now_date_wg%\电核\
	WinWait, 电核, 
	IfWinNotActive, 电核, , WinActivate, 电核, 
	WinWaitActive, 电核, 
	Sleep,500
	
	;7.2 定位光标到文件：.\CGSS2017电话核查样本更新.xlsx
	Send,{CTRLDOWN}f{CTRLUP}
	Sleep,1000
	Send,电
	WinWait, 电 - “电核”中的搜索结果, 
	IfWinNotActive, 电 - “电核”中的搜索结果, , WinActivate, 电 - “电核”中的搜索结果, 
	WinWaitActive, 电 - “电核”中的搜索结果,
	Sleep,1000
	Send,{TAB}
	Sleep,500
	Send,{TAB}
	Sleep,500
	Send,{UP}
	Sleep,500
	
	;7.3 修改文件属性，取消锁定
	Send, {APPSKEY}
	Sleep,2000
	Send, r
	WinWait, CGSS2017电话核查样本更新.xlsx 属性, 
	IfWinNotActive, CGSS2017电话核查样本更新.xlsx 属性, , WinActivate, CGSS2017电话核查样本更新.xlsx 属性, 
	WinWaitActive, CGSS2017电话核查样本更新.xlsx 属性, 
	Sleep,500
	Send, k
	Sleep,500
	Send,{ENTER}
	Sleep,500
	Send,{ALTDOWN}{F4}{ALTUP}
	Sleep,1000

	;7.4 打开文件
	Run,D:\CGSSdata\%now_date_wg%\电核\CGSS2017电话核查样本更新.xlsx
	WinWait, CGSS2017电话核查样本更新.xlsx - Excel, 
	IfWinNotActive, CGSS2017电话核查样本更新.xlsx - Excel, , WinActivate, CGSS2017电话核查样本更新.xlsx - Excel, 
	WinWaitActive, CGSS2017电话核查样本更新.xlsx - Excel,
	Sleep,2000
	Send,{LWINDOWN}{UP}{LWINUP}
	Sleep,2000
	;以上两行命令确保窗口最大化
	IfWinNotActive, CGSS2017电话核查样本更新.xlsx - Excel, , WinActivate, CGSS2017电话核查样本更新.xlsx - Excel, 
	WinWaitActive, CGSS2017电话核查样本更新.xlsx - Excel,

	;7.5 定位光标到名为“token”的列
	;方法是通过搜索定位到名为“更新日期”的列，“token”列在其右侧
	Send,{CTRLDOWN}f{CTRLUP}
	WinWait, 查找和替换, 
	IfWinNotActive, 查找和替换, , WinActivate, 查找和替换, 
	WinWaitActive, 查找和替换,
	Sleep,2000
	Send,更新日期
	Sleep,4000
	Send,{TAB}
	Sleep,2000
	Send,f
	Sleep,1000
	Loop 4
	{
		Send,{TAB}
		Sleep,500
	}
	Send,{ENTER}
	WinWait, CGSS2017电话核查样本更新.xlsx - Excel, 
	IfWinNotActive, CGSS2017电话核查样本更新.xlsx - Excel, , WinActivate, CGSS2017电话核查样本更新.xlsx - Excel, 
	WinWaitActive, CGSS2017电话核查样本更新.xlsx - Excel,
	Sleep,1000
	Send,{RIGHT}
	Sleep,500
	
	;7.6 更改“token”列的数据格式为数值型
	Send,{CTRLDOWN}{SHIFTDOWN}{DOWN}{SHIFTUP}{CTRLUP}
	Sleep,2000
	Send,{APPSKEY}
	Sleep,2000
	Send,f
	WinWait, 设置单元格格式, 
	IfWinNotActive, 设置单元格格式, , WinActivate, 设置单元格格式, 
	WinWaitActive, 设置单元格格式,
	Sleep,500
	Send,{TAB}
	Sleep,500
	Send,{DOWN}
	Sleep,500
	Send,{DOWN}
	Sleep,500
	Send,{TAB}
	Sleep,500
	Send,0
	Sleep,500
	Send,{ENTER}
	Sleep,10000
	WinWait, CGSS2017电话核查样本更新.xlsx - Excel, 
	IfWinNotActive, CGSS2017电话核查样本更新.xlsx - Excel, , WinActivate, CGSS2017电话核查样本更新.xlsx - Excel, 
	WinWaitActive, CGSS2017电话核查样本更新.xlsx - Excel,
	
	;7.7 将sheet1另存为“D:\CGSS调查问卷统计\数据质控测试文档\系统日期 dhdb.csv”
	Sleep,500
	Send, {ALT}
	Sleep,2000
	Send,f
	Sleep,3000
	Send,a
	Sleep,3000
	Send,o
	WinWait, 另存为, 
	IfWinNotActive, 另存为, , WinActivate, 另存为, 
	WinWaitActive, 另存为,
	Sleep,1000
	Loop 6
	{
		Send,{SHIFTDOWN}{TAB}{SHIFTUP}
		Sleep,500
	}
	Send,{ENTER}
	Sleep,500
	Send,D:\CGSS{SHIFTDOWN}{SHIFTUP}
	Sleep,2000
	Send,调查问卷统计\
	Sleep,1000
	Send,数据质控测试文档
	Sleep,1000
	Send,{ENTER}
	Sleep,500
	Loop,6
	{
		Send, {TAB}
		Sleep,500
	}
	Send,%now_date%{SPACE}
	Sleep,500
	Send,dhdb
	Sleep,500
	Send,{TAB}
	Sleep,500
	Send,c
	Sleep,500
	Send,{ENTER}
	WinWait, Microsoft Excel, 
	IfWinNotActive, Microsoft Excel, , WinActivate, Microsoft Excel, 
	WinWaitActive, Microsoft Excel, 
	Sleep,500
	Send,{ENTER}
	WinWait, Microsoft Excel, 
	IfWinNotActive, Microsoft Excel, , WinActivate, Microsoft Excel, 
	WinWaitActive, Microsoft Excel, 
	Sleep,500
	Send,{ENTER}
	WinWait, %now_date% dhdb.csv - Excel, 
	IfWinNotActive, %now_date% dhdb.csv - Excel, , WinActivate, %now_date% dhdb.csv - Excel, 
	WinWaitActive, %now_date% dhdb.csv - Excel,
	
	;7.8 切换到第2个sheet,废卷情况汇总
	Sleep, 2000
	Send, {CTRLDOWN}{PGDN}{CTRLUP}
	Sleep, 2000
	
	;7.9 定位光标到名为“样本编号”的列
	;方法是通过搜索定位到名为“记录日期”的列，“样本编号”列在其右侧
	Send,{CTRLDOWN}f{CTRLUP}
	WinWait, 查找和替换, 
	IfWinNotActive, 查找和替换, , WinActivate, 查找和替换, 
	WinWaitActive, 查找和替换,
	Sleep,1000
	Send,记录日期
	Sleep,1000
	Send,{TAB}
	Sleep,500
	Send,f
	Sleep,500
	Loop 4
	{
		Send,{TAB}
		Sleep,500
	}
	Send,{ENTER}
	WinWait, %now_date% dhdb.csv - Excel, 
	IfWinNotActive, %now_date% dhdb.csv - Excel, , WinActivate, %now_date% dhdb.csv - Excel, 
	WinWaitActive, %now_date% dhdb.csv - Excel,
	Sleep,500
	Send,{RIGHT}
	Sleep,500
	
	;7.10 更改“样本编号”列的数据格式为数值型
	Send,{CTRLDOWN}{SHIFTDOWN}{DOWN}{SHIFTUP}{CTRLUP}
	Sleep,3000
	Send,{APPSKEY}
	Sleep,2000
	Send,f
	WinWait, 设置单元格格式, 
	IfWinNotActive, 设置单元格格式, , WinActivate, 设置单元格格式, 
	WinWaitActive, 设置单元格格式,
	Sleep,500
	Send,{TAB}
	Sleep,500
	Send,{DOWN}
	Sleep,500
	Send,{TAB}
	Sleep,500
	Send,0
	Sleep,500
	Send,{ENTER}
	Sleep,10000
	WinWait, %now_date% dhdb.csv - Excel, 
	IfWinNotActive, %now_date% dhdb.csv - Excel, , WinActivate, %now_date% dhdb.csv - Excel, 
	WinWaitActive, %now_date% dhdb.csv - Excel,
	
	;7.11 将sheet2另存为“D:\CGSS调查问卷统计\数据质控测试文档\系统日期 fjbh.csv”
	Sleep,1000
	Send, {ALT}
	Sleep,2000
	Send,f
	Sleep,3000
	Send,a
	Sleep,3000
	Send,o
	WinWait, 另存为, 
	IfWinNotActive, 另存为, , WinActivate, 另存为, 
	WinWaitActive, 另存为,
	Sleep,1000
	Loop 6
	{
		Send,{SHIFTDOWN}{TAB}{SHIFTUP}
		Sleep,500
	}
	Sleep,500
	Send,{ENTER}
	Sleep,1000
	Send,D:\CGSS
	Sleep,1000
	Send,调查问卷统计\
	Sleep,1000
	Send,数据质控测试文档
	Sleep,1000
	Send,{ENTER}
	Sleep,500
	Loop,6
	{
		Send, {TAB}
		Sleep,500
	}
	Send,%now_date%{SPACE}
	Sleep,500
	Send,fjbh
	Sleep,500
	Send,{TAB}
	Sleep,500
	Send,c
	Sleep,500
	Send,{ENTER}
	WinWait, Microsoft Excel, 
	IfWinNotActive, Microsoft Excel, , WinActivate, Microsoft Excel, 
	WinWaitActive, Microsoft Excel, 
	Sleep,500
	Send,{ENTER}
	WinWait, Microsoft Excel, 
	IfWinNotActive, Microsoft Excel, , WinActivate, Microsoft Excel, 
	WinWaitActive, Microsoft Excel, 
	Sleep,500
	Send,{ENTER}
	WinWait, %now_date% fjbh.csv - Excel, 
	IfWinNotActive, %now_date% fjbh.csv - Excel, , WinActivate, %now_date% fjbh.csv - Excel, 
	WinWaitActive, %now_date% fjbh.csv - Excel,
	Sleep,500

	;7.12 切换到第5个sheet,高缺失率问卷,将sheet5另存为“D:\CGSS调查问卷统计\数据质控测试文档\系统日期 gqslwj.csv”
	Sleep,2000
	Loop 3
	{
		Send, {CTRLDOWN}{PGDN}{CTRLUP}
		Sleep, 500
	}
	Sleep,2000
	
	Send,{CTRLDOWN}f{CTRLUP}
	WinWait, 查找和替换, 
	IfWinNotActive, 查找和替换, , WinActivate, 查找和替换, 
	WinWaitActive, 查找和替换,
	Sleep,2000
	Send,样本编号
	Sleep,2000
	Send,{TAB}
	Sleep,1000
	Send,f
	Sleep,1000
	Loop 4
	{
		Send,{TAB}
		Sleep,500
	}
	Send,{ENTER}
	WinWait, %now_date% fjbh.csv - Excel, 
	IfWinNotActive, %now_date% fjbh.csv - Excel, , WinActivate, %now_date% fjbh.csv - Excel, 
	WinWaitActive, %now_date% fjbh.csv - Excel,
	Sleep,1000
	
	;更改“样本编号”列的数据格式为数值型
	Send,{CTRLDOWN}{SHIFTDOWN}{DOWN}{SHIFTUP}{CTRLUP}
	Sleep,5000
	Send,{APPSKEY}
	Sleep,2000
	Send,f
	WinWait, 设置单元格格式, 
	IfWinNotActive, 设置单元格格式, , WinActivate, 设置单元格格式, 
	WinWaitActive, 设置单元格格式,
	Sleep,500
	Send,{TAB}
	Sleep,500
	Send,{DOWN}
	Sleep,500
	Send,{TAB}
	Sleep,500
	Send,0
	Sleep,500
	Send,{ENTER}
	Sleep,10000
	WinWait, %now_date% fjbh.csv - Excel, 
	IfWinNotActive, %now_date% fjbh.csv - Excel, , WinActivate, %now_date% fjbh.csv - Excel, 
	WinWaitActive, %now_date% fjbh.csv - Excel,
	Sleep,2000
	
	Send, {ALT}
	Sleep,2000
	Send,f
	Sleep,3000
	Send,a
	Sleep,3000
	Send,o
	WinWait, 另存为, 
	IfWinNotActive, 另存为, , WinActivate, 另存为, 
	WinWaitActive, 另存为,
	Sleep,1000
	Loop 6
	{
		Send,{SHIFTDOWN}{TAB}{SHIFTUP}
		Sleep,500
	}
	Sleep,500
	Send,{ENTER}
	Sleep,1000
	Send,D:\CGSS
	Sleep,1000
	Send,调查问卷统计\
	Sleep,1000
	Send,数据质控测试文档
	Sleep,1000
	Send,{ENTER}
	Sleep,500
	Loop,6
	{
		Send, {TAB}
		Sleep,500
	}
	Send,%now_date%{SPACE}
	Sleep,500
	Send,gqslwj
	Sleep,500
	Send,{TAB}
	Sleep,500
	Send,c
	Sleep,500
	Send,{ENTER}
	WinWait, Microsoft Excel, 
	IfWinNotActive, Microsoft Excel, , WinActivate, Microsoft Excel, 
	WinWaitActive, Microsoft Excel, 
	Sleep,500
	Send,{ENTER}
	WinWait, Microsoft Excel, 
	IfWinNotActive, Microsoft Excel, , WinActivate, Microsoft Excel, 
	WinWaitActive, Microsoft Excel, 
	Sleep,500
	Send,{ENTER}
	WinWait, %now_date% gqslwj.csv - Excel, 
	IfWinNotActive, %now_date% gqslwj.csv - Excel, , WinActivate, %now_date% gqslwj.csv - Excel, 
	WinWaitActive, %now_date% gqslwj.csv - Excel,
	Sleep,500
	
	;7.13 切换到第6个sheet,低电话回收率访员,将sheet6另存为“D:\CGSS调查问卷统计\数据质控测试文档\系统日期 hsldfy.csv”
	Sleep, 2000
	Send, {CTRLDOWN}{PGDN}{CTRLUP}
	Sleep, 1000
	Send, {ALT}
	Sleep,2000
	Send,f
	Sleep,3000
	Send,a
	Sleep,3000
	Send,o
	WinWait, 另存为, 
	IfWinNotActive, 另存为, , WinActivate, 另存为, 
	WinWaitActive, 另存为,
	Sleep,1000
	Loop 6
	{
		Send,{SHIFTDOWN}{TAB}{SHIFTUP}
		Sleep,500
	}
	Sleep,500
	Send,{ENTER}
	Sleep,1000
	Send,D:\CGSS
	Sleep,1000
	Send,调查问卷统计\
	Sleep,1000
	Send,数据质控测试文档
	Sleep,1000
	Send,{ENTER}
	Sleep,500
	Loop,6
	{
		Send, {TAB}
		Sleep,500
	}
	Send,%now_date%{SPACE}
	Sleep,500
	Send,hsldfy
	Sleep,500
	Send,{TAB}
	Sleep,500
	Send,c
	Sleep,500
	Send,{ENTER}
	WinWait, Microsoft Excel, 
	IfWinNotActive, Microsoft Excel, , WinActivate, Microsoft Excel, 
	WinWaitActive, Microsoft Excel, 
	Sleep,500
	Send,{ENTER}
	WinWait, Microsoft Excel, 
	IfWinNotActive, Microsoft Excel, , WinActivate, Microsoft Excel, 
	WinWaitActive, Microsoft Excel, 
	Sleep,500
	Send,{ENTER}
	WinWait, %now_date% hsldfy.csv - Excel, 
	IfWinNotActive, %now_date% hsldfy.csv - Excel, , WinActivate, %now_date% hsldfy.csv - Excel, 
	WinWaitActive, %now_date% hsldfy.csv - Excel,
	Sleep,500
	
	;7.14 退出2017电核样本更新.xlsx
	Send, {ALTDOWN}{F4}{ALTUP}
	WinWait, Microsoft Excel, 
	IfWinNotActive, Microsoft Excel, , WinActivate, Microsoft Excel, 
	WinWaitActive, Microsoft Excel, 
	Sleep,500
	Send, {RIGHT}
	Sleep,500
	Send,{ENTER}
	Sleep,1000
	
;——————————————————
;8.更改问卷文件属性，将第一个sheet另存为"系统日期 wjda.csv"
;——————————————————
	
	;8.1 打开 D:\CGSSdata\系统日期\问卷\
	Run,D:\CGSSdata\%now_date_wg%\问卷\
	WinWait, 问卷, 
	IfWinNotActive, 问卷, , WinActivate, 问卷, 
	WinWaitActive, 问卷, 
	Sleep,500
	
	;8.2 定位光标到文件：.\results-survey963159.xlsx
	Send,{CTRLDOWN}f{CTRLUP}
	Sleep,1000
	Send,{SHIFT}
	Sleep,500
	Send,results-survey963159.xlsx
	WinWait, results-survey963159.xlsx - “问卷”中的搜索结果, 
	IfWinNotActive, results-survey963159.xlsx - “问卷”中的搜索结果, , WinActivate, results-survey963159.xlsx - “问卷”中的搜索结果, 
	WinWaitActive, results-survey963159.xlsx - “问卷”中的搜索结果,
	Sleep,1000
	Send,{TAB}
	Sleep,500
	Send,{TAB}
	Sleep,500
	Send,{UP}
	Sleep,500
	
	;8.3 修改文件属性，取消锁定
	Send, {APPSKEY}
	Sleep,2000
	Send, r
	WinWait, results-survey963159.xlsx 属性, 
	IfWinNotActive, results-survey963159.xlsx 属性, , WinActivate, results-survey963159.xlsx 属性, 
	WinWaitActive, results-survey963159.xlsx 属性, 
	Sleep,500
	Send, k
	Sleep,500
	Send,{ENTER}
	Sleep,500
	Send,{ALTDOWN}{F4}{ALTUP}
	Sleep,1000
	
	;8.4 打开文件
	Run,D:\CGSSdata\%now_date_wg%\问卷\results-survey963159.xlsx
	WinWait, results-survey963159.xlsx - Excel, 
	IfWinNotActive, results-survey963159.xlsx - Excel, , WinActivate, results-survey963159.xlsx - Excel, 
	WinWaitActive, results-survey963159.xlsx - Excel,
	Sleep,500
	
	;8.5 将sheet1另存为“D:\CGSS调查问卷统计\数据质控测试文档\系统日期 wjda.csv”
	Send, {ALT}
	Sleep,2000
	Send,f
	Sleep,3000
	Send,a
	Sleep,3000
	Send,o
	WinWait, 另存为, 
	IfWinNotActive, 另存为, , WinActivate, 另存为, 
	WinWaitActive, 另存为,
	Sleep,500
	Loop 6
	{
		Send,{SHIFTDOWN}{TAB}{SHIFTUP}
		Sleep,500
	}
	Send,{ENTER}
	Sleep,500
	Send,D:\CGSS{SHIFTDOWN}{SHIFTUP}
	Sleep,500
	Send,调查问卷统计\
	Sleep,500
	Send,数据质控测试文档
	Sleep,500
	Send,{ENTER}
	Sleep,500
	Loop,6
	{
		Send, {TAB}
		Sleep,500
	}
	Send,%now_date%{SPACE}
	Sleep,500
	Send,wjda
	Sleep,500
	Send,{TAB}
	Sleep,500
	Send,c
	Sleep,500
	Send,{ENTER}
	WinWait, Microsoft Excel, 
	IfWinNotActive, Microsoft Excel, , WinActivate, Microsoft Excel, 
	WinWaitActive, Microsoft Excel, 
	Sleep,500
	Send,{ENTER}
	WinWait, %now_date% wjda.csv - Excel, 
	IfWinNotActive, %now_date% wjda.csv - Excel, , WinActivate, %now_date% wjda.csv - Excel, 
	WinWaitActive, %now_date% wjda.csv - Excel,
	Sleep,500
	Send, {ALTDOWN}{F4}{ALTUP}
	WinWait, Microsoft Excel, 
	IfWinNotActive, Microsoft Excel, , WinActivate, Microsoft Excel, 
	WinWaitActive, Microsoft Excel, 
	Sleep,500
	Send, {RIGHT}
	Sleep,500
	Send,{ENTER}
	Sleep,1000
	
;——————————————————
;9.生成样单数据
;——————————————————	
	
	;9.1 打开dropbox共享文件夹
	Run,C:\Users\xiaozhou\Dropbox\2017CGSS全部样本 (1)\2017CGSS全部样本
	WinWait, 2017CGSS全部样本, 
	IfWinNotActive, 2017CGSS全部样本, , WinActivate, 2017CGSS全部样本, 
	WinWaitActive, 2017CGSS全部样本, 
	Send,{CTRLDOWN}f{CTRLUP}
	Sleep,500
	Send,{SHIFT}
	Sleep,500

	;9.2 搜索符合条件的样单文件并复制它们
	Send,.xlsx 作者:Yingfeng
	WinWait, .xlsx 作者:Yingfeng - “2017CGSS全部样本”中的搜索结果, 
	IfWinNotActive, .xlsx 作者:Yingfeng - “2017CGSS全部样本”中的搜索结果, , WinActivate, .xlsx 作者:Yingfeng - “2017CGSS全部样本”中的搜索结果, 
	WinWaitActive, .xlsx 作者:Yingfeng - “2017CGSS全部样本”中的搜索结果, 
	Sleep,4000
	Loop 2
	{
		Send,{TAB}
		Sleep,500
	}
	;复制16个文件
	Send,{SHIFTDOWN}{PGDN}{PGDN}{SHIFTUP}
	;Send,{SHIFTDOWN}{PGDN}{PGDN}{PGDN}{SHIFTUP}
	Sleep,4000
	Send,{CTRLDOWN}c{CTRLUP}
	Sleep,1000
	Send,{ALTDOWN}{F4}{ALTUP}
	Sleep,1000

	;9.3 在一个特定文件夹中粘贴样单文件
	Run,D:\CGSS调查问卷统计\样本清单信息
	WinWait, 样本清单信息, 
	IfWinNotActive, 样本清单信息, , WinActivate, 样本清单信息, 
	WinWaitActive, 样本清单信息, 
	Sleep,500
	Send,{CTRLDOWN}v{CTRLUP}
	;WinWait, 已完成 100`%, 
	;IfWinNotActive, 已完成 100`%, , WinActivate, 已完成 100`%, 
	;WinWaitActive, 已完成 100`%, 
	Sleep,20000
	Send,{HOME}
	Sleep,1000

	;9.4 转样单.xlsx为样单.csv
	;Loop 40
	Loop 16
	{
		Send,{ENTER}
		WinWait, ,xlsx, 
		IfWinNotActive, ,xlsx, , WinActivate, ,xlsx, 
		WinWaitActive, ,xlsx,
		Sleep,2000
		
		Send,{CTRLDOWN}{PGUP}{PGUP}{CTRLUP}
		Sleep,1000
		excel := ComObjActive("Excel.Application") 
		Sheet := excel.Worksheets[1]
		If(Sheet.Cells[2,1].Value!="样本编号")
			Send,{CTRLDOWN}{PGDN}{CTRLUP}
		;上六行命令用来确保保存的是正选样本sheet
		Sleep,1000
		Send,{ALT}
		Sleep,500
		Send,f
		Sleep,500
		Send,a
		Sleep,500
		Send,o
		
		WinWait, 另存为, 
		IfWinNotActive, 另存为, , WinActivate, 另存为, 
		WinWaitActive, 另存为, 
		Sleep,500
		Send, {TAB}
		Sleep,1000
		Send,c
		Sleep,1000
		Send,{SHIFT}
		Sleep,1000
		Send,{ENTER}
		WinWait, Microsoft Excel, 
		IfWinNotActive, Microsoft Excel, , WinActivate, Microsoft Excel, 
		WinWaitActive, Microsoft Excel, 
		Sleep,500
		Send,{ENTER}
		WinWait, Microsoft Excel, 
		IfWinNotActive, Microsoft Excel, , WinActivate, Microsoft Excel, 
		WinWaitActive, Microsoft Excel, 
		Sleep,1000
		Send,{ENTER}
		WinWait, ,csv, 
		IfWinNotActive, ,csv, , WinActivate, ,csv, 
		WinWaitActive, ,csv, 
		Sleep,500
		Send, {ALTDOWN}{F4}{ALTUP}
		WinWait, Microsoft Excel, 
		IfWinNotActive, Microsoft Excel, , WinActivate, Microsoft Excel, 
		WinWaitActive, Microsoft Excel, 
		Sleep,500
		Send, {RIGHT}
		Sleep,1000
		Send,{ENTER}
		WinWait, 样本清单信息, 
		IfWinNotActive, 样本清单信息, , WinActivate, 样本清单信息, 
		WinWaitActive, 样本清单信息, 
		Sleep,500
		Send, {DEL}
		Sleep,1000
		Send,{DOWN}
		Sleep,1000
		Send,{UP}
		Sleep,1000
	}

	;9.5 运行copy.bat，将其粘贴到指定文件夹中
	Run,D:\My bat code\copy.bat
	WinWait, 样本清单信息, 
	IfWinNotActive, 样本清单信息, , WinActivate, 样本清单信息, 
	WinWaitActive, 样本清单信息, 
	Sleep,2000
	Run,D:\My bat code\delete.bat
	WinWait, 样本清单信息, 
	IfWinNotActive, 样本清单信息, , WinActivate, 样本清单信息, 
	WinWaitActive, 样本清单信息,
	Sleep,2000
	Send,{ALTDOWN}{F4}{ALTUP}
	Sleep,2000

;——————————————————
;10.运行R语言分析程序
;——————————————————
	
	;10.1 运行R语言程序
	Run,D:\Program Files\R-3.4.1\bin\x64\Rgui.exe
	WinWait, RGui (64-bit), 
	IfWinNotActive, RGui (64-bit), , WinActivate, RGui (64-bit), 
	WinWaitActive, RGui (64-bit),
	Send,{SHIFT}
	Sleep,500
	Run,C:\Windows\System32\notepad.exe
	WinWait, 无标题 - 记事本, 
	IfWinNotActive, 无标题 - 记事本, , WinActivate, 无标题 - 记事本, 
	WinWaitActive, 无标题 - 记事本,	
	Sleep,500
	Send,{SHIFT}
	Sleep,500
	Send,source("D:\\CGSS调查问卷统计\\数据质控测试文档\\CGSS程序0831早.r")
	Sleep,500
	Send,{CTRLDOWN}a{CTRLUP}
	Sleep,500
	Send,{CTRLDOWN}c{CTRLUP}
	Sleep,500
	Send,{ALTDOWN}{F4}{ALTUP}
	WinWait, 记事本, 
	IfWinNotActive, 记事本, , WinActivate, 记事本, 
	WinWaitActive, 记事本,
	Sleep,500
	Send,n
	WinWait, RGui (64-bit), 
	IfWinNotActive, RGui (64-bit), , WinActivate, RGui (64-bit), 
	WinWaitActive, RGui (64-bit),
	Sleep,500
	Send,{CTRLDOWN}v{CTRLUP}
	Sleep,500
	Send,{ENTER}

	;判断R语言程序是否执行完毕的像素块信息
	;0[1 of 3]: 0
	;color[8 of 63]: 0x0000FF
	;ErrorLevel[1 of 3]: 0
	;MouseX[2 of 3]: 54
	;MouseY[3 of 3]: 629
	
	;10.2 判断R语言程序是否执行完毕
	;方法是判断屏幕上(54,629)坐标点上的像素块颜色是否为0x0000FF，如果存在，循环终止，如果不存在，循环继续。当R语言程序即将执行完毕时，该坐标点上像素块为红色，而当R语言程序未执行完毕时，该坐标点上像素块为白色0xFFFFFF。因此当R语言即将程序执行完毕时，循环便会终止，程序将紧接着执行下面的命令。
	;颜色0x0000FF的获取来自于AutoHotkey的一个示例脚本，现置于linshi.ahk中，由鼠标所在位置以及ctrl+alt+Z热键联合触发
	Loop
	{
		PixelSearch, RendX, RendY, 54, 629, 54, 629, 0x0000FF, 0, Fast
		Sleep,1000
		If ErrorLevel
			continue
		else
			break
	}
	Sleep,2000

;——————————————————
;11.上传电核数据，更新电核大表
;——————————————————

	;11.1 用notepad++更改dhyb_to_jm.csv的编码格式
	Rjlj = D:\应用程序\Notepad++\notepad++.exe
	Wjlj = D:\CGSS调查问卷统计\数据质控测试文档\%now_date% dhyb_to_jm.csv
	Run,%Rjlj% "%Wjlj%"
	WinWait, %Wjlj% - Notepad++, 
	IfWinNotActive, %Wjlj% - Notepad++, , WinActivate, %Wjlj% - Notepad++, 
	WinWaitActive, %Wjlj% - Notepad++,
	Sleep,500
	Send,{ALT}
	Sleep,500
	Send,m
	Sleep,500
	Loop 3
	{
		Send,{UP}
		Sleep,500
	}
	Send,{ENTER}
	Sleep,500
	Send,{CTRLDOWN}s{CTRLUP}
	Sleep,1000
	Send,{ALTDOWN}{F4}{ALTUP}
	Sleep,1000
	
	;11.2 修改dhyb_to_me.csv的4列数据的格式，并改变列宽
	;修改列宽
	Run,D:\CGSS调查问卷统计\数据质控测试文档\%now_date% dhyb_to_me.csv
	WinWait, %now_date% dhyb_to_me.csv - Excel, 
	IfWinNotActive, %now_date% dhyb_to_me.csv - Excel, , WinActivate, %now_date% dhyb_to_me.csv - Excel, 
	WinWaitActive, %now_date% dhyb_to_me.csv - Excel,
	Sleep,8000
	MouseClick, left, 72, 251
	Sleep,500
	Send,{APPSKEY}
	Sleep,500
	Send,c
	Sleep,500
	Send,c
	Sleep,500
	Send,{ENTER}
	WinWait, 列宽, 
	IfWinNotActive, 列宽, , WinActivate, 列宽, 
	WinWaitActive, 列宽,
	Sleep,500
	Send,10
	Sleep,500
	Send,{ENTER}
	Sleep,1000
	;修改token列数据格式
	Send,{RIGHT}
	Sleep,1000
	Send,{CTRLDOWN}{SHIFTDOWN}{DOWN}{SHIFTUP}{CTRLUP}
	Sleep,2000
	Send,{APPSKEY}
	Sleep,1000
	Send,f
	Sleep,1000
	WinWait, 设置单元格格式, 
	IfWinNotActive, 设置单元格格式, , WinActivate, 设置单元格格式, 
	WinWaitActive, 设置单元格格式,
	Sleep,2000
	Send,{TAB}
	Sleep,500
	Send,{DOWN}
	Sleep,500
	Send,{TAB}
	Sleep,500
	Send,0
	Sleep,500
	Send,{ENTER}
	Sleep,4000
	WinWait, %now_date% dhyb_to_me.csv - Excel, 
	IfWinNotActive, %now_date% dhyb_to_me.csv - Excel, , WinActivate, %now_date% dhyb_to_me.csv - Excel, 
	WinWaitActive, %now_date% dhyb_to_me.csv - Excel,
	Sleep,500
	Loop 14
	{
		Send,{RIGHT}
		Sleep,200
	}
	;修改S00列数据格式
	Send,{CTRLDOWN}{SHIFTDOWN}{DOWN}{SHIFTUP}{CTRLUP}
	Sleep,2000
	Send,{APPSKEY}
	Sleep,1000
	Send,f
	Sleep,500
	WinWait, 设置单元格格式, 
	IfWinNotActive, 设置单元格格式, , WinActivate, 设置单元格格式, 
	WinWaitActive, 设置单元格格式,
	Sleep,500
	Send,{TAB}
	Sleep,500
	Send,{DOWN}
	Sleep,500
	Send,{TAB}
	Sleep,500
	Send,0
	Sleep,500
	Send,{ENTER}
	Sleep,4000
	WinWait, %now_date% dhyb_to_me.csv - Excel, 
	IfWinNotActive, %now_date% dhyb_to_me.csv - Excel, , WinActivate, %now_date% dhyb_to_me.csv - Excel, 
	WinWaitActive, %now_date% dhyb_to_me.csv - Excel,
	Sleep,500
	Loop 4
	{
		Send,{RIGHT}
		Sleep,200
	}
	;修改Z2列数据格式
	Send,{CTRLDOWN}{SHIFTDOWN}{DOWN}{SHIFTUP}{CTRLUP}
	Sleep,2000
	Send,{APPSKEY}
	Sleep,1000
	Send,f
	Sleep,500
	WinWait, 设置单元格格式, 
	IfWinNotActive, 设置单元格格式, , WinActivate, 设置单元格格式, 
	WinWaitActive, 设置单元格格式,
	Sleep,500
	Send,{TAB}
	Sleep,500
	Send,{DOWN}
	Sleep,500
	Send,{TAB}
	Sleep,500
	Send,0
	Sleep,500
	Send,{ENTER}
	Sleep,4000
	WinWait, %now_date% dhyb_to_me.csv - Excel, 
	IfWinNotActive, %now_date% dhyb_to_me.csv - Excel, , WinActivate, %now_date% dhyb_to_me.csv - Excel, 
	WinWaitActive, %now_date% dhyb_to_me.csv - Excel,
	Sleep,500
	Loop 61
	{
		Send,{RIGHT}
		Sleep,200
	}
	;修改attribute_17列数据格式
	Sleep,500
	Send,{CTRLDOWN}{SHIFTDOWN}{DOWN}{DOWN}{DOWN}{SHIFTUP}{CTRLUP}
	Sleep,2000
	Send,{APPSKEY}
	Sleep,1000
	Send,f
	Sleep,500
	WinWait, 设置单元格格式, 
	IfWinNotActive, 设置单元格格式, , WinActivate, 设置单元格格式, 
	WinWaitActive, 设置单元格格式,
	Sleep,500
	Send,{TAB}
	Sleep,500
	Send,{DOWN}
	Sleep,500
	Send,{TAB}
	Sleep,500
	Send,0
	Sleep,500
	Send,{ENTER}
	Sleep,4000
	WinWait, %now_date% dhyb_to_me.csv - Excel, 
	IfWinNotActive, %now_date% dhyb_to_me.csv - Excel, , WinActivate, %now_date% dhyb_to_me.csv - Excel, 
	WinWaitActive, %now_date% dhyb_to_me.csv - Excel,
	Sleep,500
	;复制电核大表所需数据
	Send,{LEFT}
	Sleep,500
	Send,{HOME}
	Sleep,500
	Send,{DOWN}
	Sleep,500
	Send,{CTRLDOWN}{SHIFTDOWN}{DOWN}{SHIFTUP}{CTRLUP}
	Sleep,500
	Send,{SHIFTDOWN}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{SHIFTUP}
	Sleep,1000
	Send,{CTRLDOWN}c{CTRLUP}
	Sleep,3000
	
	;11.3 上传电核数据到limesurvey系统
	Run,http://101.200.178.132/limesurvey_2017/index.php/admin
	WinWait, LimeSurvey_2017 - Google Chrome, 
	IfWinNotActive, LimeSurvey_2017 - Google Chrome, , 	WinActivate, LimeSurvey_2017 - Google Chrome, 
	WinWaitActive, LimeSurvey_2017 - Google Chrome, 
	Sleep,1000
	Send, jiamin
	Sleep,500
	Send,{SHIFT}
	Sleep,100
	Send,{TAB}
	Sleep,100
	Send,123456
	Sleep,100
	Send,{ENTER}
	Sleep,100
	Loop
	{
		PixelSearch, LogoX, LogoY, 501, 216, 553, 271, 0x0BEDCB, 2, Fast
		Sleep,1000
		If ErrorLevel
			continue
		else
			break	
	}
	Sleep,1000
	Run,http://101.200.178.132/limesurvey_2017/index.php/admin/tokens/sa/import/surveyid/637954
	WinWait, LimeSurvey_2017 - Google Chrome, 
	IfWinNotActive, LimeSurvey_2017 - Google Chrome, , 	WinActivate, LimeSurvey_2017 - Google Chrome, 
	WinWaitActive, LimeSurvey_2017 - Google Chrome, 
	Sleep,8000
	Loop 3
	{
		Send,{DOWN}
		Sleep,1000
	}
	Sleep,1000
	Send,f
	Sleep,500
	Send,e
	WinWait, 打开, 
	IfWinNotActive, 打开, , 	WinActivate, 打开, 
	WinWaitActive, 打开,	
	Sleep,1000
	Send,{CTRLDOWN}f{CTRLUP}
	Sleep,1000
	Send,{SHIFTDOWN}{TAB}{SHIFTUP}
	Sleep,1000
	Send,{ENTER}
	Sleep,1000
	Send,D:\CGSS{SHIFTDOWN}{SHIFTUP}
	Sleep,1000
	Send,调查问卷统计\
	Sleep,1000
	Send,数据质控测试文档
	Sleep,1000
	Send,{ENTER}
	Sleep,1000
	Send,{CTRLDOWN}f{CTRLUP}
	Sleep,1000
	Send,AND “%now_date% dhyb_to_jm.csv”
	Sleep,5000
	Send,{ENTER}
	Sleep,4000
	Loop 3
	{
		Send,{TAB}
		Sleep,1000
	}
	Send,{DOWN}
	Sleep,1000
	Send,{HOME}
	Sleep,1000
	Loop 3
	{
		Send,{TAB}
		Sleep,1000
	}
	Send,{ENTER}
	Sleep,2000	
	Send,f
	Sleep,500
	Send,k
	Sleep,500
	Send,f
	Sleep,500
	Send,l
	Sleep,500
	Send,f
	Sleep,500
	Send,m
	Sleep,500
	Send,f
	Sleep,500
	Send,ss
	Sleep,500
	Loop 4
	{
		Send,{DOWN}
		Sleep,500
	}
	Send,{TAB}
	Sleep,500
	Send,f
	Sleep,500
	Send,w
	Sleep,10000
	
	;11.4 上传电核数据到电核大表
	Run,https://docs.google.com/spreadsheets/d/1tfRQ_n3-4bTEemeqWZlhvLNlV-eAjKCv9dpNVs7ETSc/edit#gid=0
	WinWait, CGSS2017电话核查样本更新 - Google 表格 - Google Chrome, 
	IfWinNotActive, CGSS2017电话核查样本更新 - Google 表格 - Google Chrome, , 	WinActivate, CGSS2017电话核查样本更新 - Google 表格 - Google Chrome, 
	WinWaitActive, CGSS2017电话核查样本更新 - Google 表格 - Google Chrome,
	Sleep,30000
	Send,{CTRLDOWN}{DOWN}{CTRLUP}
	Sleep,15000
	Send,{DOWN}
	Sleep,5000
	Send,{CTRLDOWN}{SHIFTDOWN}v{SHIFTUP}{CTRLUP}
	Sleep,20000

;——————————————————
;12.更新电核大表村居进度统计
;——————————————————

	;12.1 切换到电核大表的第4个sheet:村居进度统计表
	Loop 3
	{
		Send,{CTRLDOWN}{SHIFTDOWN}{PGDN}{SHIFTUP}{CTRLUP}
		Sleep,2000	
	}
	
	;12.2 标识更新日期
	Send,{CTRLDOWN}{HOME}{CTRLUP}
	Sleep,1000
	Send,{UP}
	Loop 3
	{
		Send,{LEFT}
		Sleep,1000
	}
	Send,省 %now_date%
	Sleep,2000
	Send,{ENTER}
	Sleep,1000
	
	;12.3 定位光标到数据粘贴位置
	Send,{CTRLDOWN}{HOME}{CTRLUP}
	Sleep,2000
	
	;12.4 打开村居进度统计本地表格：系统日期 cjtj.csv
	Run,D:\CGSS调查问卷统计\数据质控测试文档\%now_date% cjtj.csv
	WinWait, %now_date% cjtj.csv - Excel, 
	IfWinNotActive, %now_date% cjtj.csv - Excel, , 	WinActivate, %now_date% cjtj.csv - Excel, 
	WinWaitActive, %now_date% cjtj.csv - Excel,
	Sleep,2000
	Send,{LWINDOWN}{UP}{UP}{LWINUP}
	Sleep,3000
	
	;12.5 定位光标到数据复制位置
	Loop 7
	{
		Send,{RIGHT}
		Sleep,500
	}
	Send,{DOWN}
	Sleep,500
	
	;12.6 选定复制区域并复制区域内数据，然后最小化csv文件
	Send,{SHIFTDOWN}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{SHIFTUP}
	Sleep,2000
	Send,{SHIFTDOWN}{PGDN}{PGDN}{PGDN}{PGDN}{PGDN}{PGDN}{PGDN}{PGDN}{PGDN}{PGDN}{PGDN}{PGDN}{PGDN}{PGDN}{PGDN}{PGDN}{PGDN}{PGDN}{PGDN}{PGDN}{PGDN}{PGDN}{UP}{UP}{UP}{UP}{UP}{UP}{UP}{UP}{UP}{SHIFTUP}
	Sleep,2000
	Send,{CTRLDOWN}c{CTRLUP}
	Sleep,2000
	Send,{LWINDOWN}{DOWN}{LWINUP}
	Sleep,2000
	Send,{LWINDOWN}{DOWN}{LWINUP}
	Sleep,2000
	
	;12.7 更新电核大表村居进度统计
	Send,{CTRLDOWN}{SHIFTDOWN}v{SHIFTUP}{CTRLUP}
	Sleep,8000
	
;——————————————————
;13.更新电核大表高缺失率问卷
;——————————————————	

	;13.1 切换到电核大表的第5个sheet:高缺失率问卷
	Send,{CTRLDOWN}{SHIFTDOWN}{PGDN}{SHIFTUP}{CTRLUP}
	Sleep,8000
	
	;13.2 定位光标到数据粘贴位置
	Send,{CTRLDOWN}{DOWN}{CTRLUP}
	Sleep,8000
	Send,{DOWN}
	Sleep,2000
	
	;13.3 打开高缺失率问卷本地表格：系统日期 qstjwsc.csv
	Run,D:\CGSS调查问卷统计\数据质控测试文档\%now_date% qstjwsc.csv
	WinWait, %now_date% qstjwsc.csv - Excel, 
	IfWinNotActive, %now_date% qstjwsc.csv - Excel, , 	WinActivate, %now_date% qstjwsc.csv - Excel, 
	WinWaitActive, %now_date% qstjwsc.csv - Excel,
	Sleep,2000
	Send,{LWINDOWN}{UP}{UP}{LWINUP}
	Sleep,3000
	
	;13.4 更改token列数据格式
	Loop 5
	{
		Send,{RIGHT}
		Sleep,1000
	}
	Send,{CTRLDOWN}{SHIFTDOWN}{DOWN}{SHIFTUP}{CTRLUP}
	Sleep,2000
	Send,{APPSKEY}
	Sleep,1000
	Send,f
	Sleep,1000
	Send,f
	Sleep,1000
	Send,{ENTER}
	WinWait, 设置单元格格式, 
	IfWinNotActive, 设置单元格格式, , WinActivate, 设置单元格格式, 
	WinWaitActive, 设置单元格格式,
	Sleep,2000
	Send,{TAB}
	Sleep,500
	Send,{DOWN}
	Sleep,500
	Send,{TAB}
	Sleep,500
	Send,0
	Sleep,500
	Send,{ENTER}
	Sleep,4000
	WinWait, %now_date% qstjwsc.csv - Excel, 
	IfWinNotActive, %now_date% qstjwsc.csv - Excel, , WinActivate, %now_date% qstjwsc.csv - Excel, 
	WinWaitActive, %now_date% qstjwsc.csv - Excel,
	Sleep,500
	Loop 5
	{
		Send,{LEFT}
		Sleep,500
	}
	
	;13.5 选定复制区域并复制区域内数据,然后最小化csv文件
	Send,{DOWN}
	Sleep,2000
	Send,{CTRLDOWN}{SHIFTDOWN}{DOWN}{SHIFTUP}{CTRLUP}
	Sleep,2000
	Send,{SHIFTDOWN}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{SHIFTUP}
	Sleep,2000
	Send,{CTRLDOWN}c{CTRLUP}
	Sleep,2000
	Send,{LWINDOWN}{DOWN}{LWINUP}
	Sleep,2000
	Send,{LWINDOWN}{DOWN}{LWINUP}
	Sleep,2000
	
	;13.6 更新高缺失率问卷
	Send,{CTRLDOWN}{SHIFTDOWN}v{SHIFTUP}{CTRLUP}
	Sleep,8000

;——————————————————
;14.更新电核大表低电话回收率访员
;——————————————————	
	
	;14.1 切换到电核大表的第6个sheet:低电话回收率问卷
	Send,{CTRLDOWN}{SHIFTDOWN}{PGDN}{SHIFTUP}{CTRLUP}
	Sleep,8000
	
	;14.2 定位光标到数据粘贴位置
	Send,{CTRLDOWN}{DOWN}{CTRLUP}
	Sleep,8000
	Send,{DOWN}
	Sleep,4000
	
	;14.3 打开低电话回收率访员本地表格：系统日期 低电话回收率未上传访员名单.csv
	Run,D:\CGSS调查问卷统计\数据质控测试文档\%now_date% 低电话回收率未上传访员名单.csv
	WinWait, %now_date% 低电话回收率未上传访员名单.csv - Excel, 
	IfWinNotActive, %now_date% 低电话回收率未上传访员名单.csv - Excel, , 	WinActivate, %now_date% 低电话回收率未上传访员名单.csv - Excel, 
	WinWaitActive, %now_date% 低电话回收率未上传访员名单.csv - Excel,
	Sleep,2000
	Send,{LWINDOWN}{UP}{UP}{LWINUP}
	Sleep,3000
	
	;14.4 选定复制区域并复制区域内数据,然后最小化csv文件
	Send,{DOWN}
	Sleep,2000
	Send,{CTRLDOWN}{SHIFTDOWN}{DOWN}{SHIFTUP}{CTRLUP}
	Sleep,2000
	Send,{CTRLDOWN}{SHIFTDOWN}{RIGHT}{SHIFTUP}{CTRLUP}
	Sleep,2000
	Send,{CTRLDOWN}c{CTRLUP}
	Sleep,2000
	Send,{LWINDOWN}{DOWN}{LWINUP}
	Sleep,2000
	Send,{LWINDOWN}{DOWN}{LWINUP}
	Sleep,2000
	
	;14.5 更新高缺失率问卷
	Send,{CTRLDOWN}{SHIFTDOWN}v{SHIFTUP}{CTRLUP}
	Sleep,8000

;——————————————————
;15.给昕烨、严欣发邮件
;——————————————————

	;15.1 登录QQ邮箱
	Run,https://mail.qq.com
	WinWait, 登录QQ邮箱 - Google Chrome, 
	IfWinNotActive, 登录QQ邮箱 - Google Chrome, , 	WinActivate, 登录QQ邮箱 - Google Chrome, 
	WinWaitActive, 登录QQ邮箱 - Google Chrome,
	Sleep,500
	Send,Xiaozhou13171317
	Sleep,1000
	Send,{ENTER}
	WinWait, QQ邮箱, 
	IfWinNotActive, QQ邮箱, , 	WinActivate, QQ邮箱, 
	WinWaitActive, QQ邮箱,
	Sleep,4000
	
;——————————————————
;16.给雨薇发邮件
;——————————————————

	;16.1 进入写信页面,填写收件人、邮件主题和邮件正文
	MouseClick, left, 97, 210
	Sleep,2000
	Send,{SHIFTDOWN}{SHIFTUP}
	Sleep,2000
	Send,snailallez@163.com;
	;Send,737130291@qq.com
	Sleep,500
	Send,{ENTER}
	Sleep,2000
	Send,{TAB}
	Sleep,500
	Send,{TAB}
	Sleep,500
	Send,%now_date% 调查进度与无联系记录样本
	Sleep,1000
	Send,{TAB}
	Run,D:\CGSS调查问卷统计\数据质控测试文档\%now_date% toyw.txt
	WinWait, %now_date% toyw.txt - 记事本, 
	IfWinNotActive, %now_date% toyw.txt - 记事本, , 	WinActivate, %now_date% toyw.txt - 记事本, 
	WinWaitActive, %now_date% toyw.txt - 记事本,
	Sleep,1000
	Send,{CTRLDOWN}a{CTRLUP}
	Sleep,1000
	Send,{CTRLDOWN}c{CTRLUP}
	Sleep,1000
	Send,{ALTDOWN}{F4}{ALTUP}
	WinWait, QQ邮箱, 
	IfWinNotActive, QQ邮箱, , 	WinActivate, QQ邮箱, 
	WinWaitActive, QQ邮箱,
	Sleep,1000
	Send,{CTRLDOWN}v{CTRLUP}
	Sleep,1000
	MouseClick, left, 550, 718
	Sleep, 2000
	
	;16.2 添加附件
	MouseClick, left, 385, 227
	WinWait, 打开, 
	IfWinNotActive, 打开, , 	WinActivate, 打开, 
	WinWaitActive, 打开,
	Sleep,1000
	Send,{CTRLDOWN}f{CTRLUP}
	Sleep,1000
	Send,{SHIFTDOWN}{TAB}{SHIFTUP}
	Sleep,500
	Send,{ENTER}
	Sleep,500
	Send,D:\CGSS{SHIFTDOWN}{SHIFTUP}
	Sleep,500
	Send,调查问卷统计\
	Sleep,500
	Send,数据质控测试文档
	Sleep,500
	Send,{ENTER}
	Sleep,500
	Send,{CTRLDOWN}f{CTRLUP}
	Sleep,500
	Send,NOT “北京地区” AND “%now_date% 调查进度” OR “%now_date% 无联系”
	Sleep,4000
	Loop 3
	{
		Send,{TAB}
		Sleep,500
	}
	Send,{CTRLDOWN}a{CTRLUP}
	Sleep,500
	Loop 3
	{
		Send,{TAB}
		Sleep,500
	}
	Send,{ENTER}
	Sleep,25000
	
	;16.3 发送邮件
	MouseClick, left, 294, 712
	Sleep, 1000
	Loop 6
	{
		Send,{DOWN}
		Sleep,500
	}
	MouseClick, left, 294, 712
	Sleep, 8000
	
	;16.4 判断邮件是否发送完毕
	Loop
	{
		PixelSearch, MailendX, MailendY, 292,252, 292,252, 0xBAF2D7, 0, Fast
		Sleep,1000
		If ErrorLevel
			continue
		else
			break
	}

	Sleep,5000

;——————————————————
;17.给福鑫、贾敏发邮件
;——————————————————	
	
	;17.1 进入写信页面,填写收件人、邮件主题和邮件正文
	MouseClick, left, 97, 210
	Sleep,2000
	Send,18170146171@163.com;
	Sleep,2000
	Send,{ENTER}
	Send,1131738451@qq.com;
	Sleep,2000
	Send,{ENTER}
	;Send,737130291@qq.com
	Sleep,2000
	Send,{TAB}
	Sleep,500
	Send,{TAB}
	Sleep,500
	Send,%now_date% 调查进度与原始数据
	Sleep,1000
	Send,{TAB}
	Run,D:\CGSS调查问卷统计\数据质控测试文档\tofxjm.txt
	WinWait, tofxjm.txt - 记事本, 
	IfWinNotActive, tofxjm.txt - 记事本, , 	WinActivate, tofxjm.txt - 记事本, 
	WinWaitActive, tofxjm.txt - 记事本,
	Sleep,1000
	Send,{CTRLDOWN}a{CTRLUP}
	Sleep,1000
	Send,{CTRLDOWN}c{CTRLUP}
	Sleep,1000
	Send,{ALTDOWN}{F4}{ALTUP}
	WinWait, QQ邮箱, 
	IfWinNotActive, QQ邮箱, , 	WinActivate, QQ邮箱, 
	WinWaitActive, QQ邮箱,
	Sleep,1000
	Send,{CTRLDOWN}v{CTRLUP}
	Sleep,1000
	Send,{TAB}
	Sleep,1000
	MouseClick, left, 550, 718
	Sleep, 2000
	
	;17.2 添加附件
	MouseClick, left, 385, 227
	WinWait, 打开, 
	IfWinNotActive, 打开, , 	WinActivate, 打开, 
	WinWaitActive, 打开,
	Sleep,1000
	Send,{CTRLDOWN}f{CTRLUP}
	Sleep,1000
	Send,{SHIFTDOWN}{TAB}{SHIFTUP}
	Sleep,500
	Send,{ENTER}
	Sleep,500
	Send,D:\CGSS{SHIFTDOWN}{SHIFTUP}
	Sleep,500
	Send,调查问卷统计\
	Sleep,500
	Send,数据质控测试文档
	Sleep,500
	Send,{ENTER}
	Sleep,500
	Send,{CTRLDOWN}f{CTRLUP}
	Sleep,500
	Send,NOT “北京地区” AND “%now_date% 调查进度” OR “%now_date% 无电话已完成” OR “%now_date% 原始数据”
	Sleep,4000
	Loop 3
	{
		Send,{TAB}
		Sleep,500
	}
	Send,{CTRLDOWN}a{CTRLUP}
	Sleep,500
	Loop 3
	{
		Send,{TAB}
		Sleep,500
	}
	Send,{ENTER}
	Sleep,8000
	MouseClick,left,711,502
	Sleep,25000
	
	;17.3 发送邮件
	MouseClick, left, 294, 712
	Sleep, 1000
	Loop 6
	{
		Send,{DOWN}
		Sleep,500
	}
	MouseClick, left, 294, 712
	Sleep, 8000

;——————————————————
;18.更新google进度大表
;——————————————————	
	Run,D:\CGSS调查问卷统计\数据质控测试文档\%now_date% 调查进度.xlsx
	WinWait, %now_date% 调查进度.xlsx - Excel, 
	IfWinNotActive, %now_date% 调查进度.xlsx - Excel, , 	WinActivate, %now_date% 调查进度.xlsx - Excel, 
	WinWaitActive, %now_date% 调查进度.xlsx - Excel,
	Sleep,4000
	Send,{LWINDOWN}{UP}{UP}{LWINUP}
	Sleep,2000
	Loop 6
	{
		Send,{RIGHT}
		Sleep,2000
	}
	Send,{DOWN}
	Sleep,2000
	Send,{SHIFTDOWN}{PGDN}{DOWN}{RIGHT}{SHIFTUP}
	Sleep,2000
	Send,{APPSKEY}
	Sleep,2000
	Send,f
	WinWait, 设置单元格格式, 
	IfWinNotActive, 设置单元格格式, , WinActivate, 设置单元格格式, 
	WinWaitActive, 设置单元格格式,
	Sleep,2000
	Send,{TAB}
	Sleep,2000
	Loop 6
	{
		Send,{DOWN}
		Sleep,2000
	}
	Send,{TAB}
	Sleep,2000
	Send,{ENTER}
	Sleep,2000
	Loop 5
	{
		Send,{LEFT}
		Sleep,2000
	}
	Send,{SHIFTDOWN}{PGDN}{DOWN}{SHIFTUP}
	Sleep,2000
	Send,{SHIFTDOWN}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{SHIFTUP}
	Sleep,2000
	Send,{CTRLDOWN}c{CTRLUP}
	Sleep,2000
	Run,https://docs.google.com/spreadsheets/d/1-pBKYvO8Fd32Rmk76ghDrIcKPC4R5renbWGtE5-DYLg/edit#gid=0
	WinWait, CGSS2017 - Google 表格 - Google Chrome, 
	IfWinNotActive, CGSS2017 - Google 表格 - Google Chrome, , 	WinActivate, CGSS2017 - Google 表格 - Google Chrome, 
	WinWaitActive, CGSS2017 - Google 表格 - Google Chrome,
	Sleep,30000
	Send,{RIGHT}
	Sleep,2000
	Send,{CTRLDOWN}{HOME}{CTRLUP}
	Sleep,5000
	Loop 12
	{
		Send,{RIGHT}
		Sleep,2000
	}
	Send,{CTRLDOWN}{SHIFTDOWN}v{SHIFTUP}{CTRLUP}
	Sleep,5000
	
;——————————————————
;19.休眠，关闭硬盘，为内存通电
;——————————————————
	
	DllCall("PowrProf\SetSuspendState", "int", 1, "int", 1, "int", 0)

Return
