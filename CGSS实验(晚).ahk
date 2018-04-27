#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

;——————————————————————
;——————————————————————
;**自定义函数**
;——————————————————————
;——————————————————————

;——————————————————————
;1.等待窗口函数
;——————————————————————

;wn为窗口完整名称, wt为窗口完整名称的子字符串
Ddck(wn,ti=0.1,wt="")
{
	WinWait, %wn%, %wt%,
	IfWinNotActive, %wn%, %wt%, WinActivate, %wn%, %wt%, 
	WinWaitActive, %wn%, %wt%,
	ti := ti*1000
	Sleep, ti
}

;——————————————————————
;2.输入文字函数
;——————————————————————

;word为需要输入的文字，wait为ctrl+v后的睡眠时间
Srwz(word,ti=0.2)
{
	Clipboard = %word%
	Sm()
	Send,^v
	ti := ti*1000
	Sleep, ti
}

;——————————————————————
;3.睡眠函数
;——————————————————————

;函数旨在将Sleep函数的时间参数单位由ms修改为s
Sm(ti=0.1)
{
	ti := ti*1000
	Sleep, ti
}

;——————————————————————
;4.发送按键、子串函数
;——————————————————————

;函数旨在为所有Send命令添加一行Sleep命令，避免手工多次输入Sleep的麻烦
;默认休息时间为0.1秒，休息时间由该函数第2个参数确定
Fs(na,ti=0.1)
{
	Send,%na%
	ti := ti*1000
	Sleep, ti
}

;——————————————————————
;5.寻找颜色函数
;——————————————————————

Sc(a,b,c,d,color,var,ti=0.5)
{
	Loop
	{
		PixelSearch, LogoX, LogoY, a, b, c, d, %color%, var, Fast
		Sleep,100
		If ErrorLevel
			continue
		else
			break	
	}
	ti := ti*1000
	Sleep, ti
}

;——————————————————————
;6.左键单击函数
;——————————————————————

Zjdj(a,b,ti=0.5)
{
	MouseClick,left,a,b
	ti := ti*1000
	Sleep, ti
}



;——————————————————————
;——————————————————————
;**备份CGSS数据命令**
;——————————————————————
;——————————————————————

:*:bfcgsssjw::

;——————————————————————
;1.获取系统当前日期，日期格式为yyyy-MM-dd和yyyyMMdd
;——————————————————————

	FormatTime, now_date, %A_Now%, yyyy-MM-dd
	Sm()
	FormatTime, now_date_wg, %A_Now%, yyyyMMdd
	Sm()

;——————————————————————
;2.在CGSSdata目录下新建以系统日期命名的文件夹
;——————————————————————

	Run, cmd.exe
	Ddck("C:\WINDOWS\SYSTEM32\cmd.exe")
	Srwz("md d:\CGSSdata\")
	Send,%now_date_wg%2
	Sm()
	Fs("{ENTER}")
	Fs("!{F4}")
	
;——————————————————————
;3.登录与下载
;——————————————————————

	;3.1 打开limesurvey登录页面并登录limesurvey系统
	Run,http://101.200.178.132/limesurvey_2017/index.php/admin
	Ddck("LimeSurvey_2017 - Google Chrome")
	Srwz("longzhengfan")
	Fs("{TAB}")
	Fs("123456")
	Fs("{ENTER}")
	
	;3.2 判断是否登录了limesurvey系统
	Sc(501,216,553,271,"0x0BEDCB",2)
	
	;3.3 下载电核大表
	Run,https://docs.google.com/spreadsheets/d/1tfRQ_n3-4bTEemeqWZlhvLNlV-eAjKCv9dpNVs7ETSc/edit?usp=drive_web
	Sc(1259,126,1328,152,"0xF58743",2)
	Zjdj(80,144)
	Loop,7
	{
		Fs("{UP}",0.3)
	}
	Fs("{RIGHT}")
	Fs("{ENTER}")
	
	Ddck("另存为",1)
	Fs("^f",0.2)
	Fs("+{TAB}")
	Fs("{ENTER}")
	Srwz("D:\CGSSdata\")
	Send,%now_date_wg%2
	Sm(0.2)
	Fs("{ENTER}")
	Loop,3
	{
		Fs("{TAB}")
	}
	Fs("{ENTER}")


	;3.4 下载问卷数据
	Run,http://101.200.178.132/limesurvey_2017/index.php/admin/export/sa/exportresults/surveyid/963159
	Ddck("LimeSurvey_2017 - Google Chrome",3)
	
	Fs("f",0.5)
	Fs("sk",0.5)
	Zjdj(585,460)
	Loop,19
	{
		Fs("{DOWN}")
	}
	Fs("f",0.3)
	Fs("e",0.3)
	Fs("f",0.3)
	Fs("sk",0.3)
	Fs("f",0.3)
	Fs("sl",0.3)
	Fs("f",0.3)
	Fs("p",0.3)
	Fs("f",0.3)
	Fs("sw",0.3)
	Fs("^a",0.3)
	Fs("{TAB}",0.3)
	Fs("f",0.3)
	Fs("a",0.3)
	Ddck("另存为",1)
	Fs("{ENTER}",0.3)

;——————————————————————
;4.处理电核大表
;——————————————————————	
	Sm(4)
	;4.1 打开 D:\CGSSdata\系统日期\ 目录
	Run,D:\CGSSdata\%now_date_wg%2\
	Ddck("",0.5,"2017")
	
	;4.2 定位光标到文件：.\CGSS2017电话核查样本更新.xlsx
	Fs("^f",0.5)
	Srwz("电")
	Ddck("",1,"搜索结果")
	Fs("{TAB}")
	Fs("{TAB}")
	Fs("{DOWN}")
	
	;4.3 修改文件属性，取消锁定
	Fs("!{ENTER}")
	Ddck("CGSS2017电话核查样本更新.xlsx 属性")
	Fs("k")
	Fs("{ENTER}")
	Fs("!{F4}")

	;4.4 处理电核大表
	Run,D:\CGSSdata\%now_date_wg%2\CGSS2017电话核查样本更新.xlsx
	Ddck("CGSS2017电话核查样本更新.xlsx - Excel",0.5)
	;ctrl+t用来执行处理电核大表的excel visual basic程序,该快捷键由用户自定义
	Fs("^t")
	Ddck("",0.5,"csv")
	Fs("!{F4}",1)
	
;——————————————————————
;5.问卷下载结束后，退出limesurvey系统
;——————————————————————		

	;5.1 检测问卷数据是否下载完毕
	DetectHiddenWindows, On
	WinActivate, LimeSurvey_2017 - Google Chrome, 
	Sm(0.5)
	MouseMove, 152, 721
	Sc(47,719,47,719,"0xDFDFDF",0)
	
	;5.2 退出limesurvey系统
	Zjdj(1058,535)
	Loop,22
	{
		Fs("{UP}")
	}
	Zjdj(1236,123) 
	Loop,2
	{
		Fs("{DOWN}")
	}
	Fs("{ENTER}",1)
	Fs("!{F4}",1)

;——————————————————
;6.处理问卷数据
;——————————————————

	;6.1 打开 D:\CGSSdata\系统日期\
	Run,D:\CGSSdata\%now_date_wg%2\
	Ddck("",0.5,"2017")
	
	;6.2 定位光标到文件：.\results-survey963159.xlsx
	Fs("^f",1)
	Srwz("results-survey963159.xlsx")
	Ddck("",1,"搜索结果")
	Fs("{TAB}")
	Fs("{TAB}")
	Fs("{DOWN}")
	
	;6.3 修改文件属性，取消锁定
	Fs("!{ENTER}")
	Ddck("results-survey963159.xlsx 属性")
	Fs("k")
	Fs("{ENTER}",1)
	Fs("!{F4}")
	
	;6.4 打开文件
	Run,D:\CGSSdata\%now_date_wg%2\results-survey963159.xlsx
	Ddck("results-survey963159.xlsx - Excel",0.5)
	;ctrl+u用来执行处理电核大表的excel visual basic程序,该快捷键由用户自定义
	Fs("^u")
	Ddck("",2,"csv")
	Fs("!{F4}",1)

;——————————————————
;7.运行R语言分析程序
;——————————————————
	
	Run,D:\Program Files\R-3.4.1\bin\x64\Rgui.exe
	Ddck("RGui (64-bit)",2)
	Srwz("source('D:\\CGSS调查问卷统计\\数据质控测试文档\\CGSS程序0831早.r')")
	Fs("{ENTER}")
	Sc(54,629,54,629,"0x0000FF",0)

;——————————————————
;8.上传电核数据到limesurvey系统
;——————————————————

	;8.1 用notepad++更改dhyb_to_jm.csv的编码格式
	Rjlj = D:\应用程序\Notepad++\notepad++.exe
	Wjlj = D:\CGSS调查问卷统计\数据质控测试文档\%now_date% dhyb_to_jm.csv
	Run,%Rjlj% "%Wjlj%"
	WinWait, D:\CGSS调查问卷统计\数据质控测试文档\%now_date% dhyb_to_jm.csv - Notepad++, 
	IfWinNotActive, D:\CGSS调查问卷统计\数据质控测试文档\%now_date% dhyb_to_jm.csv - Notepad++,  WinActivate, D:\CGSS调查问卷统计\数据质控测试文档\%now_date% dhyb_to_jm.csv - Notepad++,  
	WinWaitActive, D:\CGSS调查问卷统计\数据质控测试文档\%now_date% dhyb_to_jm.csv - Notepad++, 
	Fs("{ALT}",0.5)
	Fs("{m}",0.5)
	Loop,3
	{
		Fs("{UP}",0.2)
	}
	Fs("{ENTER}")
	Fs("^s",1)
	;Fs("!{F4}")
	
	;8.2 登录limesurvey系统
	Run,http://101.200.178.132/limesurvey_2017/index.php/admin
	Ddck("LimeSurvey_2017 - Google Chrome")
	Srwz("jiamin")
	Fs("{TAB}")
	Fs("123456")
	Fs("{ENTER}")
	Sc(501,216,553,271,"0x0BEDCB",2)

	;8.2 上传电核数据
	Run,http://101.200.178.132/limesurvey_2017/index.php/admin/tokens/sa/import/surveyid/637954
	Ddck("LimeSurvey_2017 - Google Chrome",5)
	Zjdj(1090,405,0.5)
	Loop 3
	{
		Fs("{DOWN}",1)
	}
	Fs("f",0.5)
	Fs("e")
	
	Ddck("打开",1)
	Fs("^f",1)
	Fs("+{TAB}",0.5)
	Fs("{ENTER}",0.5)
	Srwz("D:\CGSS调查问卷统计\数据质控测试文档")
	Fs("{ENTER}",0.5)
	Fs("^f",0.5)
	Fs("{SHIFT}",0.5)
	Send,AND “%now_date% dhyb_to_jm.csv”
	Sm(1)
	Fs("{ENTER}")
	Sm(4)
	Loop 3
	{
		Fs("{TAB}",0.5)
	}
	Fs("{DOWN}",0.5)
	Fs("{HOME}",0.5)
	Loop,3
	{
		Fs("{TAB}",0.5)
	}
	Fs("{ENTER}",1)
	
	Fs("f",0.3)
	Fs("k",0.3)
	Fs("f",0.3)
	Fs("l",0.3)
	Fs("f",0.3)
	Fs("m",0.3)
	Fs("f",0.3)
	Fs("ss",0.3)
	Loop 4
	{
		Fs("{DOWN}",0.3)
	}
	Fs("{TAB}",0.3)
	Fs("f",0.3)
	Fs("W",0.3)

;——————————————————
;9.上传数据到电核大表
;——————————————————	
	
	;9.1 处理电核样本
	Run,D:\CGSS调查问卷统计\数据质控测试文档\%now_date% dhyb_to_me.csv
	Ddck("",1,"2017")
	Fs("^y",2)
	Fs("!{F4}")
	Ddck("Microsoft Excel")
	Fs("{RIGHT}")
	Fs("{ENTER}")
	Ddck("Microsoft Excel")
	Fs("{ENTER}")
	
	;9.2 上传电核数据到电核大表
	Run,https://docs.google.com/spreadsheets/d/1tfRQ_n3-4bTEemeqWZlhvLNlV-eAjKCv9dpNVs7ETSc/edit#gid=0
	Sc(1259,126,1328,152,"0xF58743",2)
	Zjdj(106,315)
	Fs("^{DOWN}")
	Sc(154,307,154,307,"0xDADADA",2)
	Fs("{DOWN}",2)
	Fs("^+v")
	Sc(243,686,243,686,"0xFFF3EC",0)

	;9.3 上传村居进度到电核大表
	Fs("^+{PGDN}")
	Sc(145,264,145,264,"0xA8D7B6",0)
	Fs("^+{PGDN}")
	Sc(145,264,145,264,"0xC9C4A2",0)
	Fs("^+{PGDN}")
	Sc(145,264,145,264,"0xBDA6D5",0)
	Fs("^{HOME}")
	Fs("{UP}")
	Loop 3
	{
		Fs("{LEFT}")
	}
	Send, 省 %now_date%
	Sm(0.5)
	Fs("{ENTER}")
	Fs("^{HOME}")
	
	Run,D:\CGSS调查问卷统计\数据质控测试文档\%now_date% cjtj.csv
	Ddck("",1,"2017")
	Fs("^e",2)
	Fs("!{F4}")
	Ddck("Microsoft Excel")
	Fs("{ENTER}")
	DetectHiddenWindows, On
	WinActivate, CGSS2017电话核查样本更新 - Google 表格 - Google Chrome, 
	Sm(1)
	Fs("^+v")
	Sc(739,367,739,367,"0xFFF3EC",0) 
	
	;9.4 上传高缺失率问卷到电核大表
	Fs("^+{PGDN}")
	Sc(145,264,145,264,"0x99E5FF",0)
	Fs("^{DOWN}")
	Sc(11,674,11,674,"0xDDDDDD",0)
	Fs("{DOWN}")
	FileGetSize, OutputVar, D:\CGSS调查问卷统计\数据质控测试文档\%now_date% qstjwsc.csv, K
	If ErrorLevel
	{
		Sm()
	}
	else
	{
		Run,D:\CGSS调查问卷统计\数据质控测试文档\%now_date% qstjwsc.csv
		Ddck("",1,"2017")
		Fs("^j",2)
		Fs("!{F4}")
		Ddck("Microsoft Excel")
		Fs("{RIGHT}")
		Fs("{ENTER}",1)		
		DetectHiddenWindows, On
		WinActivate, CGSS2017电话核查样本更新 - Google 表格 - Google Chrome, 
		Sm(1)
		Fs("^+v")
		Sc(292,679,292,679,"0xFFF3EC",0)
	}

	;9.5 上传低电话回收访员到电核大表
	Fs("^+{PGDN}")
	Sc(145,264,145,264,"0xA8D7B6",0)
	Fs("^{DOWN}")
	Sc(11,674,11,674,"0xDDDDDD",0)
	Fs("{DOWN}")
	FileGetSize, OutputVar, D:\CGSS调查问卷统计\数据质控测试文档\%now_date% 低电话回收率未上传访员名单.csv, K
	If ErrorLevel
	{
		Sm()
	}
	else
	{
		Run,D:\CGSS调查问卷统计\数据质控测试文档\%now_date% 低电话回收率未上传访员名单.csv
		Ddck("",1,"2017")
		Fs("^j",2)
		Fs("!{F4}")	
		Ddck("Microsoft Excel")
		Fs("{RIGHT}")
		Fs("{ENTER}")			
		DetectHiddenWindows, On
		WinActivate, CGSS2017电话核查样本更新 - Google 表格 - Google Chrome, 
		Sm(1)
		Fs("^+v")
	}

	;15.1 登录QQ邮箱
	Run,https://mail.qq.com
	Ddck("登录QQ邮箱 - Google Chrome",0.5)
	Fs("Xiaozhou13171317",1)
	Fs("{ENTER}")
	Ddck("QQ邮箱",1)
	
;——————————————————
;16.给雨薇发邮件
;——————————————————

	;16.1 进入写信页面,填写收件人、邮件主题和邮件正文
	Zjdj(97,210,2)
	Srwz("snailallez@163.com;")
	Fs("{ENTER}",0.5)
	Loop 2
	{
		Fs("{TAB}")
	}
	Send,%now_date% 调查进度与无联系记录样本
	Sm(1)
	Fs("{TAB}")
	Run,D:\CGSS调查问卷统计\数据质控测试文档\%now_date% toyw.txt
	Sm(2)
	Fs("^a")
	Fs("^c")
	Fs("!{F4}")
	Ddck("QQ邮箱")
	Fs("^v")
	Zjdj(550,718)
	
	;16.2 添加附件
	Zjdj(385,227)
	Ddck("打开",1)
	Fs("^f",1)
	Fs("+{TAB}",0.5)
	Fs("{ENTER}",0.5)
	Srwz("D:\CGSS调查问卷统计\数据质控测试文档")
	Fs("{ENTER}",0.5)
	Fs("^f",1)
	Send,NOT “北京地区” AND “%now_date% 调查进度” OR “%now_date% 无联系”
	Sm(4)
	Fs("{ENTER}",0.5)
	Loop 3
	{
		Fs("{TAB}",0.5)
	}
	Fs("^a")
	Loop 3
	{
		Fs("{TAB}",0.5)
	}
	Fs("{ENTER}",0.5)
	
	;16.3 发送邮件
	Zjdj(294,712)
	Loop 6
	{
		Fs("{DOWN}")
	}
	Zjdj(294, 712)
	
	;16.4 判断邮件是否发送完毕
	Sc(292,252,292,252,"0xBAF2D7",0)


;——————————————————
;17.给福鑫、贾敏发邮件
;——————————————————	
	
	;17.1 进入写信页面,填写收件人、邮件主题和邮件正文
	Zjdj(97,210,2)
	Srwz("18170146171@163.com;")
	Srwz("1131738451@qq.com;")
	Fs("{ENTER}",0.5)
	Fs("{TAB}",0.5)
	Fs("{TAB}",0.5)
	Send,%now_date% 调查进度与原始数据
	Sm(1)
	Fs("{TAB}",0.5)
	Run,D:\CGSS调查问卷统计\数据质控测试文档\tofxjm.txt
	Sm(2)
	Fs("^a")
	Fs("^c")
	Fs("!{F4}")
	Ddck("QQ邮箱")
	Fs("^v")
	Fs("{TAB}")
	Zjdj(550,718)
	
	;17.2 添加附件
	Zjdj(385,227)
	Ddck("打开",1)
	Fs("^f",1)
	Fs("+{TAB}",0.5)
	Fs("{ENTER}",0.5)
	Srwz("D:\CGSS调查问卷统计\数据质控测试文档")
	Fs("{ENTER}",0.5)
	Fs("^f",1)
	Send,NOT “北京地区” AND “%now_date% 调查进度” OR “%now_date% 无电话已完成” OR “%now_date% 原始数据”
	Sm(4)
	Fs("{ENTER}",0.5)
	Loop 3
	{
		Fs("{TAB}",0.5)
	}
	Fs("^a")
	Loop 3
	{
		Fs("{TAB}",0.5)
	}
	Fs("{ENTER}",3)
	
	Zjdj(711,502,1)
	
	;17.3 发送邮件
	Zjdj(294,712)
	Loop 6
	{
		Fs("{DOWN}")
	}
	Zjdj(294, 712)

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
;11.休眠，关闭硬盘，为内存通电
;——————————————————

	DllCall("PowrProf\SetSuspendState", "int", 1, "int", 1, "int", 0)	
	
Return
