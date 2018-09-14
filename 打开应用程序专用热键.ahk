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
	SendInput,^v
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
	SendInput,%na%
	ti := ti*1000
	Sleep, ti
}

;——————————————————————
;5.寻找颜色函数
;——————————————————————

Sc(a,b,c,d,color,var,ti=0.5)
{

		PixelSearch, LogoX, LogoY, a, b, c, d, %color%, var, Fast
		Sleep,100
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


;以下三个为联网程序

:*:dklantern::
  Run,C:\Users\xiaozhou\AppData\Roaming\Lantern\lantern.exe
Return

;以下四个为浏览器程序

:*:dkie::
  Run,C:\Program Files\Internet Explorer\iexplore.exe
Return

:*:dkedge::
  Run,C:\Windows\SystemApps\Microsoft.MicrosoftEdge_8wekyb3d8bbwe\MicrosoftEdge.exe
Return

:*:dkchrome::
  Run,C:\Program Files (x86)\Google\Chrome\Application\chrome.exe
Return

:*:ggss::
	InputBox, keyword, Google 搜索, 亲爱的主人，请输入您需要搜索的内容
	Run,https://www.google.com.hk
	Ddck("Google - Google Chrome",1)
	SendInput, {SHIFT}%keyword%{ENTER}{SHIFT}
Return

:*:bdss::
	InputBox, keyword, Baidu 搜索, 亲爱的主人，请输入您需要搜索的内容
	Run,https://www.baidu.com	
	Ddck("百度一下，你就知道 - Google Chrome",1)
	SendInput, {SHIFT}%keyword%{ENTER}{SHIFT}
Return

;以下三个为办公软件程序

:*:dkword::
  Run,C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE
Return

:*:dkexcel::
  Run,C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE
Return

:*:dkppt::
  Run,C:\Program Files (x86)\Microsoft Office\root\Office16\POWERPNT.EXE
Return

;以下四个为统计软件程序

:*:dkryy::
  Run,D:\Program Files\R-3.4.1\bin\x64\Rgui.exe
Return

;以下两个为数学软件程序

:*:dkmatlab::
  Run,D:\Matlab 2014b\bin\matlab.exe
Return

:*:dkmathematica::
  Run,D:\应用程序\Mathematica 10\SystemFiles\FrontEnd\Binaries\Windows-x86-64\Mathematica.exe
Return

:*:dkcyy::
  Run,D:\应用程序\Dev-Cpp\devcpp.exe
Return

;以下为脚本编辑器程序

:*:dknotepad::
  Run,D:\应用程序\Notepad++\notepad++.exe
Return

:*:dksublime::
  Run,D:\Program Files\Sublime Text 3\sublime_text.exe
Return

;以下四个为多媒体程序

:*:dkwyy::  		
	Run,D:\应用程序\网易云音乐\CloudMusic\cloudmusic.exe
Return

:*:dkhshy::
  Run,D:\应用程序\Corel VideoStudio Pro X6\vstudio.exe
Return

:*:dkps::
  Run,D:\应用程序安装包\Adobe_Photoshop_CS6_x64\Adobe Photoshop CS6 (64 Bit)\Adobe Photoshop CS6 (64 Bit)\Photoshop.exe
Return

;以下两个为聊天程序

:*:dkqq::
  Run,D:\应用程序\qq\Bin\QQ.exe
Return

:*:dlqq::
  Run,D:\应用程序\qq\Bin\QQ.exe
  Ddck("QQ",4)
  SendInput,Xiaozhou13171317{enter}
Return

:*:dkwechat::
  Run,D:\应用程序\WeChat\WeChat.exe
Return

;以下为其他程序

:*:dkbdy::
  Run,D:\应用程序\百度云\BaiduYunGuanjia\baiduyunguanjia.exe
Return

:*:dkautosw::
  Run,D:\应用程序\AutoScriptWriter\AutoScriptWriter.exe
Return

:*:dk360::
  Run,D:\应用程序\360\360Safe\360Safe.exe
Return

:*:dkjscb::
  Run,D:\应用程序\Kingsoft\Power Word 2016\2016.3.3.0316\PowerWord.exe
Return

;以下为打开autohotkey脚本程序

:*:dkahk::

;以下为打开Git、ruby程序

:*:dkgit::
	Run,C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Git\Git Bash
	Ddck("mintty")
	SendInput,{ENTER}
Return

:*:dkjkfw::
	WinActivate, Start Command Prompt with Ruby
	IfWinNotActive, Start Command Prompt with Ruby
	{
		Run,C:\Users\xiaozhou\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Ruby 2.4.4-1-x64 with MSYS2\Start Command Prompt with Ruby
		Ddck("Start Command Prompt with Ruby")
		SendInput,{SHIFT}jekyll serve{ENTER}
	}
	else{
		SendInput,jekyll serve{ENTER}
	}
Return

:*:gbjkfw::
	WinActivate, jekyll  serve
	SendInput,^c
	Sleep,1000
	SendInput,Y{ENTER}
Return

; 以下为湘潭大学校园网认证程序

#w::
	Run, http://172.16.0.32:8080/eportal/index.jsp?wlanuserip=0ab5b1c3c6fd84fe5b8faa120d1d0763&wlanacname=776e154f4e46ba9668906e203954b23a&ssid=&nasip=0d09a1de724f795b5b8c3bae719e0bd7&snmpagentip=&mac=8047be3af5818b81223df485a83a950b&t=wireless-v2&url=d20b79da3e12046d789ef330d087c33e451856b9c4d64bb892fec938e15f8ae20fe2efa844813086a2abb2f0115b731bd44d644760b565f8d58b4ea40c8bc49f&apmac=&nasid=776e154f4e46ba9668906e203954b23a&vid=a48bd0765eb8bc00&port=30e2ab4d2df9b52e&nasportid=5b9da5b08a53a5405fcca25f7e4ce14433a731b72e9fbd08c52fa2551b6cad1a
	Ddck("上网认证 - Google Chrome",1)
	Fs("{SHIFT}201821511193{TAB}054717{ENTER}")
Return

#c::
	WinActivate, 登录成功 - Google Chrome
	Sleep, 100
	PixelSearch, Px, Py, 0, 0, 1362, 249, 0x78CF43, 3, Fast
	MouseClick, left, %Px%, %Py%
	Sleep, 1000
	PixelSearch, Rx, Ry, 0, 0, 1362, 400, 0xF48542, 3, Fast
	MouseClick, left, %Rx%, %Ry%
	Ddck("您已成功下线 - Google Chrome")
	SendInput, ^w
Return

#s::
	loop
	{
		SendInput, {DOWN}{HOME}{RIGHT}
		Sleep, 1500
		SendInput, ^c
		Sleep, 300
		promf := clipboard
		SendInput, {RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}
		Sleep, 300
		SendInput, ^c
		Sleep, 300
		promt := clipboard
		SendInput, {RIGHT}
		Sleep, 1500
		SendInput, ^c
		Sleep, 300
		proms := clipboard
		MsgBox, 3, 是否继续, 是否涉及量化分析
		IfMsgBox, YES 
			SendInput, {LEFT}{LEFT}{LEFT}0{TAB}
		else
			IfMsgBox, NO
			{
				Sleep, 300
				SendInput, {LEFT}{LEFT}{LEFT}{LEFT}
				Sleep, 300
				InputBox, cate, 类别输入, 请输入该论文所属类别——%promf%——%promt%, , 1000, 600
				Sleep, 300
				SendInput, {SHIFT}%cate%{SHIFT}
				Sleep, 300
				SendInput, {ENTER}{UP}{RIGHT}1
				Sleep, 300
				SendInput, {ENTER}
				Sleep, 300
				SendInput, {UP}{RIGHT}
				Sleep, 300
				InputBox, mode, 模型输入, 请输入该论文所用模型——%proms%
				Sleep, 300
				SendInput, {SHIFT}%mode%{SHIFT}{ENTER}{UP}{RIGHT}
			}
			else
				break
	}
Return