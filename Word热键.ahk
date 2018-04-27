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

;—————————————
;—————————————
;程序主体部分
;—————————————
;—————————————

;—————————
;插入公式
;—————————

:*:crgs::
   SendInput,{Alt}{n}{e}{i}{SHIFT}{CTRLDOWN}i{CTRLUP}
Return

	;公式模式下的小括号
	:*:xkh::
	   SendInput,{BACKSPACE}{Alt}{j}{e}{b}{enter}{left}
	Return

	
;—————————
;切换字体
;—————————

#z::
	InputBox, fontname, 字体名称, 您需要更改为哪种字体
	SendInput,{APPSKEY}f%fontname%{ENTER}
Return

	;切换字体为华文楷体
	:*:qhkt::
		SendInput,{APPSKEY}f
		SendInput,{raw}华文楷体
		SendInput,{ENTER}
	Return

	;切换字体为宋体
	:*:qhst::
		SendInput,{APPSKEY}f
		SendInput,{raw}宋体
		SendInput,{ENTER}
	Return

	
;—————————
;文字选择
;—————————	

^+x::
	MsgBox, 4, 方式选择, 如您选择按行选择，单击YES，否则单击NO, 8
	IfMsgBox, YES
	{
		InputBox, charnum, 行数选择, 请输入所需选择的行数
		MsgBox, 4, 方向选择, 如您选择向下，单击YES，否则单击NO, 8
		IfMsgBox, YES
			SendInput,{SHIFTDOWN}{DOWN %charnum%}{SHIFTUP}
		else
			SendInput,{SHIFTDOWN}{UP %charnum%}{SHIFTUP}
	}
	else
	{
		InputBox, charnum, 文字选择, 请输入所需选择的文字数
		MsgBox, 4, 方向选择, 如您选择向右，单击YES，否则单击NO, 8
		IfMsgBox, YES
			SendInput,{SHIFTDOWN}{RIGHT %charnum%}{SHIFTUP}	
		else
			SendInput,{SHIFTDOWN}{LEFT %charnum%}{SHIFTUP}
	}
Return
	
	
;—————————
;光标移动
;—————————
	
	;光标左移
	:*:gbzy::
		InputBox, leftnum, 光标左移, 您需要左移多少个单位
		SendInput,{LEFT %leftnum%}
	Return

	;光标右移
	:*:gbyy::
		InputBox, rightnum, 光标右移, 您需要右移多少个单位
		SendInput,{RIGHT %rightnum%}
	Return

	;光标上移
	:*:gbsy::
		InputBox, upnum, 光标上移, 您需要上移多少个单位
		SendInput,{UP %upnum%}
	Return

	;光标下移
	:*:gbxy::
		InputBox, downnum, 光标下移, 您需要下移多少个单位
		SendInput,{DOWN %downnum%}
	Return

/*
;一段用于搜索网页关键字的无聊的程序，折腾了我一天，最后把google整得神魂颠倒
;2018年1月27日……	
^a::
	;年代尾号
	an := 9
	;结果计数
	i := 0
	Loop
	{
		an := an + 1
		;班级尾号
		cn := 15
		Loop
		{
			cn := cn + 1
			;班级代码
			bn := an*100 + cn
			;网页地址年级尾号
			ynu := 2*floor(an/2)
			ynd := ynu + 1
			If an < 10
				Run,http://58.47.177.235:8818/Admin/WinXin/ClassList.aspx?name=高0%bn%班&grade=200%ynu%年-200%ynd%年
			else
				Run,http://58.47.177.235:8818/Admin/WinXin/ClassList.aspx?name=高%bn%班&grade=20%ynu%年-20%ynd%年
			Ddck("校友名录 - Google Chrome")
			Fs("^f")
			Fs("{raw}隆")
			Fs("{ENTER}")
			Sc(0,0,1355,743,"0x3296FF",0)
			If ErrorLevel
				Mouseclick, left, 1348, 20
			else
			{
				i := i + 1
				RS%i% := bn
				;MsgBox, 4, 通知, 发现目标%bn%班，是否继续
				Mouseclick, left, 1348, 20				
			}
			
			if cn>20
				break
			else
				continue
		}
		If an > 14
			break
		else
			continue
	}
	ListVars
Return
*/