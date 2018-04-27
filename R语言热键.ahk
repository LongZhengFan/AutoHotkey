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


;启动R语言
#s::
	Run,D:\Program Files\R-3.4.1\bin\x64\Rgui.exe
	Ddck("RGui (64-bit)",2)
	Fs("{ALT}")
	Loop 3
	{
		Fs("{DOWN}")
	}
	Fs("{ENTER}")
Return

;创建新程序脚本
:*:wyxjb::
	Fs("!{F4}")
	Fs("{ALT}")
	Loop 2
	{
		Fs("{DOWN}")
	}
	Fs("{ENTER}")
Return

;R语言清空内存
:*:qknc::
	Srwz("rm(list=ls())")
	Fs("{ENTER}{ENTER}")
Return

;R语言安装程序包
:*:azcxb::
	InputBox, cxbna, 程序包名称, 请输入所要加载程序包的名称
	Srwz("install.packages(")
	Send,"%cxbna%"
	Srwz(")")
	Send,{ENTER}
Return

;R语言加载程序包
:*:jzcxb::
	InputBox, cxbna, 程序包名称, 请输入所要加载程序包的名称
	Srwz("library(")
	Send,"%cxbna%"
	Srwz(")")
	Send,{ENTER}
Return

;R语言赋值运算符<-（左箭头）
:*:zjt::
	Srwz("<-")
	Fs("{SPACE}")
Return

;R语言自定义函数快捷键
:*:zdyhs::
	Srwz("function(x){")
	Fs("{ENTER}{ENTER}")
	Srwz("}")
	Fs("{UP}{TAB}")
Return

;R语言函数括号（小括号）
:*:xkh::
	Fs("{BACKSPACE}")
	Srwz("()")
	Fs("{LEFT}")
Return

;R语言索引（中括号）
:*:zkh::
	Fs("{BACKSPACE}")
	Srwz("[]")
	Fs("{LEFT}")
Return

;R语言条件/循环（大括号）
:*:dkh::
	Srwz("{")
	Fs("{ENTER}{ENTER}")
	Srwz("}")
	Fs("{UP}{TAB}")
Return

;R语言结束条件/循环/函数体编辑
:*:jsbd::
	Fs("{DOWN}{ENTER}{ENTER}")
Return

;R语言进入下一行编写
:*:xyh::
	Fs("{BACKSPACE}{END}{ENTER}")
Return

;R语言跳出括号
:*:rrr::
	Fs("{BACKSPACE}{Right}")
Return

;R语言跳到下一个引号内
:*:rrs::
	Fs("{BACKSPACE}{Right}{Right}{Right}")
Return

;R语言单个单引号
:*:dyh::
	Send,{BACKSPACE}{'}{'}{LEFT}
Return

;R语言单个双引号
:*:yh::
	Send,{BACKSPACE}{"}{"}{LEFT}
Return

;R语言多个双引号
:*:wyyh::
	InputBox, num, 引号个数, 请输入所需引号个数
	Send,{BACKSPACE}
	Loop, %num%
	{
		Send,{"}{"}{,}
	}
	Fs("{BACKSPACE}")
	num2 := 3*num-2
	Loop, %num2%
	{
		Send,{LEFT}
	}
Return

;R语言字符串向量
:*:wyzfc::
	
	ArrayCount := 0
	InputBox, num, 字符串个数输入框, 亲爱的主人，您需要输入多少个字符串？`n`n如果您现在还不清楚的话，直接敲回车键就好了
	
	;依据用户是否输入了字符串个数分别执行有限循环和无限循环
	If num =
	{
		Loop
		{
			ArrayCount += 1
			InputBox, Array%ArrayCount%, 字符串内容输入框, 亲爱的主人，请您输入第 %ArrayCount% 个字符串`n`n如需结束，请您在回车后立即按住 End 键不放
			Sm(0.5)
			GetKeyState, state, End
			If state = D
			{
				MsgBox, 4, 结束提醒, 亲爱的主人，到目前为止，您共计输入了 %ArrayCount% 个字符串，确定要结束本次输入吗？
			}
			IfMsgBox, YES 
				break
			else
				continue
		}
	}
	else
	{
		Loop, %num%
		{
			ArrayCount += 1
			InputBox, Array%ArrayCount%, 字符串内容输入框, 亲爱的主人，请您输入第 %ArrayCount% 个字符串 
		}
	}
	
	Clipboard = "
	Fs("^v")
	Loop, %ArrayCount%
	{
		element := Array%A_Index%
		Clipboard = %element%","
		Fs("^v")
	}
	Fs("{BACKSPACE}{BACKSPACE}")		

Return

;R语言字符串向量
:*:wysz::
	
	ArrayCount := 0
	InputBox, num, 数字个数输入框, 亲爱的主人，您需要输入多少个数字？`n`n如果您现在还不清楚的话，直接敲回车键就好了
	
	;依据用户是否输入了数字个数分别执行有限循环和无限循环
	If num =
	{
		Loop
		{
			ArrayCount += 1
			InputBox, Array%ArrayCount%, 数字内容输入框, 亲爱的主人，请您输入第 %ArrayCount% 个数字`n`n如需结束，请您在回车后立即按住 End 键不放
			Sm(0.5)
			GetKeyState, state, End
			If state = D
			{
				MsgBox, 4, 结束提醒, 亲爱的主人，到目前为止，您共计输入了 %ArrayCount% 个数字，确定要结束本次输入吗？
			}
			IfMsgBox, YES 
				break
			else
				continue
		}
	}
	else
	{
		Loop, %num%
		{
			ArrayCount += 1
			InputBox, Array%ArrayCount%, 数字内容输入框, 亲爱的主人，请您输入第 %ArrayCount% 个数字 
		}
	}
	
	Loop, %ArrayCount%
	{
		element := Array%A_Index%
		Clipboard = %element%,
		Fs("^v")
	}
	Fs("{BACKSPACE}")		

Return

#z::
	Loop 1
	{
		MouseClick,right,,,
		Sleep,100
		SendInput,d
		Sleep,1000

	}
Return