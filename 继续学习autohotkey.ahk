#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

;一、自定义函数部分

;——————————————————————
;1.睡眠函数
;——————————————————————

;函数旨在将Sleep函数的时间参数单位由ms修改为s
Sm(ti=0.1)
{
	ti := ti*1000
	Sleep, ti
}

;——————————————————————
;2.等待窗口函数
;——————————————————————

;wn为窗口完整名称, wt为窗口完整名称的子字符串
Ddck(wn,ti=0.1,wt="")
{
	WinWait, %wn%, %wt%,
	IfWinNotActive, %wn%, %wt%, WinActivate, %wn%, %wt%, 
	WinWaitActive, %wn%, %wt%,
	Sm(ti)
}

;——————————————————————
;3.输入文字函数
;——————————————————————

;word为需要输入的文字，wait为ctrl+v后的睡眠时间
Srwz(word,ti=0.2)
{
	Clipboard = %word%
	Sm()
	Send,^v
	Sm(ti)
}

;——————————————————————
;4.发送按键函数
;——————————————————————

;函数旨在为所有Send命令添加一行Sleep命令，避免手工多次输入Sleep的麻烦
;默认休息时间为0.1秒，休息时间由该函数第2个参数确定
Fs(na,ti=0.1)
{
	Send,%na%
	Sm(ti)
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
	Sm(ti)
}

;——————————————————————
;6.左键单击函数
;——————————————————————

Zjdj(a,b,ti=0.5)
{
	MouseClick,left,a,b
	Sm(ti)
}

:*:ffss::
	;SoundBeep第一个参数代表声音频率，第二个代表持续时间
	SoundBeep,423,500*6
	Sm(1)
	;SoundPlay用于播放系统声音时，本系统仅下面两种可选声音
	SoundPlay,*-1
	Sm(1)
	SoundPlay,*16
	Sm(1)
	;SoundPlay也可用于播放外部声音文件(支持mav,mp3)
	SoundPlay,D:\personal\音乐\陈楚生_-_有没有人告诉你.mp3
Return