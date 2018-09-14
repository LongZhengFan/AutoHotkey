#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.


^+l::
	InputBox, maxnum, 添加内容, 请输入您所需添加内容的行数
	Loop, %maxnum%
	{
		SendInput,{APPSKEY}
		Sleep,100
		SendInput,m
		Sleep,100
		SendInput,{HOME}
		Sleep,100
		Loop,12
		{
			SendInput,{Right}
		}
		SendInput,{DELETE}
		Sleep,100
		SendInput,+T
		Sleep,100
		SendInput,{ENTER}
		Sleep,100
		SendInput,{DOWN}
	}	
Return

:*:xyg::
	SendInput,{END}{DOWN}{END}{ENTER}{ENTER}
	Sleep, 100
	SendInput, {SHIFT}<blockquote>
	Sleep,200
	SendInput, {ENTER}{ENTER}{UP}{SHIFT}
Return