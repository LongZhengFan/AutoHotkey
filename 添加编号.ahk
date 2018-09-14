#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.


:*:hwtjnr::
	InputBox, maxnum, 添加内容, 请输入您所需添加内容的行数
	;InputBox, content, 添加内容，请输入您所需添加的内容
	;originnum := 0
	Loop, %maxnum%
	{
		;originnum := originnum + 1
		Loop,10
		{
			SendInput,{END}
		}
		Sleep,100
		;SendInput,{ENTER}
		SendInput,<br>
		Sleep,100
		SendInput,{DOWN}
	}	
Return