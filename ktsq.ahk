#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

#s::
	loop
	{
		SendInput, {DOWN}{HOME}{RIGHT}
		Sleep, 2000
		SendInput, ^c
		Sleep, 500
		promf := clipboard
		SendInput, {RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}
		Sleep, 500
		SendInput, ^c
		Sleep, 500
		promt := clipboard
		SendInput, {RIGHT}
		Sleep, 2000
		SendInput, ^c
		Sleep, 500
		proms := clipboard
		MsgBox, 3, 是否继续, 是否涉及量化分析
		IfMsgBox, YES 
			SendInput, {LEFT}{LEFT}{LEFT}0{TAB}
		else
			IfMsgBox, NO
			{
				Sleep, 500
				SendInput, {LEFT}{LEFT}{LEFT}{LEFT}
				Sleep, 500
				InputBox, cate, 类别输入, 请输入该论文所属类别——%promf%——%promt%, , 1000, 600
				Sleep, 500
				SendInput, {SHIFT}%cate%{SHIFT}
				Sleep, 500
				SendInput, {ENTER}{UP}{RIGHT}1
				Sleep, 500
				SendInput, {ENTER}
				Sleep, 500
				SendInput, {UP}{RIGHT}
				Sleep, 500
				InputBox, mode, 模型输入, 请输入该论文所用模型——%proms%
				Sleep, 500
				SendInput, {SHIFT}%mode%{SHIFT}{ENTER}{UP}{RIGHT}
			}
			else
				break
	}
Return