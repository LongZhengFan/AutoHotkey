#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.


#t::

	InputBox, time, 休息提醒器, 请输入一个时间（单位：分），系统将在这个时间之后提醒您休息
	; 弹出一个输入框，标题是“休息提醒器”，内容是“请输入一个时间……”
	tiou := time*60*1000
	; 如果一个变量要做计算的话，一定要像这样写，和平常的算式相比，多了一个冒号
	Sleep,%tiou%

	SetTimer, ChangeButtonNames, 50 
	MsgBox, 3, 休息提醒, 亲爱的主人，您已经辛勤工作 %time% 分钟了，休息一下吧！
	IfMsgBox, YES 
		MsgBox, , 一叶扁舟, 我的心心送给你，主人（づ￣3￣）づ╭❤～ ,3
	else 
		IfMsgBox, No
			MsgBox, , 一叶扁舟, 记得尽早休息哦，主人(づ￣ 3￣)づ)
		else
			MsgBox, , 一叶扁舟, 主人你不理我(；′⌒`)),3

Return 

ChangeButtonNames: 
	IfWinNotExist, 休息提醒
		return  ; Keep waiting.
	SetTimer, ChangeButtonNames, off 
	WinActivate 
	ControlSetText, Button1, 我这就休息 
	ControlSetText, Button2, 我再干会儿
Return


#G::
	MsgBox, 3, 休眠提醒, 亲爱的主人，您确定要让计算机进入休眠状态吗?, 8
	
	IfMsgBox, Timeout
	{
		MsgBox, , 休眠提醒, 超时了
	}
	
	IfMsgBox, YES
	{
		MsgBox, , 休眠提醒, 再见，主人（づ￣3￣）づ╭❤～,2
		;Sleep,2400000
		DllCall("PowrProf\SetSuspendState", "int", 1, "int", 1, "int", 0)
	}
	 
	IfMsgBox, No
		MsgBox, , 休眠提醒, 好的，主人 (づ￣ 3￣)づ), 2
	
	IfMsgBox, Cancel
		MsgBox, , 休眠提醒, 主人你不理我(；′⌒`)), 2
	
Return


^Left::
	MouseMove, -25,  0, 0, R
Return

^Up::    
	MouseMove,  0, -25, 0, R
Return

^Right:: 
	MouseMove,  25,  0, 0, R
Return

^Down::  
	MouseMove,  0,  25, 0, R
Return

!Left::
	MouseMove, -5,  0, 0, R
Return

!Up::    
	MouseMove,  0, -5, 0, R
Return

!Right:: 
	MouseMove,  5,  0, 0, R
Return

!Down::  
	MouseMove,  0,  5, 0, R
Return

^ENTER::  
	MouseClick,left,0,0,1,0,,R
Return

^APPSKEY:: 
	MouseClick,right,0,0,1,0,,R
Return

/*
:*:z::
	InputBox, num, , 
	Loop,%num%
	{
		MouseClick,left,660,532
		Sleep,400
	}
Return
*/

:*:wykj::

	InputBox, num, 前进步数,请输入快进次数
	Loop, %num%
	{
		Send,{RIGHT}
		Sleep,500
	}
Return

:*:wykt::

	InputBox, num, 后退步数,请输入快退次数
	Loop, %num%
	{
		Send,{LEFT}
		Sleep,500
	}
Return

#A::
	SendInput, <p>{DELETE}{DELETE}{DELETE}{DELETE}
	Sleep, 100
	Loop, 10
	{
		SendInput,{END}
	}
	Sleep, 100
	SendInput, </p>{DOWN}{DOWN}{HOME}
Return