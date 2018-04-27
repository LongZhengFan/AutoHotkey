#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

Fs(na,ti=0.1)
{
	SendInput,%na%
	ti := ti*1000
	Sleep, ti
}

;git add .
:*:gadd::
	SendInput,git add .{ENTER}
Return

;git commit -m ""
:*:gcom::
	InputBox, word, 提交说明, 请输入对此次更改的必要文字说明 
	SendInput,git commit -m "%word%"{ENTER}
Return

;git remote add origin git@github.com:username/caname.git
:*:grad::
	InputBox, caname, 目录名, 请输入github上的目录名
	SendInput, git remote add origin git@github.com:tiny-boat/%caname%.git{ENTER}
Return

;git push -u origin master
:*:gpuu::
	SendInput, git push -u origin master{ENTER}
Return

;git push origin master
:*:gpus::
	SendInput, git push origin master{ENTER}
Return

:*:yjtj::
	InputBox, bdname, 目录名, 请输入本地目录名
	SendInput, cd '%bdname%'{ENTER}
	Sleep,2000
	SendInput, git init{ENTER}
	Sleep,1000
	SendInput, git add .{ENTER}
	InputBox, word, 提交说明, 请输入对此次更改的必要文字说明 
	SendInput,git commit -m "%word%"{ENTER}
	Sleep,1000
	InputBox, caname, 目录名, 请输入github上的目录名
	SendInput, git remote add origin git@github.com:tiny-boat/%caname%.git{ENTER}
	Sleep,1000
	SendInput, git push -u origin master{ENTER}
Return