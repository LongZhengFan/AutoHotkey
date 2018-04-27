#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

;触发定时关机程序
:*:dgj::
  Send,{LWin down}r{LWin up}
  Sleep,1000
  Send,{Shift}Shutdown -s -t{space}
Return