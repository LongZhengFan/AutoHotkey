#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

:*:bgfl::
  MouseClick, left,  75,  136
  Send,{ALT}{a}{e}  
  Sleep, 1000
  Send,{n}
  Sleep, 1000
  Send,{t}{s}{n}
  Sleep,1000
  Send,{f}
  Sleep,1000
  Send,{enter}
Return