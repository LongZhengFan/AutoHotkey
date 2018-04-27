#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.


:*:tempt::
  Send, :*:dk::{enter}
  Send, {space}{space}Send,{enter}
  Send, Return{up}{up}{left}
Return

:*:scqyh1::
  Send, {home}{right}{right}{right}{right}{right}{right}{delete}{down}{enter}{enter}
Return

:*:scqyh2::
  Send, {home}{up}{right}{right}{right}{right}{right}{right}{delete}{down}{down}{enter}{enter}
Return