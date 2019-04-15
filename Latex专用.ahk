#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

:*:lhn::
	SendInput, {Raw}$$
	SendInput, {Left 1}
Return

:*:lhw::
	SendInput, {Raw}$$$$
	SendInput, {Left 2}
Return

:*:ljc::
	SendInput, {Raw}\boldsymbol{}
	SendInput, {Left}
Return

:*:lyy::
	SendInput, {right 1}
Return

:*:lzy::
	SendInput, {left 1}
Return

:*:lxkh::
	SendInput, {Raw}\left( \right )
	SendInput, {Left 9}
Return

:*:ltcx::
	SendInput, {right 9}
Return

:*:lzkh::
	SendInput, {Raw}\left[ \right ]
	SendInput, {Left 9}
Return

:*:ltcz::
	SendInput, {right 9}
Return

:*:ldkh::
	SendInput, {Raw}\left \{  \right \}
	SendInput, {Left 10}
Return

:*:ltcd::
	SendInput, {right 10}
Return

:*:lfs::
	SendInput, {Raw}\left \|  \right \|
	SendInput, {Left 10}
Return

:*:ljdz::
	SendInput, {Raw}\left |  \right |
	SendInput, {Left 9}
Return

:*:ljz::
	SendInput, {Raw}\bar{}
	SendInput, {Left}
Return

:*:lsb::
	SendInput, {Raw}^{}
	SendInput, {Left}
Return

:*:lxb::
	SendInput, {Raw}_{}
	SendInput, {Left}
Return


:*:lqh::
	SendInput, {Raw}\sum_{}^{}
	SendInput, {Left 4}
Return