
;--- StatusBar

SB_ISSIMPLE		equ (WM_USER+14)
SB_SETTIPTEXT	equ (WM_USER+16)

StatusBar_GetText macro hWnd, iPart, pText
	invoke SendMessage, hWnd, SB_GETTEXT, iPart, pText
	endm 
StatusBar_SetText macro hWnd, iPart, pText
	invoke SendMessage, hWnd, SB_SETTEXT, iPart, pText
	endm 
StatusBar_GetTextLength macro hWnd, iPart
	invoke SendMessage, hWnd, SB_GETTEXTLENGTH, iPart, 0
	endm 
StatusBar_SetSimpleMode macro hWnd, bFlag
	invoke SendMessage, hWnd, SB_SIMPLE, bFlag, 0
	endm 
StatusBar_SetParts macro hWnd, iParts, pdwWidths
	invoke SendMessage, hWnd, SB_SETPARTS, iParts, pdwWidths
	endm 
StatusBar_IsSimple macro hWnd
	invoke SendMessage, hWnd, SB_ISSIMPLE, 0, 0
	endm
StatusBar_SetTipText macro hWnd, iPart, pszText
	invoke SendMessage, hWnd, SB_SETTIPTEXT, iPart, pszText
	endm
