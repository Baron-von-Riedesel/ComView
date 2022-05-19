
;*** show dialog of event protocols

	.486
	.model flat, stdcall
	option casemap :none
	option proc :private

	include COMView.inc
INSIDE_CLOGWINDOW equ 1
	include classes.inc
	include rsrc.inc
	include debugout.inc

?USEDEBUGOUT	equ 0	;modify entry OutputDebugString in IAT (dont use anymore!)

BEGIN_CLASS CLogWindow, CDlg
hWndLV	HWND	?
iLine	DWORD	?
bNewLine BOOLEAN	?
END_CLASS

__this	textequ <ebx>
_this	textequ <[__this].CLogWindow>
thisarg	textequ <this@:ptr CLogWindow>

	MEMBER hWnd, pDlgProc, hWndLV, iLine, bNewLine

	.data

g_pNotificationDlg		LPVOID NULL
g_rect					RECT {0,0,0,0}
g_hFont					HFONT NULL
if ?USEDEBUGOUT
g_oldOutputDebugString	DWORD 0
endif
g_bLogActive			BOOLEAN FALSE

	.data?
g_lf					LOGFONT <>

	.code

;*** user pressed right mouse button, show context menu ***


ShowContextMenu proc uses esi

local	pt:POINT
local	hPopupMenu:HMENU

		invoke GetSubMenu,g_hMenu, ID_SUBMENU_NOTIFICATIONDLG
		.if (eax != 0)
			mov hPopupMenu, eax
			invoke GetCursorPos,addr pt
			invoke ListView_GetSelectedCount( m_hWndLV)
			.if (eax)
				mov esi, MF_ENABLED
			.else
				mov esi, MF_GRAYED or MF_DISABLED
			.endif
			invoke EnableMenuItem, hPopupMenu, IDM_CUT, esi
			invoke EnableMenuItem, hPopupMenu, IDM_COPY, esi
			invoke EnableMenuItem, hPopupMenu, IDM_DELETE, esi
			invoke TrackPopupMenu, hPopupMenu, TPM_LEFTALIGN or TPM_LEFTBUTTON,\
					pt.x,pt.y,0,m_hWnd,NULL
		.endif


		ret
		align 4

ShowContextMenu endp


;*** WM_INITDIALOG

OnInitDialog proc

local lvc:LVCOLUMN
local rect:RECT

	invoke GetDlgItem, m_hWnd, IDC_LIST1
	mov m_hWndLV,eax

	.if (g_rect.left)
		invoke SetWindowPos, m_hWnd, NULL, g_rect.left, g_rect.top,\
			g_rect.right, g_rect.bottom, SWP_NOZORDER or SWP_NOACTIVATE
	.endif

	invoke GetClientRect, m_hWndLV, addr rect
	invoke GetSystemMetrics, SM_CXVSCROLL
	sub rect.right,eax
	mov eax, rect.right
	mov lvc.cx_,eax

	mov lvc.mask_,LVCF_TEXT or LVCF_WIDTH

	mov lvc.pszText,CStr("Text")
	invoke ListView_InsertColumn( m_hWndLV, 0, addr lvc)

	invoke ListView_SetExtendedListViewStyle( m_hWndLV, LVS_EX_FULLROWSELECT or LVS_EX_HEADERDRAGDROP or LVS_EX_INFOTIP)

	ret
	align 4

OnInitDialog endp

DeleteItems proc
		invoke SetWindowRedraw( m_hWndLV, FALSE)
		invoke ListView_GetNextItem( m_hWndLV, -1, LVNI_SELECTED)
		.while (eax != -1)
			invoke ListView_DeleteItem( m_hWndLV, eax)
			invoke ListView_GetNextItem( m_hWndLV, -1, LVNI_SELECTED)
		.endw
		invoke SetWindowRedraw( m_hWndLV, TRUE)
		invoke ListView_GetItemCount( m_hWndLV)
		mov m_iLine, eax
		ret
DeleteItems endp

CopyItems proc
		invoke Create@CProgressDlg, m_hWndLV, NULL, SAVE_CLIPBOARD, 1
		invoke DialogBoxParam, g_hInstance, IDD_PROGRESSDLG, m_hWnd, classdialogproc, eax
		ret
CopyItems endp

CutItems proc
		invoke CopyItems
		invoke DeleteItems
		ret
CutItems endp

SelectAllItems	proc
if 0
		SetWindowRedraw m_hWndLV, FALSE
		ListView_GetItemCount m_hWndLV
		mov ecx, eax
		xor eax, eax
		.while (ecx)
			push ecx
			push eax
			ListView_SetItemState m_hWndLV, eax, LVIS_SELECTED, LVIS_SELECTED 
			pop eax
			pop ecx
			inc eax
			dec ecx
		.endw
		SetWindowRedraw m_hWndLV, TRUE
else
		ListView_SetItemState m_hWndLV, -1, LVIS_SELECTED, LVIS_SELECTED 
endif
		ret
SelectAllItems endp

;*** WM_NOTIFY

OnNotify proc uses esi lParam:ptr NMHDR

	mov esi, lParam
	.if ([esi].NMHDR.code == NM_RCLICK)
		invoke ShowContextMenu
	.elseif ([esi].NMHDR.code == LVN_KEYDOWN)
		invoke GetKeyState, VK_CONTROL
		and	al,80h
		.if (!ZERO?)
			.if ([esi].NMLVKEYDOWN.wVKey == 'C')
				invoke CopyItems
			.elseif ([esi].NMLVKEYDOWN.wVKey == 'X')
				invoke CutItems
if 0
			.elseif ([esi].NMLVKEYDOWN.wVKey == 'A')
				invoke SelectAllItems
endif
			.endif
		.elseif ([esi].NMLVKEYDOWN.wVKey == VK_DELETE)
			invoke DeleteItems
		.endif
	.else
		xor eax, eax
	.endif
	ret
	align 4

OnNotify endp

MySetFont proc

local	hDC:HDC
local	hFont:HFONT
local	cf:CHOOSEFONT
local	tm:TEXTMETRIC

		invoke ZeroMemory, addr cf, sizeof cf
		mov cf.lStructSize, sizeof cf
		mov eax, m_hWnd
		mov cf.hwndOwner, eax
		mov cf.Flags, CF_SCREENFONTS
		lea eax, g_lf
		mov cf.lpLogFont, eax
		.if (g_hFont)
			or cf.Flags, CF_INITTOLOGFONTSTRUCT
		.endif
		invoke ChooseFont, addr cf
		.if (eax)
			invoke CreateFontIndirect, addr g_lf
			.if (eax)
				.if (g_hFont)
					push eax
					invoke DeleteObject, g_hFont
					pop eax
				.endif
				mov g_hFont, eax
				invoke SendMessage, m_hWndLV, WM_SETFONT, eax, 0
				invoke InvalidateRect, m_hWndLV, NULL, TRUE
			.endif
		.endif
		ret
MySetFont endp

;*** WM_COMMAND

OnCommand proc wParam:WPARAM, lParam:LPARAM


	movzx eax,word ptr wParam
	.if (eax == IDCANCEL)
		invoke PostMessage, m_hWnd, WM_CLOSE, 0, 0
	.elseif (eax == IDOK)
		invoke PostMessage, m_hWnd, WM_CLOSE, 0, 0
	.elseif (eax == IDM_DELETE)
		invoke DeleteItems
	.elseif (eax == IDM_CUT)
		invoke CutItems
	.elseif (eax == IDM_COPY)
		invoke CopyItems
	.elseif (eax == IDM_SELECTALL)
		invoke SelectAllItems
	.elseif (eax == IDM_CHOOSEFONT)
		invoke MySetFont
	.endif
	xor eax,eax
	ret
	align 4

OnCommand endp

;*** dialog proc to select a control

NotificationDlg proc uses __this thisarg, uMsg:DWORD, wParam:WPARAM, lParam:LPARAM

local	rect:RECT
local	wp:WINDOWPLACEMENT

;	mov __this, this@
;;	DebugOut "NotificationDlg, uMsg=%X, wParam=%X, lParam=%X", uMsg, wParam, lParam
	mov __this, g_pNotificationDlg

	mov eax,uMsg
	.if (eax == WM_INITDIALOG)

		invoke OnInitDialog
		.if (g_hFont)
			invoke SendMessage, m_hWndLV, WM_SETFONT, g_hFont, 0
		.endif
		mov eax,1

	.elseif (eax == WM_COMMAND)

		invoke OnCommand, wParam, lParam

	.elseif (eax == WM_NOTIFY)

		invoke OnNotify, lParam

	.elseif (eax == WM_SIZE)
		
		.if (wParam != SIZE_MINIMIZED)
			movzx eax, word ptr lParam+0
			movzx edx, word ptr lParam+2
			invoke SetWindowPos, m_hWndLV, NULL, 0, 0, eax, edx, SWP_NOMOVE or SWP_NOZORDER or SWP_NOACTIVATE
			invoke GetClientRect, m_hWndLV, addr rect
			invoke ListView_SetColumnWidth( m_hWndLV, 0, rect.right)
		.endif

	.elseif (eax == WM_CLOSE)

		mov wp.length_, sizeof WINDOWPLACEMENT
		invoke GetWindowPlacement, m_hWnd, addr wp
		invoke CopyRect, addr g_rect, addr wp.rcNormalPosition
		mov eax, g_rect.right
		sub eax, g_rect.left
		mov g_rect.right, eax
		mov eax, g_rect.bottom
		sub eax, g_rect.top
		mov g_rect.bottom, eax

		invoke Destroy@CLogWindow

	.elseif (eax == WM_ACTIVATE)

		movzx eax,word ptr wParam
		.if (eax == WA_INACTIVE)
			mov g_hWndDlg, NULL
		.else
			mov eax, m_hWnd
			mov g_hWndDlg, eax
		.endif

	.else

		xor eax,eax

	.endif
	ret
	align 4

NotificationDlg endp


;*** constructor


Create@CLogWindow proc public uses __this

	mov __this, g_pNotificationDlg
	.if (__this)
		invoke RestoreAndActivateWindow, m_hWnd
		return __this
	.endif

	invoke malloc, sizeof CLogWindow
	.if (eax == NULL)
		ret
	.endif
	mov __this,eax

	mov m_pDlgProc, NotificationDlg
	mov g_pNotificationDlg, __this

	mov m_bNewLine, TRUE

	invoke CreateDialogParam, g_hInstance, IDD_LOGWINDOW, NULL, classdialogproc, __this
	.if (eax)
if ?USEDEBUGOUT
		push edx
		invoke VirtualProtect, offset _imp__OutputDebugStringA@4, sizeof DWORD, PAGE_READWRITE, esp
		pop edx
		mov eax, AddLine
		xchg eax, _imp__OutputDebugStringA@4
		mov g_oldOutputDebugString, eax
endif
		invoke Show@CLogWindow, SW_SHOWNORMAL
	.endif

	invoke UpdateLogSwitch@CLogWindow

	return __this
	align 4

Create@CLogWindow endp

Destroy@CLogWindow proc public uses __this

;	mov __this,this@
	mov __this, g_pNotificationDlg
	.if (__this == NULL)
		ret
	.endif

	invoke DestroyWindow, m_hWnd

if ?USEDEBUGOUT
	mov eax, g_oldOutputDebugString
	mov _imp__OutputDebugStringA@4, eax
endif

	mov g_pNotificationDlg, NULL
	invoke UpdateLogSwitch@CLogWindow
	invoke free, __this

	ret
	align 4

Destroy@CLogWindow endp

;--- static method

Show@CLogWindow proc public uses __this dwFlags:DWORD

;	mov __this,this@
	mov __this,g_pNotificationDlg
	.if (__this == NULL)
		ret
	.endif
	invoke ShowWindow, m_hWnd, dwFlags
	ret
	align 4

Show@CLogWindow endp

;--- add a line to listbox

AddLine proc uses esi edi __this pStr:ptr BYTE

local	dwSize:DWORD
local	lvi:LVITEM

;	mov __this,this@
	mov __this, g_pNotificationDlg
	.if (__this == NULL)
		ret
	.endif

	mov lvi.mask_,LVIF_TEXT
	mov lvi.iSubItem,0
	mov eax, m_iLine
	mov lvi.iItem, eax
	invoke lstrlen, pStr
	add eax, 4
	and al, 0FCh
	mov dwSize, eax
	sub esp, eax
	mov lvi.pszText, esp
	mov edi, esp
	mov esi, pStr
	.while (1)
		lodsb
		.break .if ((al == 0) || (al == 13) || (al == 10))
		stosb
	.endw
	mov byte ptr [edi],0
	movzx edi, al
	.if (m_bNewLine)
		invoke ListView_InsertItem( m_hWndLV, addr lvi)
		invoke ListView_EnsureVisible( m_hWndLV, lvi.iItem, TRUE)
	.else
		mov eax, dwSize
		add eax, 128
		mov esi, esp
		sub esp, eax
		add dwSize, eax
		mov lvi.pszText, esp
		mov lvi.cchTextMax, 128
		invoke ListView_GetItem( m_hWndLV, addr lvi)
		mov edx, esp
		invoke lstrcat, edx, esi
		invoke ListView_SetItem( m_hWndLV, addr lvi)
	.endif
	.if (edi)
		mov m_bNewLine, TRUE
		inc m_iLine
	.else
		mov m_bNewLine, FALSE
	.endif
	add esp, dwSize
	ret

	align 4

AddLine endp

printf@CLogWindow proc c public pszFormat:LPSTR, varargs:VARARG

	.if (g_bLogActive)
		pushad
		sub esp, 1024
		mov edx, esp
		invoke wvsprintf, edx, pszFormat, addr varargs
		mov edx, esp
		.if (g_bLogToDebugWnd)
			.if (byte ptr [edx+eax-1] == 0Ah)
				mov dword ptr [edx+eax-1], 0A0Dh
			.endif
			invoke OutputDebugString, edx
		.else
			invoke AddLine, edx
		.endif
		add esp, 1024
		popad
	.endif
	ret
	align 4

printf@CLogWindow endp

UpdateLogSwitch@CLogWindow proc public
	.if ((g_pNotificationDlg) || (g_bLogToDebugWnd))
		mov g_bLogActive, TRUE
	.else
		mov g_bLogActive, FALSE
	.endif
	ret
UpdateLogSwitch@CLogWindow endp

	end
