	.386
	.model flat,stdcall
	option casemap:none
	option proc:private

	.nolist
	.nocref
WIN32_LEAN_AND_MEAN equ 1
	include windows.inc
	include commctrl.inc
	include richedit.inc
	include windowsx.inc

	include macros.inc
	include debugout.inc
	include CHexEdit.inc
	include rsrc.inc
	.list
	.cref

;--- externals used

malloc	proto :DWORD
free	proto :HANDLE
externdef g_hCsrWait:HCURSOR
externdef g_dwFontWidth:DWORD


MYWINDATA	equ GWL_USERDATA

?BOOKMARK	equ 0
?COLOR		equ 0
?FINDTEXT	equ 0
?VCENTER	equ 0
?LINENUMBER	equ 0
?STATE		equ 0
?USEGLOBAL	equ 0
?SIZECHARBUFF equ 16384

;include Data.inc

MAXCHARMEM			equ 128*1024
MAXSTREAM			equ 128*1024
MAXUNDOMEM			equ 32*1024
MINSIZE				equ 30	;minimum size of lower window

UNDO_CHARINSERT 	equ 1
UNDO_CHAROVERWRITE	equ 2
UNDO_INSERTBLOCK	equ 3
UNDO_DELETEBLOCK	equ 4

SBWT				equ 16	;initial value scrollbar width
BTNWT				equ 14	;linenumber button width
BTNHT				equ 6	;splitt button height
SELWT				equ 0	;	14
LNRWT				equ 28	;line number width
SPCWT				equ 3

HEUNDO struct
	rpPrev		dd ?		;Relative pointer to previous
	cp			dd ?		;Character position
	cb			dd ?		;Size in bytes
	fun 		db ?		;Function
HEUNDO ends

TIMER struct
	hwnd		HWND ?
	umsg		dd ?
	lparam		LPARAM ?
	wparam		WPARAM ?
TIMER ends

HEBRUSH struct
	hBrBck		HBRUSH ?	;Back color brush
	hBrSelBck	HBRUSH ?	;Sel back color brush
	hBrSelBar	HBRUSH ?	;Selection bar
	hPenSelbar	dd ?		;Selection bar pen
HEBRUSH ends

HEFONTINFO struct
	fntwt		dd ?		;Font width
	fntht		dd ?		;Font height
	linespace	dd ?		;Extra line spacing
HEFONTINFO ends

HEEDT struct
	hwnd		HWND ?		;Handle of edit a or b
	hvscroll	HWND ?		;Handle of scroll bar
	nline		dd ?		;Scroll position
	rc			RECT <?>	;Edit a or b rect
HEEDT ends

CHexEdit struct
	hwnd		HWND ?		;Handle of main window
	fstyle		dd ?		;Window style
	ID			dd ?		;Window ID
	hwndParent	HWND ?		;Handle of parent window
	edta		HEEDT <>
	edtb		HEEDT <>
	hhscroll	HWND ?		;Handle of horizontal scrollbar
	hgrip		HWND ?		;Handle of sizegrip
	hnogrip 	HWND ?		;Handle of nosizegrip
	hsbtn		HWND ?		;Handle of splitt button
if ?LINENUMBER
	hlin		HWND ?		;Handle of linenumber button
endif
if ?STATE
	hsta		HWND ?		;Handle of state window
endif
	htt 		HWND ?		;Handle of tooltip
	fresize 	dd ?		;Resize in action flag
	fsplitt 	dd ?		;Splitt factor
	nsplitt 	dd ?		;Splitt height

	hmem		LPVOID ?	;byte buffer
	nbytes		dd ?		;size of byte buffer
	nsize		dd ?		;max Size

	hundo		LPVOID ?
	rpundo		dd ?
	cbundo		dd ?

	rc			RECT <?>	;Main rect
	selbarwt	dd ?		;Width of selection bar
if ?LINENUMBER
	nlinenrwt	dd ?		;Initial width of linenumber bar
endif
	linenrwt	dd ?		;Width of linenumber bar
	cpMin		dd ?		;Selection min
	cpMax		dd ?		;Selection max
	cpx 		dd ?		;Scroll position
	hwndFocus	HWND ?		;Handle of edit having focus
	fOvr		BOOLEAN ?	;Insert / Overwrite
	fHasFocus	BOOLEAN ?	;control has focus
	fCaretHide	BOOLEAN ?	;Caret is hidden
	fChanged	BOOLEAN ?	;Content changed
	fHideSel	BOOLEAN ?	;Hide selection
	bEnabled	BOOLEAN ?	;window enabled
	filler		db 2 dup (?)
	clr 		HECOLOR <?>
	br			HEBRUSH <?>
	fnt 		HEFONT <?>
	fntinfo 	HEFONTINFO <?>
if ?BOOKMARK
	lpBmCB		dd ?		;Bookmark paint callback
endif
	nchange 	dd ?		;# of changes (Used by EN_SELCHANGE)
	nlastchange dd ?		;Used by EN_SELCHANGE
	addrxp		dd ?		;x pos address area
	addrwt		dd ?		;width address area
	dataxp		dd ?		;x pos data area
	datawt		dd ?		;width data area
	asciixp 	dd ?		;x pos ascii area
	asciiwt 	dd ?		;width ascii area
CHexEdit ends

__this	textequ <ebx>
_this	textequ <[__this].CHexEdit>

	MEMBER hwnd, fstyle, ID, hwndParent, edta, edtb, hhscroll
	MEMBER hgrip, hnogrip, hsbtn, htt, fresize, fsplitt, nsplitt
	MEMBER hmem, nbytes, nsize, hundo, rpundo, cbundo
	MEMBER rc, selbarwt, cpMax, cpMin, fOvr, cpx
	MEMBER hwndFocus, fHasFocus, fCaretHide, fChanged, fHideSel, bEnabled
	MEMBER clr, br, fnt, fntinfo, nchange, nlastchange 
	MEMBER addrxp, addrwt, dataxp, datawt, asciixp, asciiwt
	MEMBER linenrwt

SaveUndo			PROTO	:DWORD,:DWORD,:DWORD,:DWORD
ScrollCaret 		PROTO	:HWND
FindBookmark		PROTO	:DWORD,:DWORD
ResetMem			PROTO
InsertChars			proto ncp:DWORD, nChars:DWORD
DeleteChars			proto ncp:DWORD, nChars:DWORD
DeleteSelection		proto cpMin:DWORD, cpMax:DWORD
SelChange			proto nSelTyp:DWORD
EditCopy			proto
EditPaste			proto
EditSelectAll		proto hWin:HWND
InvalidateSelection proto dwIndex:DWORD, cpMin:DWORD, cpMax:DWORD


STYLE				equ WS_CHILD or WS_VISIBLE

	.const

szClassName 		db 'HEXEDIT',0
szEditClassName 	db 'HEXEDITCHILD',0
szToolTip			db 'HexEdit control',0

szScrollBar 		db 'ScrollBar',0
szStatic			db 'Static',0
if ?STATE
szButton			db 'Button',0
szChanged			db 'Changed indicator',0
endif
szToolTips			db 'Tooltips_class32',0

szSplitterBar		db 'Splitter Bar',0
if ?LINENUMBER
szLineNumber		db 'Show/Hide Linenumbers',0
endif
szX 				db '0123456789ABCDEF',0

szMemFailChar		db 'Memory allocation failed! (Char)',0
szMemFailUndo		db 'Memory allocation failed! (Undo)',0

	.data

hInstance			HINSTANCE NULL
TimerID 			dd 0
hex 				dd 'A'-10
g_hFont				HFONT NULL
g_hMenu				HMENU NULL
g_sbwt				dd SBWT
g_OldFontWidth		dd 0
g_pCharBuff			LPVOID NULL
g_dwMyCBFormat		DD NULL

	.data?

g_hHSCur			HCURSOR ?
if ?STATE
OldStateProc		dd ?
endif
OldSplittBtnProc	dd ?
tmr1				TIMER <?>
tmr2				TIMER <?>
fSize				dd ?
fSelState			dd ?


	.code

ShowContextMenu proc uses esi

local	pt:POINT
local	hPopupMenu:HMENU

		.if (g_hMenu)
			invoke GetSubMenu, g_hMenu, 0
			.if (eax != 0)
				mov hPopupMenu, eax

				mov esi, MF_GRAYED or MF_DISABLED
				.if (m_rpundo)
					mov esi, MF_ENABLED
				.endif
				invoke EnableMenuItem, hPopupMenu, IDM_UNDO, esi

				mov esi, MF_GRAYED or MF_DISABLED
				invoke SendMessage, m_hwnd, EM_CANREDO, 0, 0
				.if (eax)
					mov esi, MF_ENABLED
				.endif
				invoke EnableMenuItem, hPopupMenu, IDM_REDO, esi

				mov eax, m_cpMin
				mov esi, MF_GRAYED or MF_DISABLED
				.if (eax != m_cpMax)
					mov esi, MF_ENABLED
				.endif
				invoke EnableMenuItem, hPopupMenu, IDM_CUT, esi
				invoke EnableMenuItem, hPopupMenu, IDM_COPY, esi
				invoke EnableMenuItem, hPopupMenu, IDM_DELETE, esi

				mov esi, MF_GRAYED or MF_DISABLED
				invoke IsClipboardFormatAvailable, CF_TEXT
				.if eax
					mov esi, MF_ENABLED
				.else
					invoke IsClipboardFormatAvailable, g_dwMyCBFormat
					.if eax
						mov esi, MF_ENABLED
					.endif
				.endif
				invoke EnableMenuItem, hPopupMenu, IDM_PASTE, esi
				invoke GetCursorPos,addr pt
				invoke TrackPopupMenu, hPopupMenu,\
					TPM_LEFTALIGN or TPM_LEFTBUTTON or TPM_RETURNCMD,\
					pt.x, pt.y, 0, m_hwnd, NULL
				xor ecx, ecx
				.if (eax == IDM_CUT)
					mov ecx, WM_CUT
				.elseif (eax == IDM_COPY)
					mov ecx, WM_COPY
				.elseif (eax == IDM_PASTE)
					mov ecx, WM_PASTE
				.elseif (eax == IDM_DELETE)
					mov ecx, WM_CLEAR
				.elseif (eax == IDM_SELECTALL)
					invoke EditSelectAll, m_hwndFocus
					xor ecx, ecx
				.elseif (eax == IDM_UNDO)
					mov ecx, EM_UNDO
				.elseif (eax == IDM_REDO)
					mov ecx, EM_REDO
				.endif
				.if (ecx)
					invoke SendMessage, m_hwnd, ecx, 0, 0
				.endif
			.endif
		.endif
		ret
		align 4

ShowContextMenu endp


;include Caret.asm


SetCaret proc uses ebx esi hWin:HWND
	LOCAL	xp:DWORD
	LOCAL	yp:DWORD
	LOCAL	buffer[32]:BYTE

	.if hWin
		invoke GetWindowLong,hWin,MYWINDATA
		mov ebx,eax
		mov eax,hWin
		.if eax==m_edta.hwnd
			lea esi,m_edta
		.else
			lea esi,m_edtb
		.endif
		invoke GetFocus
		.if eax==hWin
			.if m_fOvr
				mov edx,m_fntinfo.fntwt
			.else
				mov edx,2
			.endif
			invoke CreateCaret,hWin,NULL,edx,m_fntinfo.fntht
			mov eax,m_cpMin
			mov ecx,eax
			shr eax,1
			and eax,0Fh
			mov edx,3
			mul edx
			sub eax,m_cpx
			mov edx,m_fntinfo.fntwt
			mul edx
			add eax,m_linenrwt
			add eax,m_selbarwt
			add eax,m_dataxp
			test ecx,1
			.if !ZERO?
				add eax,m_fntinfo.fntwt
			.endif
			mov xp,eax
			mov eax,ecx
			shr eax,5
			sub eax,[esi].HEEDT.nline
			mov edx,m_fntinfo.fntht
			mul edx
			mov yp,eax
			invoke SetCaretPos,xp,yp
			mov eax,m_linenrwt
			add eax,m_selbarwt
			mov edx,[esi].HEEDT.rc.right
			sub edx,m_fntinfo.fntwt
			.if sdword ptr eax>xp || sdword ptr edx<xp
				.if !m_fCaretHide
					invoke HideCaret,hWin
					mov m_fCaretHide,TRUE
				.endif
			.else
				mov eax,m_cpMin
				.if eax!=m_cpMax
					.if !m_fCaretHide
						invoke HideCaret,hWin
						mov m_fCaretHide,TRUE
					.endif
				.else
					invoke ShowCaret,hWin
					mov m_fCaretHide,FALSE
				.endif
			.endif
		.endif
	.endif
	ret
	align 4

SetCaret endp

TestScrollX proc uses esi,hWin:HWND

	LOCAL xp:DWORD

	mov eax,hWin
	.if eax==m_edta.hwnd
		lea esi,m_edta
	.else
		lea esi,m_edtb
	.endif
	mov eax,m_cpMin
	mov ecx,eax
	shr eax,1
	and eax,0Fh
	mov edx,3
	mul edx
	sub eax,m_cpx
	mov edx,m_fntinfo.fntwt
	mul edx
	add eax,m_linenrwt
	add eax,m_selbarwt
	add eax,m_dataxp
	test ecx,1
	.if !ZERO?
		add eax,m_fntinfo.fntwt
	.endif
	mov xp,eax
	mov eax,m_linenrwt
	add eax,m_selbarwt
	.if sdword ptr eax > xp
		sub eax,xp
		mov ecx,m_fntinfo.fntwt
		xor edx,edx
		div ecx
		sub eax,m_cpx
		neg eax
		sub eax,4
	.else
		mov eax,[esi].HEEDT.rc.right
		sub eax,m_fntinfo.fntwt
		.if sdword ptr eax < xp
			sub eax,xp
			neg eax
			mov ecx,m_fntinfo.fntwt
			xor edx,edx
			div ecx
			add eax,m_cpx
			add eax,4
		.else
			mov eax,m_cpx
		.endif
	.endif
	ret
	align 4

TestScrollX endp

TestScrollY proc uses esi,hWin:HWND
	LOCAL sinf:SCROLLINFO

	mov eax,hWin
	.if eax==m_edta.hwnd
		lea esi,m_edta
	.else
		lea esi,m_edtb
	.endif
	mov sinf.cbSize,sizeof sinf
	mov sinf.fMask,SIF_ALL
	invoke GetScrollInfo,[esi].HEEDT.hvscroll,SB_CTL,addr sinf
	mov eax,m_cpMin
	mov edx,[esi].HEEDT.nline
	mov ecx,edx
	add ecx,sinf.nPage
	dec ecx
	shr eax,5
	.if eax < edx
		sub eax,edx
	.elseif eax > ecx
		sub eax,ecx
	.else
		xor eax,eax
	.endif
	add eax,[esi].HEEDT.nline
	ret
	align 4

TestScrollY endp

DoScroll proc uses ebx esi edi,hWin:HWND,x:DWORD,y:DWORD
	LOCAL sinf:SCROLLINFO
	LOCAL rect:RECT

	invoke GetWindowLong,hWin,MYWINDATA
	mov ebx,eax
	mov eax,hWin
	.if eax==m_edta.hwnd
		lea esi,m_edta
	.else
		lea esi,m_edtb
	.endif
	mov sinf.cbSize,sizeof sinf
	mov sinf.fMask,SIF_ALL or SIF_DISABLENOSCROLL
	invoke GetScrollInfo,[esi].HEEDT.hvscroll,SB_CTL,addr sinf
	mov eax,y
	.if eax!=sinf.nPos
		mov sinf.nPos,eax
		mov sinf.nMin,0
		mov eax,m_nbytes
		shr eax,4
		inc eax
		mov ecx,m_fntinfo.fntht
		mul ecx
		.if eax < [esi].HEEDT.rc.bottom
			mov eax,[esi].HEEDT.rc.bottom
		.endif
		xor edx,edx
		div ecx
		dec eax
		mov sinf.nMax,eax
		mov eax,[esi].HEEDT.rc.bottom
		xor edx,edx
		div ecx
		mov sinf.nPage,eax
		invoke SetScrollInfo,[esi].HEEDT.hvscroll,SB_CTL,addr sinf,TRUE
		invoke GetScrollInfo,[esi].HEEDT.hvscroll,SB_CTL,addr sinf
		mov eax,sinf.nPos
		mov ecx,[esi].HEEDT.nline
		.if eax!=ecx
			mov [esi].HEEDT.nline,eax
			sub eax,ecx
			neg eax
			mov ecx,m_fntinfo.fntht
			mul ecx
			mov ecx,eax
			invoke ScrollWindow,hWin,0,ecx,addr [esi].HEEDT.rc,NULL
			invoke UpdateWindow,hWin
			invoke SetCaret,hWin
		.endif
	.endif
	mov sinf.cbSize,sizeof sinf
	mov sinf.fMask,SIF_ALL or SIF_DISABLENOSCROLL
	invoke GetScrollInfo,m_hhscroll,SB_CTL,addr sinf
	mov eax,x
	.if eax!=sinf.nPos
		mov sinf.nPos,eax
		mov sinf.nMin,0
		mov eax,[esi].HEEDT.rc.right
		sub eax,m_selbarwt
		sub eax,m_linenrwt
		mov ecx,m_fntinfo.fntwt
		xor edx,edx
		div ecx
		mov sinf.nPage,eax
		mov eax,m_asciixp
		add eax,m_asciiwt
		mov ecx,m_fntinfo.fntwt
		xor edx,edx
		div ecx
		mov sinf.nMax,eax
		invoke SetScrollInfo,m_hhscroll,SB_CTL,addr sinf,TRUE
		invoke GetScrollInfo,m_hhscroll,SB_CTL,addr sinf
		mov eax,sinf.nPos
		mov ecx,m_cpx
		.if eax!=ecx
			mov m_cpx,eax
			sub eax,ecx
			neg eax
			mov ecx,m_fntinfo.fntwt
			mul ecx
			mov edi,eax
			invoke CopyRect,addr rect,addr m_edta.rc
			mov eax,m_linenrwt
			add eax,m_selbarwt
			add eax,SPCWT
			mov rect.left,eax
			invoke ScrollWindow,m_edta.hwnd,edi,0,addr rect,addr rect
			invoke CopyRect,addr rect,addr m_edtb.rc
			mov eax,m_linenrwt
			add eax,m_selbarwt
			add eax,SPCWT
			mov rect.left,eax
			invoke ScrollWindow,m_edtb.hwnd,edi,0,addr rect,addr rect
			invoke UpdateWindow,hWin
			invoke SetCaret,hWin
		.endif
	.endif
	ret
	align 4

DoScroll endp

GetCharFromPos proc uses esi hWin:HWND, x:DWORD, y:DWORD

	LOCAL	cp:DWORD

	mov 	eax,hWin
	.if eax==m_edta.hwnd
		lea 	esi,m_edta
	.else
		lea 	esi,m_edtb
	.endif
	mov 	eax,y
	.if sdword ptr eax<0
		xor 	eax,eax
	.elseif eax>[esi].HEEDT.rc.bottom
		mov 	eax,[esi].HEEDT.rc.bottom
	.endif
	mov 	ecx,m_fntinfo.fntht
	xor 	edx,edx
	div 	ecx
	add 	eax,[esi].HEEDT.nline
	shl 	eax,4
	.if eax>m_nbytes
		mov 	eax,m_nbytes
		and 	eax,0FFFFFFF0h
	.endif
	shl eax,1
	mov cp,eax
	mov eax,m_cpx
	mov ecx,m_fntinfo.fntwt
	mul ecx
	mov edx,eax
	mov eax,x
	sub eax,m_dataxp
	sub eax,m_linenrwt
	sub eax,m_selbarwt
	add eax,edx
	.if sdword ptr eax < 0
		xor eax,eax
	.endif
	mov ecx,m_fntinfo.fntwt
	mov edx,ecx
	add ecx,ecx
	add ecx,edx
	shl eax,1
	add eax,edx
	xor edx,edx
	div ecx
	.if eax > 31
		mov eax,31
	.endif
	add eax,cp
	mov ecx,m_nbytes
	shl ecx,1
	.if eax > ecx
		mov eax,ecx
	.endif
	ret
	align 4

GetCharFromPos endp

ScrollCaret proc uses esi hWin:HWND

	.if hWin
		mov eax,hWin
		.if eax == m_edta.hwnd
			lea esi,m_edta
		.else
			lea esi,m_edtb
		.endif
		invoke TestScrollY, hWin
		.if eax != [esi].HEEDT.nline
			invoke DoScroll, hWin, m_cpx,eax
		.endif
		invoke TestScrollX, hWin
		.if eax != m_cpx
			invoke DoScroll,hWin,eax,[esi].HEEDT.nline
		.endif
		invoke SetCaret,hWin
	.endif
	ret
	align 4

ScrollCaret endp


;include Misc.asm


SetFont proc uses edi lpRafont:ptr HEFONT
	LOCAL	hDC:HDC
	LOCAL	dtp:DRAWTEXTPARAMS
	LOCAL	rect:RECT
	LOCAL	pt:POINT
	LOCAL	tm:TEXTMETRIC

	mov 	edx,lpRafont
	mov 	eax,[edx].HEFONT.hFont
	mov 	m_fnt.hFont,eax
	mov 	eax,[edx].HEFONT.hLnrFont
	mov 	m_fnt.hLnrFont,eax
	invoke GetDC,m_hwnd
	mov 	hDC,eax
	invoke SelectObject,hDC,m_fnt.hFont
	push	eax
	;Get height & width
	invoke GetTextExtentPoint32,hDC,addr szX,16,addr pt
	mov 	eax,pt.x
	shr 	eax,4
	mov 	pt.x,eax
	mov 	m_fntinfo.fntwt,eax
	mov 	eax,pt.y
	add 	eax,m_fntinfo.linespace
	mov 	m_fntinfo.fntht,eax
	pop 	eax
	invoke SelectObject,hDC,eax
	invoke ReleaseDC, m_hwnd, hDC
	test	m_fstyle,STYLE_NOADDRESS
	.if ZERO?
		mov 	m_addrxp,SPCWT
		mov 	eax,pt.x
		shl 	eax,3
		mov 	m_addrwt,eax
		add 	eax,SPCWT*2+1
		add 	eax,m_addrxp
	.else
		mov 	m_addrxp,0
		mov 	m_addrwt,0
		mov 	eax,SPCWT
	.endif
	mov 	m_dataxp,eax
	mov 	eax,pt.x
	shl 	eax,4
	mov 	ecx,eax
	shl 	eax,1
	add 	eax,ecx
	sub 	eax,pt.x
	mov 	m_datawt,eax
	test	m_fstyle,STYLE_NOASCII
	.if ZERO?
		add 	eax,SPCWT*2+1
		add 	eax,m_dataxp
		mov 	m_asciixp,eax
		mov 	eax,pt.x
		shl 	eax,4
		mov 	m_asciiwt,eax
	.else
		mov 	eax,m_dataxp
		add 	eax,m_datawt
		mov 	m_asciixp,eax
		mov 	m_asciiwt,0
	.endif
	ret
	align 4

SetFont endp

DestroyBrushes proc

	mov eax,m_br.hBrBck
	.if eax
		invoke DeleteObject,eax
	.endif
	mov eax,m_br.hBrSelBck
	.if eax
		invoke DeleteObject,eax
	.endif
	mov eax,m_br.hBrSelBar
	.if eax
		invoke DeleteObject,eax
	.endif
	mov eax,m_br.hPenSelbar
	.if eax
		invoke DeleteObject,eax
	.endif
	ret
	align 4

DestroyBrushes endp


CreateBrushes proc

	invoke DestroyBrushes
	invoke CreateSolidBrush,m_clr.bckcol
	mov m_br.hBrBck,eax
	invoke CreateSolidBrush,m_clr.selbckcol
	mov m_br.hBrSelBck,eax
	invoke CreateSolidBrush,m_clr.selbarbck
	mov m_br.hBrSelBar,eax
	invoke CreatePen,PS_SOLID,1,m_clr.selbarpen
	mov m_br.hPenSelbar,eax
	ret
	align 4

CreateBrushes endp

DwToAscii proc uses ebx esi edi dwVal:DWORD,lpAscii:DWORD

	mov eax,dwVal
	mov edi,lpAscii
	or eax,eax
	jns pos
	mov byte ptr [edi],'-'
	neg eax
	inc edi
  pos:
	mov ecx,429496730
	mov esi,edi
  @@:
	mov ebx,eax
	mul ecx
	mov eax,edx
	lea edx,[edx*4+edx]
	add edx,edx
	sub ebx,edx
	add bl,'0'
	mov [edi],bl
	inc edi
	or eax,eax
	jne @b
	mov byte ptr [edi],al
	.while esi < edi
		dec edi
		mov al,[esi]
		mov ah,[edi]
		mov [edi],al
		mov [esi],ah
		inc esi
	.endw
	ret
	align 4

DwToAscii endp


ConvertHexString proc uses esi edi lpStr:LPSTR, pCharBuff:LPVOID

	mov 	esi,lpStr
	mov 	edi,pCharBuff
  NxtByte:
	mov 	al,[esi]
	.if al
		inc 	esi
		.if (al>='0' && al<='9') || (al>='A' && al<='F') || (al>='a' && al<='f')
			.if al>='a'
				and 	al,5Fh
			.endif
			.if al>='A'
				sub 	al,'A'-10
			.else
				sub 	al,'0'
			.endif
			shl 	al,4
			mov 	ah,al
		.else
			jmp 	NxtByte
		.endif
	  NxtNybble:
		mov 	al,[esi]
		.if al
			inc 	esi
			.if (al>='0' && al<='9') || (al>='A' && al<='F') || (al>='a' && al<='f')
				.if al>='a'
					and 	al,5Fh
				.endif
				.if al>='A'
					sub 	al,'A'-10
				.else
					sub 	al,'0'
				.endif
				or		ah,al
			.else
				jmp 	NxtNybble
			.endif
		.endif
		mov 	[edi],ah
		inc 	edi
		mov		ecx, pCharBuff
		add		ecx, ?SIZECHARBUFF
		cmp 	edi, ecx
		jne 	NxtByte
	.endif
	mov 	eax, edi
	sub 	eax, pCharBuff
	mov 	edx, esi
	sub 	edx, lpStr
	ret
	align 4

ConvertHexString endp


SelChange proc nSelTyp:DWORD
	LOCAL	hesel:HESELCHANGE

	mov 	eax,m_hwnd
	mov 	hesel.nmhdr.hwndFrom,eax
	mov 	eax,m_ID
	mov 	hesel.nmhdr.idFrom,eax
	mov 	hesel.nmhdr.code,EN_SELCHANGE

	mov 	eax,m_cpMin
	mov 	hesel.chrg.cpMin,eax
	mov 	eax,m_cpMax
	mov 	hesel.chrg.cpMax,eax

	mov 	eax,nSelTyp
	mov 	hesel.seltyp,ax
	mov 	eax,m_cpMin
	shr 	eax,5
	mov 	hesel.line,eax
	mov 	eax,m_nbytes
	shr 	eax,4
	mov 	hesel.nlines,eax

	movzx 	eax,m_fChanged
	mov 	hesel.fchanged,eax
	invoke SendMessage,m_hwndParent,WM_NOTIFY,m_ID,addr hesel
	ret
	align 4

SelChange endp


;include Edit.asm


GetChar proc uses edi ncp:DWORD

	xor 	eax,eax
	mov 	edi,ncp
	.if edi<m_nbytes
if ?USEGLOBAL
		invoke GlobalLock,m_hmem
else
		mov		eax, m_hmem
endif
		movzx	edi,byte ptr [edi+eax]
if ?USEGLOBAL
		invoke GlobalUnlock,m_hmem
endif
		mov 	eax,edi
	.endif
	ret
	align 4

GetChar endp

;--- add some bytes to byte buffer

PutChar proc uses esi edi ncp:DWORD, lpChar:LPVOID, nChars:DWORD

	mov 	edi,ncp
	mov 	eax, m_nbytes
	add 	eax, nChars
	.if eax>=m_nsize
		and 	eax,-1 xor (MAXCHARMEM-1)
		add 	eax,MAXCHARMEM
		mov 	m_nsize,eax
if ?USEGLOBAL
		invoke GlobalReAlloc,m_hmem,eax,GMEM_MOVEABLE or GMEM_ZEROINIT
else
		push eax
		invoke GetProcessHeap
		pop ecx
		invoke HeapReAlloc, eax, NULL, m_hmem, ecx
endif
		.if !eax
			invoke MessageBox,NULL,addr szMemFailChar,addr szToolTip,MB_OK
			mov 	eax,m_hmem
		.endif
		mov 	m_hmem,eax
	.endif
if ?USEGLOBAL
	invoke GlobalLock,m_hmem
else
	mov eax, m_hmem
endif
	add 	edi,eax
	mov 	esi,lpChar
	mov 	ecx,nChars
	mov		dl, cl
	shr		ecx, 2
	rep		movsd
	mov		cl, dl
	and		cl, 3
	rep		movsb
if ?USEGLOBAL
	invoke GlobalUnlock,m_hmem
endif
	ret
	align 4

PutChar endp

;--- private method

InsertChars proc uses esi edi ncp:DWORD, nChars:DWORD
	LOCAL	lpMem:DWORD

	mov 	edi,ncp
	.if edi>m_nbytes
		mov 	edi,m_nbytes
		mov 	ncp,edi
	.endif
	mov 	eax,m_nbytes
	add 	eax,nChars
	.if eax>=m_nsize
		and 	eax,-1 xor (MAXCHARMEM-1)
		add 	eax,MAXCHARMEM
		mov 	m_nsize,eax
if ?USEGLOBAL
		invoke GlobalReAlloc, m_hmem, eax, GMEM_MOVEABLE or GMEM_ZEROINIT
else
		push eax
		invoke GetProcessHeap
		pop ecx
		invoke HeapReAlloc, eax, NULL, m_hmem, ecx
endif
		.if !eax
			invoke MessageBox,NULL,addr szMemFailChar,addr szToolTip,MB_OK
			mov 	eax,m_hmem
		.endif
		mov 	m_hmem,eax
	.endif
if ?USEGLOBAL
	invoke GlobalLock,m_hmem
else
	mov		eax, m_hmem
endif
	mov 	lpMem,eax
	mov 	esi,eax
	add 	esi,m_nbytes
	mov 	edi,esi
	add 	edi,nChars
	mov 	ecx,m_nbytes
	inc 	ecx
	sub 	ecx,ncp
	std
	rep movsb
	cld
	mov 	edi,ncp
	add 	edi,lpMem
	mov 	ecx,nChars
	xor 	eax,eax
	rep stosb
if ?USEGLOBAL
	invoke GlobalUnlock,m_hmem
endif
	mov 	eax,nChars
	add 	m_nbytes,eax
	ret
	align 4

InsertChars endp

DeleteChars proc uses esi edi ncp:DWORD, nChars:DWORD
	LOCAL	lpMem:DWORD

	mov 	eax,ncp
	.if eax>m_nbytes
		mov 	eax,m_nbytes
		mov 	ncp,eax
	.endif
	add 	eax,nChars
	.if eax>m_nbytes
		mov 	eax,m_nbytes
		sub 	eax,ncp
		mov 	nChars,eax
	.endif
	mov 	edi,ncp
if ?USEGLOBAL
	invoke GlobalLock,m_hmem
else
	mov		eax, m_hmem
endif
	mov 	lpMem,eax
	add 	edi,eax
	mov 	esi,edi
	add 	esi,nChars
	mov 	ecx,m_nbytes
	add 	ecx,lpMem
	sub 	ecx,esi
	rep movsb
if ?USEGLOBAL
	invoke GlobalUnlock,m_hmem
endif
	mov 	eax,nChars
	sub 	m_nbytes,eax
	ret
	align 4

DeleteChars endp

DeleteSelection proc cpMin:DWORD, cpMax:DWORD

	mov 	eax,cpMin
	mov 	ecx,cpMax
	.if eax>ecx
		xchg	eax,ecx
	.endif
	mov 	cpMin,eax
	mov 	cpMax,ecx
	shr 	eax,1
	shr 	ecx,1
	sub 	ecx,eax
	shl 	eax,1
	invoke SaveUndo, UNDO_DELETEBLOCK,eax,0,ecx
	mov 	eax,cpMin
	mov 	ecx,cpMax
	shr 	eax,1
	shr 	ecx,1
	sub 	ecx,eax
	.if ecx
		invoke DeleteChars, eax, ecx
		mov 	eax,cpMin
		mov 	m_cpMin,eax
		mov 	m_cpMax,eax
		invoke InvalidateRect,m_edta.hwnd,NULL,FALSE
		invoke InvalidateRect,m_edtb.hwnd,NULL,FALSE
		mov 	m_fChanged,TRUE
		inc 	m_nchange
		invoke ScrollCaret, m_hwndFocus
if ?STATE
		invoke InvalidateRect,m_hsta,NULL,TRUE
endif
		invoke SetCaret,m_hwndFocus
		invoke SelChange, SEL_TEXT
		push	word ptr EN_CHANGE
		push	word ptr m_ID
		pop		ecx
		invoke SendMessage, m_hwndParent, WM_COMMAND, ecx, m_hwnd
	.endif
	ret
	align 4

DeleteSelection endp

InsertBlock proc ncp:DWORD,lpStr:DWORD,nBytes:DWORD

	.if nBytes
		invoke InsertChars, ncp, nBytes
		invoke PutChar, ncp, lpStr, nBytes
		mov 	edx,ncp
		shl 	edx,1
		invoke SaveUndo, UNDO_INSERTBLOCK, edx, 0, nBytes
	.endif
	ret
	align 4

InsertBlock endp

;--- returns size in EAX

InsertHexString proc uses esi ncp:DWORD,lpStr:DWORD

	.if (!g_pCharBuff)
		invoke malloc, ?SIZECHARBUFF
		mov g_pCharBuff, eax
	.endif
	mov 	esi,lpStr
	push	ncp
	.while (1)
		invoke ConvertHexString, esi, g_pCharBuff
		.break .if (!edx)
		add 	esi,edx
		push	eax
		invoke InsertChars, ncp, eax
		pop 	eax
		push	eax
		invoke PutChar, ncp, g_pCharBuff, eax
		pop 	eax
		add 	ncp,eax
	.endw
	pop 	eax
	push	eax
	mov 	edx,eax
	sub 	eax,ncp
	neg 	eax
	shl 	edx,1
	invoke SaveUndo, UNDO_INSERTBLOCK, edx, 0, eax
	pop 	eax
	sub 	eax,ncp
	neg 	eax
	ret
	align 4

InsertHexString endp

InsertAsciiString proc ncp:DWORD,lpStr:DWORD

	invoke lstrlen,lpStr
	push eax
	invoke InsertBlock, ncp, lpStr, eax
	pop eax
	ret
	align 4

InsertAsciiString endp

StreamIn proc uses esi edi lParam:DWORD

	LOCAL	dwRead:DWORD
	LOCAL	hCMem:LPVOID

	invoke ResetMem
	invoke malloc, MAXSTREAM
	mov hCMem,eax
	xor edi,edi
	.repeat
		mov esi,lParam
		mov [esi].EDITSTREAM.dwError,0
if 0
		lea eax,dwRead
		push eax
		push MAXSTREAM
		push hCMem
		push [esi].EDITSTREAM.dwCookie
		call [esi].EDITSTREAM.pfnCallback
else
		invoke [esi].EDITSTREAM.pfnCallback, [esi].EDITSTREAM.dwCookie, hCMem, MAXSTREAM, addr dwRead
endif
		.break .if (eax)
		.if (dwRead)
			mov esi,hCMem
			invoke PutChar, edi, esi, dwRead
			add edi,dwRead
			mov m_nbytes,edi
		.endif
	.until (dwRead != MAXSTREAM)
	invoke free,hCMem
	ret
	align 4

StreamIn endp

StreamOut proc uses esi edi lParam:DWORD

	LOCAL	dwWrite:DWORD
	LOCAL	hCMem:DWORD

	invoke malloc, MAXSTREAM
	mov hCMem,eax
	xor esi,esi
	.repeat
		call FillCMem
		.break .if (!ecx)
		mov edx,lParam
		mov [edx].EDITSTREAM.dwError,0
if 0
		lea eax,dwWrite
		push eax
		push ecx
		push hCMem
		push [edx].EDITSTREAM.dwCookie
		call [edx].EDITSTREAM.pfnCallback
else
		invoke [edx].EDITSTREAM.pfnCallback, [edx].EDITSTREAM.dwCookie, hCMem, ecx, addr dwWrite
endif
	.until (eax)
	invoke free, hCMem
	ret

FillCMem:
ifdef @StackBase
@StackBase = @StackBase + 4
endif
	mov edi,hCMem
	mov ecx, m_nbytes
	sub ecx, esi
	.if (ecx > MAXSTREAM)
		mov ecx, MAXSTREAM
	.elseif (!ecx)
		retn
	.endif
if ?USEGLOBAL
	push ecx
	invoke GlobalLock,m_hmem
	pop ecx
else
	mov eax, m_hmem
endif
	push esi
	mov esi,eax
	add esi,[esp]
	push ecx
	mov dl, cl
	shr ecx,2
	rep movsd
	mov cl, dl
	and cl, 3
	rep movsb
if ?USEGLOBAL
	invoke GlobalUnlock,m_hmem		;changed to GlobalUnlock
endif
	pop ecx
	pop esi
	add esi, ecx
	retn
ifdef @StackBase
@StackBase = @StackBase - 4
endif
	align 4

StreamOut endp


;include Paint.asm


HexAddr:
	push ecx
	mov ecx,8
  @@:
	call Nybble
	mov [edi],dl
	inc edi
	dec ecx
	jne @b
	mov [edi],cl
	pop ecx
	retn

HexByte:
	mov al,[esi+ebx]
	shl eax,24
	call Nybble
	mov [edi],dl
	inc edi
	call Nybble
	mov [edi],dl
	inc edi
	inc ecx
	inc ebx
	.if ecx==8
		mov word ptr [edi],'-'
	.else
		mov word ptr [edi],' '
	.endif
	inc edi
	retn

Nybble:
	rol eax,4
	push eax
	and eax,0Fh
	.if eax<=9
		add eax,'0'
	.else
		add eax,hex
	.endif
	mov edx,eax
	pop eax
	retn

Ascii:
	.if ecx<16
		push ecx
		sub ecx,16
		neg ecx
		mov eax,ecx
		shl ecx,1
		add ecx,eax
		mov al,' '
		rep stosb
		pop ecx
	.endif
	sub ebx,ecx
	push ecx
	.while ecx
		mov al,[esi+ebx]
		.if al<20h || al>7Eh
			mov al,'.'
		.endif
		mov [edi],al
		inc edi
		inc ebx
		dec ecx
	.endw
	pop ecx
	sub ecx,16
	neg ecx
	mov al,' '
	rep stosb
	retn

HexLine proc uses ebx esi edi,lpMem:DWORD,nBytes:DWORD,nLine:DWORD,lpString:DWORD

	mov 	esi,lpMem
	mov 	edi,lpString
	mov 	byte ptr [edi],0
	mov 	ebx,nLine
	shl 	ebx,4
	.if ebx<=nBytes
		mov 	eax,ebx
		call	HexAddr
		xor 	ecx,ecx
		.while ebx<nBytes && ecx<16
			call	HexByte
		.endw
		call	Ascii
		xor 	eax,eax
		inc 	eax
	.else
		xor 	eax,eax
	.endif
	ret
	align 4

HexLine endp

ifdef @StackBase
	option stackbase:ebp
endif

OnPaint proc uses esi edi, hWin:HWND

	LOCAL	ps:PAINTSTRUCT
	LOCAL	rect:RECT
	LOCAL	lpMem:DWORD
	LOCAL	buffer[256]:BYTE
	LOCAL	cpx:DWORD
	LOCAL	hRgn:DWORD
	LOCAL	nSt:DWORD
	LOCAL	nEn:DWORD
	LOCAL	nLeft:DWORD
	LOCAL	nHilite:DWORD
	LOCAL	nRight:DWORD

	mov eax,hWin
	.if eax==m_edta.hwnd
		lea esi,m_edta
	.else
		lea esi,m_edtb
	.endif

if ?STATE
	mov eax,m_linenrwt
	add eax,m_selbarwt
	mov edx,SW_SHOWNA
	.if !eax
		mov edx,SW_HIDE
	.endif
	invoke ShowWindow,m_hsta,edx
endif
	test m_fstyle,STYLE_NOUPPERCASE
	.if ZERO?
		mov hex,'A'-10
	.else
		mov hex,'a'-10
	.endif
	mov eax,m_cpMin
	mov ecx,m_cpMax
	shr eax,1
	shr ecx,1
	.if eax>ecx
		xchg eax,ecx
	.endif
	mov nSt,eax
	mov nEn,ecx
	mov eax,m_cpx
	mov ecx,m_fntinfo.fntwt
	mul ecx
	mov cpx,eax
	invoke GetClientRect,hWin,addr rect
	invoke BeginPaint,hWin,addr ps
	invoke SelectObject,ps.hdc,m_fnt.hFont
	push eax
	invoke SelectObject,ps.hdc,m_br.hPenSelbar
	push eax
	invoke SetBkMode,ps.hdc,TRANSPARENT
if ?USEGLOBAL
	invoke GlobalLock,m_hmem
else
	mov eax, m_hmem
endif
	mov lpMem,eax
	mov eax,m_linenrwt
	add eax,m_selbarwt
	.if !eax
		dec eax
	.endif
	add eax,SPCWT
	mov rect.left,eax
	mov ecx, g_sbwt
	sub rect.right,ecx
	invoke CreateRectRgn,rect.left,rect.top,rect.right,rect.bottom
	mov hRgn,eax
	sub rect.left,SPCWT-1
	xor edx,edx
	mov edi,[esi].HEEDT.nline
	.while edx<rect.bottom
		push edx
		mov eax,rect.top
		.if eax<ps.rcPaint.bottom
			add eax,m_fntinfo.fntht
			.if eax>ps.rcPaint.top
				push rect.bottom
				mov rect.bottom,eax
				.if (m_bEnabled)
					mov eax, m_br.hBrBck
				.else
					invoke GetSysColorBrush, COLOR_BTNFACE
				.endif
				invoke FillRect, ps.hdc, addr rect, eax
				invoke SelectClipRgn,ps.hdc,hRgn
				push	edi
				test	m_fstyle,STYLE_NOADDRESS
				.if ZERO?
					mov 	edi,m_dataxp
					add 	edi,m_linenrwt
					add 	edi,m_selbarwt
					sub 	edi,cpx
					sub 	edi,SPCWT+1
					invoke MoveToEx,ps.hdc,edi,0,NULL
					invoke LineTo,ps.hdc,edi,rect.bottom
				.endif
				test	m_fstyle,STYLE_NOASCII
				.if ZERO?
					mov 	edi,m_asciixp
					add 	edi,m_linenrwt
					add 	edi,m_selbarwt
					sub 	edi,cpx
					sub 	edi,SPCWT+1
					invoke MoveToEx,ps.hdc,edi,0,NULL
					invoke LineTo,ps.hdc,edi,rect.bottom
				.endif
				pop edi
				invoke HexLine,lpMem,m_nbytes,edi,addr buffer
				mov esi,eax
				.if eax
					call	DrawSelBack
					invoke SelectObject,ps.hdc,m_fnt.hFont
					test	m_fstyle,STYLE_NOADDRESS
					.if ZERO?
;---- dont show an address if we are disabled and buffer is empty
						.if (m_bEnabled || m_nbytes)
							invoke SetTextColor,ps.hdc,m_clr.adrtxtcol
							mov 	edx,m_addrxp
							add 	edx,m_linenrwt
							add 	edx,m_selbarwt
							sub 	edx,cpx
							invoke TextOut,ps.hdc,edx,rect.top,addr buffer,8
						.endif
					.endif
					mov 	edx,m_dataxp
					add 	edx,m_linenrwt
					add 	edx,m_selbarwt
					sub 	edx,cpx
					.if nHilite
						.if nLeft
							push	edx
							invoke SetTextColor,ps.hdc,m_clr.dtatxtcol
							pop 	edx
							mov 	eax,nLeft
							shl 	eax,1
							add 	eax,nLeft
							push	eax
							push	edx
							invoke TextOut,ps.hdc,edx,rect.top,addr buffer[8],eax
							pop 	ecx
							pop 	eax
							mov 	edx,m_fntinfo.fntwt
							mul 	edx
							mov 	edx,ecx
							add 	edx,eax
						.endif
						push	edx
						.if (m_fHasFocus)
							invoke SetTextColor,ps.hdc,m_clr.seltxtcol
						.else
							invoke SetTextColor,ps.hdc, 0
						.endif
						pop 	edx
						mov 	eax,nHilite
						shl 	eax,1
						add 	eax,nHilite
						mov 	ecx,nLeft
						shl 	ecx,1
						add 	ecx,nLeft
						push	eax
						push	edx
						invoke TextOut,ps.hdc,edx,rect.top,addr buffer[ecx+8],eax
						pop 	ecx
						pop 	eax
						dec 	eax
						mov 	edx,m_fntinfo.fntwt
						mul 	edx
						mov 	edx,ecx
						add 	edx,eax
						.if nRight
							push	edx
							invoke SetTextColor,ps.hdc,m_clr.dtatxtcol
							pop 	edx
							mov 	eax,nRight
							shl 	eax,1
							add 	eax,nRight
							mov 	ecx,nLeft
							add 	ecx,nHilite
							shl 	ecx,1
							add 	ecx,nLeft
							add 	ecx,nHilite
							dec 	ecx
							inc 	eax
							invoke TextOut,ps.hdc,edx,rect.top,addr buffer[ecx+8],eax
						.endif
					.else
						push	edx
						invoke SetTextColor,ps.hdc,m_clr.dtatxtcol
						pop 	edx
						invoke TextOut,ps.hdc,edx,rect.top,addr buffer[8],47
					.endif
					test	m_fstyle,STYLE_NOASCII
					.if ZERO?
						invoke SetTextColor,ps.hdc,m_clr.asctxtcol
						mov 	edx,m_asciixp
						add 	edx,m_linenrwt
						add 	edx,m_selbarwt
						sub 	edx,cpx
						invoke TextOut,ps.hdc,edx,rect.top,addr buffer[8+47+1],16
					.endif
				.endif
				invoke SelectClipRgn,ps.hdc,NULL
				push	rect.left
				push	rect.right
				mov 	rect.left,0
				mov 	eax,m_linenrwt
				add 	eax,m_selbarwt
				.if eax
					mov 	rect.right,eax
					invoke FillRect,ps.hdc,addr rect,m_br.hBrSelBar
					invoke MoveToEx,ps.hdc,rect.right,0,NULL
					invoke LineTo,ps.hdc,rect.right,rect.bottom
if ?BOOKMARK
					.if esi
						invoke FindBookmark,m_hwnd,edi
						.if eax!=-1
							mov 	eax,m_selbarwt
							add 	eax,m_linenrwt
							sub 	eax,12
							mov 	edx,m_fntinfo.fntht
							sub 	edx,7
							shr 	edx,1
							add 	edx,rect.top
;;							invoke ImageList_Draw,hIml,2,ps.hdc,eax,edx,ILD_NORMAL
						.endif
					.endif
endif
if ?LINENUMBER
					mov 	eax,m_linenrwt
					.if eax && esi
						mov 	rect.right,eax
						dec 	rect.bottom
						invoke SelectObject,ps.hdc,m_fnt.hLnrFont
						mov 	edx,edi
						inc 	edx
						invoke DwToAscii,edx,addr buffer
						invoke SetTextColor,ps.hdc,m_clr.lnrcol
						invoke DrawText,ps.hdc,addr buffer,-1,addr rect,DT_RIGHT or DT_SINGLELINE or DT_BOTTOM; or DT_VCENTER;
					.endif
endif
				.endif
				pop 	rect.right
				pop 	rect.left
				pop 	rect.bottom
			.endif
		.endif
		pop 	edx
		inc 	edi
		add 	edx,m_fntinfo.fntht
		mov 	rect.top,edx
	.endw
if ?USEGLOBAL
	invoke GlobalUnlock,m_hmem
endif
	;Restore pen
	pop eax
	invoke SelectObject,ps.hdc,eax
	;Restore font
	pop eax
	invoke SelectObject,ps.hdc,eax
	invoke DeleteObject,hRgn
	invoke EndPaint,hWin,addr ps
	ret

;--- draw background of selected area

DrawSelBack:
	xor 	eax,eax
	mov 	nLeft,eax
	mov 	nHilite,eax
	mov 	nRight,eax
	mov 	eax,nSt
	mov 	ecx,nEn
	mov 	edx,edi
	shl 	edx,4
	.if eax>=edx && eax!=ecx
		sub 	eax,edx
		.if eax<16
			mov 	nLeft,eax
			add 	edx,16
			.if ecx<=edx
				sub 	edx,ecx
				mov 	nRight,edx
				sub 	ecx,eax

				mov 	edx,edi
				shl 	edx,4
				sub 	ecx,edx
				mov 	nHilite,ecx
			.else
				.if edx>16
					mov 	edx,16
				.endif
				sub 	edx,eax
				mov 	nHilite,edx
			.endif
		.endif
	.elseif ecx>edx && eax!=ecx
		sub 	ecx,edx
		.if ecx>16
			mov 	ecx,16
		.endif
		mov 	nHilite,ecx
		mov 	eax,16
		sub 	eax,ecx
		mov 	nRight,eax
	.endif
	.if nHilite
		push	rect.left
		push	rect.right
		mov 	eax,m_cpx
		mov 	edx,m_fntinfo.fntwt
		mul 	edx
		mov 	edx,eax
		mov 	eax,m_dataxp
		add 	eax,m_linenrwt
		add 	eax,m_selbarwt
		sub 	eax,edx
		mov 	rect.left,eax
		add 	eax,m_datawt
		mov 	rect.right,eax
		mov 	ecx,m_fntinfo.fntwt
		mov 	eax,nLeft
		mul 	ecx
		mov 	edx,3
		mul 	edx
		add 	rect.left,eax
		mov 	eax,nRight
		mul 	ecx
		mov 	edx,3
		mul 	edx
		sub 	rect.right,eax
		.if (m_fHasFocus)
			mov eax, m_br.hBrSelBck
		.else
			invoke GetSysColorBrush, COLOR_BTNFACE
		.endif
		invoke FillRect, ps.hdc, addr rect, eax
		pop rect.right
		pop rect.left
	.endif
	retn
	align 4

OnPaint endp

ifdef @StackBase
	option stackbase:esp
endif

InvalidateLine proc uses ebx esi hWin:HWND,cp:DWORD
	LOCAL	rect:RECT

	invoke GetWindowLong,hWin,MYWINDATA
	mov 	ebx,eax
	mov 	eax,hWin
	.if eax==m_edta.hwnd
		lea 	esi,m_edta
	.else
		lea 	esi,m_edtb
	.endif
	invoke CopyRect,addr rect,addr [esi].HEEDT.rc
	mov 	eax,cp
	shr 	eax,5
	sub 	eax,[esi].HEEDT.nline
	.if !CARRY?
		mov 	ecx,m_fntinfo.fntht
		mul 	ecx
		.if eax<rect.bottom
			mov 	rect.top,eax
			add 	eax,ecx
			mov 	rect.bottom,eax
			invoke InvalidateRect,hWin,addr rect,TRUE
		.endif
	.endif
	ret
	align 4

InvalidateLine endp

InvalidateSelection proc uses esi dwIndex:DWORD,cpMin:DWORD,cpMax:DWORD
	LOCAL	rect:RECT

	.if (dwIndex == 0)
		lea esi, m_edta
	.else
		lea esi, m_edtb
	.endif
	mov eax,cpMin
	mov ecx,cpMax
	shr eax,5
	shr ecx,5
	.if eax > ecx
		xchg eax,ecx
	.endif
	.if eax<[esi].HEEDT.nline
		mov eax,[esi].HEEDT.nline
	.endif
	.while eax<=ecx
		push eax
		push ecx
		push eax
		invoke CopyRect,addr rect,addr [esi].HEEDT.rc
		pop eax
		sub eax,[esi].HEEDT.nline
		mov ecx,m_fntinfo.fntht
		mul ecx
		.if eax<rect.bottom
			mov rect.top,eax
			add eax,ecx
			mov rect.bottom,eax
			invoke InvalidateRect, [esi].HEEDT.hwnd, addr rect, TRUE
			clc
		.else
			stc
		.endif
		pop ecx
		pop eax
		.break .if CARRY?
		inc eax
	.endw
	ret
	align 4

InvalidateSelection endp


;include Find.asm

if ?FINDTEXT

FindTextUp proc uses esi edi lpText:DWORD,len:DWORD,cpMin:DWORD,cpMax:DWORD
	LOCAL	lpMem:DWORD

if ?USEGLOBAL
	invoke GlobalLock,m_hmem
else
	mov		eax, m_hmem
endif
	mov 	lpMem,eax
	mov 	esi,eax
	mov 	eax,cpMin
	mov 	ecx,cpMax
	.if eax>ecx
		xchg	eax,ecx
	.endif
	shr 	eax,1
	shr 	ecx,1
	mov 	cpMin,eax
	add 	ecx,len
	mov 	cpMax,ecx
	mov 	eax,m_nbytes
	.if eax<cpMax
		mov 	cpMax,eax
	.endif
	add 	cpMin,esi
	add 	cpMax,esi
	mov 	esi,cpMax
	.while esi>=cpMin
		call	Compare
		.break .if !eax
		dec 	esi
	.endw
	.if eax
		xor 	eax,eax
		dec 	eax
	.else
		mov 	eax,esi
		sub 	eax,lpMem
		shl 	eax,1
	.endif
if ?USEGLOBAL
	push	eax
	invoke GlobalUnlock,m_hmem
	pop 	eax
endif
	ret

Compare:
ifdef @StackBase
if @ProcStatus and 4
@StackBase = @StackBase + 4
endif
endif
	mov 	edi,lpText
	mov 	ecx,len
	lea 	eax,[esi+ecx]
	cmp 	eax,cpMax
	jg		Ex
	xor 	eax,eax
  @@:
	dec 	ecx
	js		Ex
	mov 	al,[esi+ecx]
	sub 	al,[edi+ecx]
	je		@b
  Ex:
	retn
ifdef @StackBase
if @ProcStatus and 4
@StackBase = @StackBase - 4
endif
endif
	align 4

FindTextUp endp

FindTextDown proc uses esi edi lpText:DWORD,len:DWORD,cpMin:DWORD,cpMax:DWORD
	LOCAL	lpMem:DWORD

if ?USEGLOBAL
	invoke GlobalLock,m_hmem
else
	mov eax, m_hmem
endif
	mov lpMem,eax
	mov esi,eax
	mov eax,cpMin
	mov ecx,cpMax
	.if eax > ecx
		xchg eax,ecx
	.endif
	shr eax,1
	shr ecx,1
	mov cpMin,eax
	mov cpMax,ecx
	mov eax,m_nbytes
	.if eax < cpMax
		mov cpMax,eax
	.endif
	add cpMin,esi
	add cpMax,esi
	mov esi,cpMin
	.while esi < cpMax
		call Compare
		.break .if !eax
		inc esi
	.endw
	.if eax
		xor eax,eax
		dec eax
	.else
		mov eax,esi
		sub eax,lpMem
		shl eax,1
	.endif
if ?USEGLOBAL
	push eax
	invoke GlobalUnlock,m_hmem
	pop eax
endif
	ret

Compare:
ifdef @StackBase
if @ProcStatus and 4
@StackBase = @StackBase + 4
endif
endif
	mov edi,lpText
	mov ecx,len
	lea eax,[esi+ecx]
	cmp eax,cpMax
	jg Ex
	xor eax,eax
  @@:
	dec ecx
	js Ex
	mov al,[esi+ecx]
	sub al,[edi+ecx]
	je @b
  Ex:
	retn
ifdef @StackBase
if @ProcStatus and 4
@StackBase = @StackBase - 4
endif
endif
	align 4

FindTextDown endp

FindTextEx proc uses esi edi fFlag:DWORD,lpFindTextEx:DWORD
	LOCAL lpText:DWORD
	LOCAL len:DWORD

	mov esi,lpFindTextEx
	mov eax,[esi].FINDTEXTEX.lpstrText
	mov lpText,eax
	invoke lstrlen,eax
	.if eax
		mov len,eax
		mov eax,fFlag
		and eax,FR_HEX
		.if eax
			.if (!g_pCharBuff)
				invoke malloc, ?SIZECHARBUFF
				mov g_pCharBuff, eax
			.endif
			invoke ConvertHexString,lpText, g_pCharBuff
			mov len,eax
			mov eax, g_pCharBuff
			mov lpText, eax
		.endif
		mov eax,[esi].FINDTEXTEX.chrg.cpMin
		.if eax<[esi].FINDTEXTEX.chrg.cpMax
			;Down
			invoke FindTextDown, lpText,len,[esi].FINDTEXTEX.chrg.cpMin,[esi].FINDTEXTEX.chrg.cpMax
		.elseif eax > [esi].FINDTEXTEX.chrg.cpMax
			;Up
			invoke FindTextUp, lpText,len,[esi].FINDTEXTEX.chrg.cpMax,[esi].FINDTEXTEX.chrg.cpMin
		.else
			mov eax,-1
		.endif
		.if eax!=-1
			mov [esi].FINDTEXTEX.chrgText.cpMin,eax
			mov edx,len
			shl edx,1
			add edx,eax
			mov [esi].FINDTEXTEX.chrgText.cpMax,edx
		.endif
	.else
		xor eax,eax
		dec eax
	.endif
	ret
	align 4

FindTextEx endp
endif

;include ClipBoard.asm

;--- nMode: 0 == private, 1 = CF_TEXT

SetClipData proc nMode:DWORD, lpData:LPSTR, dwSize:dword

	LOCAL hMem:HANDLE
	LOCAL pMem:dword

	mov eax,dwSize
	inc eax
	invoke GlobalAlloc,GHND or GMEM_DDESHARE,eax
	test eax,eax
	je @exit2
	mov hMem,eax
	invoke GlobalLock,eax	;hGlob
	test eax,eax
	je @exit1
	mov pMem,eax
	invoke RtlMoveMemory,eax,lpData,dwSize
	mov eax,pMem
	add eax,dwSize
	mov byte ptr [eax],0
	invoke GlobalUnlock,hMem
	.if (nMode)
		mov ecx, CF_TEXT
	.else
		mov ecx, g_dwMyCBFormat
	.endif
	invoke SetClipboardData, ecx ,hMem
	xor eax,eax 	;0 - Ok
	jmp @exit3
  @exit1:
	invoke GlobalFree,hMem
	xor eax,eax
  @exit2:
	dec eax 		 ; -1 - error
  @exit3:
	ret
	align 4

SetClipData endp

;--- nMode: 0 == private, 1 = CF_TEXT

EditCopy_ proc uses esi edi nMode:DWORD

	LOCAL hCMem:LPVOID
	LOCAL lpCMem:DWORD
	LOCAL cpMin:DWORD
	LOCAL cpMax:DWORD

	mov eax,m_cpMin
	mov edx,m_cpMax
	shr eax,1
	shr edx,1
	.if eax != edx
		.if eax > edx
			xchg eax,edx
		.endif
		mov cpMin,eax
		mov cpMax,edx
		sub edx,eax
		mov eax,edx

		.if (!nMode)
			shr edx,4
			add eax,edx
			shl eax,1
			add eax,3
		.endif
		invoke malloc, eax
		mov lpCMem,eax
		mov edi,lpCMem
if ?USEGLOBAL
		invoke GlobalLock,m_hmem
else
		mov eax, m_hmem
endif
		mov esi,eax
		xor ecx,ecx
		mov edx,cpMin
		.while edx < cpMax
			push edx
			mov al,[esi+edx]
			.if (nMode)
				stosb
			.else
				shl eax,24
				call Nybble
				mov [edi],dl
				inc edi
				call Nybble
				mov [edi],dl
				inc edi
				inc ecx
				.if ecx==16
					mov byte ptr [edi],0Dh
					inc edi
					mov byte ptr [edi],0Ah
					inc edi
					xor ecx,ecx
				.endif
			.endif
			pop edx
			inc edx
		.endw
		.if (ecx && (nMode == 0))
			mov byte ptr [edi],0Dh
			inc edi
			mov byte ptr [edi],0Ah
			inc edi
		.endif
if ?USEGLOBAL
		invoke GlobalUnlock,m_hmem
endif
		sub edi,lpCMem
		invoke SetClipData, nMode, lpCMem, edi
		invoke free, lpCMem
	.endif
	ret
	align 4

EditCopy_ endp

EditCopy proc
	invoke OpenClipboard,NULL
	.if (eax)
		invoke EmptyClipboard
		invoke EditCopy_, 0
		invoke EditCopy_, 1
		invoke CloseClipboard
	.endif
	ret
EditCopy endp

EditPaste proc

local nMode:DWORD

	invoke OpenClipboard,m_hwnd
	.if eax
		mov nMode, 0
		invoke GetClipboardData, g_dwMyCBFormat
		.if (!eax)
			invoke GetClipboardData, CF_TEXT
			inc nMode
		.endif

		.if eax
			push eax
			invoke GlobalLock,eax
			push eax
			invoke SendMessage,m_hwnd,WM_CLEAR,0,0
			pop eax
			mov edx,m_cpMin
			.if edx > m_cpMax
				mov edx,m_cpMax
			.endif
			.if (nMode)
				push edx
				shr edx, 1
				invoke InsertAsciiString, edx, eax
				pop edx
			.else
				shr edx,1
				shl edx,1
				push edx
				shr edx,1
				invoke InsertHexString, edx, eax
				pop edx
			.endif
			shl eax,1
			add eax,edx
			mov m_cpMin,eax
			mov m_cpMax,eax
			pop eax
			invoke GlobalUnlock,eax
			invoke InvalidateRect,m_edta.hwnd,NULL,FALSE
			invoke InvalidateRect,m_edtb.hwnd,NULL,FALSE
			mov m_fChanged,TRUE
			inc m_nchange
if ?STATE
			invoke InvalidateRect,m_hsta,NULL,FALSE
endif
			invoke ScrollCaret, m_hwndFocus
			invoke SelChange, SEL_TEXT
			push word ptr EN_CHANGE
			push word ptr m_ID
			pop ecx
			invoke SendMessage, m_hwndParent, WM_COMMAND, ecx, m_hwnd
		.endif
		invoke CloseClipboard
	.endif
	ret
	align 4

EditPaste endp

EditSelectAll proc hWin:HWND
	mov eax,m_nbytes
	shl eax,1
	mov m_cpMin,0
	mov m_cpMax,eax
	invoke InvalidateRect,m_edta.hwnd,NULL,FALSE
	invoke InvalidateRect,m_edtb.hwnd,NULL,FALSE
	invoke SetCaret,hWin
	invoke SelChange, SEL_TEXT
	ret
	align 4
EditSelectAll endp

EditDelete proc

	mov eax, m_cpMin
	mov ecx, m_cpMax
	.if (eax > ecx)
		xchg eax, ecx
	.endif
	.if (eax != ecx)
		invoke DeleteSelection, eax, ecx
	.else
		and eax,0FFFFFFFEh
		mov edx,eax
		add edx,2
		invoke DeleteSelection, eax, edx
	.endif
	ret
	align 4
EditDelete endp

;include Undo.asm


ExpandUndoMem proc uses edi cb:DWORD

	invoke GetProcessHeap
	mov edi,eax
	mov eax,m_rpundo
	add eax,cb
	add eax,sizeof HEUNDO*2
	.if eax > m_cbundo
		and eax,-1 xor (MAXUNDOMEM-1)
		add eax,MAXUNDOMEM
		mov m_cbundo,eax
		invoke HeapReAlloc,edi,HEAP_GENERATE_EXCEPTIONS or HEAP_NO_SERIALIZE or HEAP_ZERO_MEMORY,m_hundo,m_cbundo
		.if !eax
			invoke MessageBox,m_hwnd,offset szMemFailUndo,offset szToolTip,MB_OK
			xor eax,eax
		.else
			mov m_hundo,eax
		.endif
	.endif
	ret
	align 4

ExpandUndoMem endp

DoUndo proc uses esi edi

	mov edi,m_hundo
	mov edx,m_rpundo
	.if edx
		mov edx,[edi+edx].HEUNDO.rpPrev
		mov m_rpundo,edx
		mov al,[edi+edx].HEUNDO.fun
		.if al==UNDO_CHARINSERT
			lea edi,[edi+edx]
			mov esi,[edi].HEUNDO.cp
			mov m_cpMin,esi
			mov m_cpMax,esi
			shr esi,1
			invoke DeleteChars, esi, 1
		.elseif al==UNDO_CHAROVERWRITE
			lea edi,[edi+edx]
			mov esi,[edi].HEUNDO.cp
			mov m_cpMin,esi
			mov m_cpMax,esi
			shr esi,1
			invoke GetChar, esi
			push eax
			invoke PutChar, esi, addr [edi+sizeof HEUNDO],1
			pop eax
			mov [edi+sizeof HEUNDO],al
		.elseif al==UNDO_INSERTBLOCK
			lea edi,[edi+edx]
			mov esi,[edi].HEUNDO.cp
			mov m_cpMin,esi
			mov m_cpMax,esi
			shr esi,1
			invoke DeleteChars, esi, [edi].HEUNDO.cb
		.elseif al==UNDO_DELETEBLOCK
			lea edi,[edi+edx]
			mov esi,[edi].HEUNDO.cp
			mov m_cpMin,esi
			shr esi,1
			invoke InsertChars, esi, [edi].HEUNDO.cb
			invoke PutChar, esi, addr [edi+sizeof HEUNDO],[edi].HEUNDO.cb
			add esi,[edi].HEUNDO.cb
			shl esi,1
			mov m_cpMax,esi
		.endif
	.endif
	ret
	align 4

DoUndo endp

DoRedo proc uses esi edi

	mov edi,m_hundo
	mov edx,m_rpundo
	mov eax,[edi+edx].HEUNDO.cb
	.if eax
		mov al,[edi+edx].HEUNDO.fun
		.if al==UNDO_CHARINSERT
			lea edi,[edi+edx]
			mov esi,[edi].HEUNDO.cp
			mov m_cpMin,esi
			mov m_cpMax,esi
			shr esi,1
			invoke InsertChars, esi, 1
			invoke PutChar, esi, addr [edi+sizeof HEUNDO], 1
			inc m_cpMin
			inc m_cpMax
			add m_rpundo,sizeof HEUNDO+1
		.elseif al==UNDO_CHAROVERWRITE
			lea edi,[edi+edx]
			mov esi,[edi].HEUNDO.cp
			mov m_cpMin,esi
			mov m_cpMax,esi
			shr esi,1
			invoke GetChar, esi
			push eax
			invoke PutChar, esi, addr [edi+sizeof HEUNDO], 1
			pop eax
			mov [edi+sizeof HEUNDO],al
			inc m_cpMin
			inc m_cpMax
			add m_rpundo,sizeof HEUNDO+1
		.elseif al==UNDO_INSERTBLOCK
			lea edi,[edi+edx]
			mov esi,[edi].HEUNDO.cp
			mov m_cpMin,esi
			mov m_cpMax,esi
			shr esi,1
			invoke InsertChars, esi, [edi].HEUNDO.cb
			invoke PutChar, esi, addr [edi+sizeof HEUNDO], [edi].HEUNDO.cb
			mov eax,[edi].HEUNDO.cb
			add eax,sizeof HEUNDO
			add m_rpundo,eax
		.elseif al==UNDO_DELETEBLOCK
			lea edi,[edi+edx]
			mov esi,[edi].HEUNDO.cp
			mov m_cpMin,esi
			mov m_cpMax,esi
			shr esi,1
			invoke DeleteChars, esi, [edi].HEUNDO.cb
			mov eax,[edi].HEUNDO.cb
			add eax,sizeof HEUNDO
			add m_rpundo,eax
		.endif
	.endif
	ret
	align 4

DoRedo endp

SaveUndo proc uses esi edi nFun:DWORD,cp:DWORD,cr:DWORD,cb:DWORD

	invoke ExpandUndoMem, cb
	or eax,eax
	je Ex
	mov edi,m_hundo
	mov edx,m_rpundo
	mov eax,nFun
	.if eax==UNDO_CHARINSERT || eax==UNDO_CHAROVERWRITE
		lea edi,[edi+edx]
		mov [edi].HEUNDO.fun,al
		mov eax,cp
		mov [edi].HEUNDO.cp,eax
		mov [edi].HEUNDO.cb,1
		mov eax,cr
		mov [edi+sizeof HEUNDO],al
		mov [edi+sizeof HEUNDO+1].HEUNDO.rpPrev,edx
		add m_rpundo,sizeof HEUNDO+1
	.elseif eax==UNDO_INSERTBLOCK || eax==UNDO_DELETEBLOCK
		lea edi,[edi+edx]
		mov [edi].HEUNDO.fun,al
		mov esi,cp
		mov [edi].HEUNDO.cp,esi
		mov ecx,cb
		mov [edi].HEUNDO.cb,ecx
		mov eax,sizeof HEUNDO
		add eax,cb
		mov [edi+eax].HEUNDO.rpPrev,edx
		add eax,edx
		mov m_rpundo,eax

		shr esi,1
		add edi,sizeof HEUNDO
		.while ecx
			push ecx
			invoke GetChar, esi
			pop ecx
			dec ecx
			mov [edi],al
			inc esi
			inc edi
		.endw
	.endif
  Ex:
	ret
	align 4

SaveUndo endp

EditUndo proc hWin:DWORD

	invoke DoUndo
	invoke ScrollCaret,hWin
	invoke InvalidateRect,m_edta.hwnd,NULL,FALSE
	invoke InvalidateRect,m_edtb.hwnd,NULL,FALSE
	invoke SetCaret,hWin
	invoke SelChange, SEL_TEXT
	ret
	align 4

EditUndo endp

EditRedo proc hWin:DWORD

	invoke DoRedo
	invoke ScrollCaret,hWin
	invoke InvalidateRect,m_edta.hwnd,NULL,FALSE
	invoke InvalidateRect,m_edtb.hwnd,NULL,FALSE
	invoke SetCaret,hWin
	invoke SelChange, SEL_TEXT
	ret
	align 4

EditRedo endp

if ?BOOKMARK

;include Bookmark.asm


	.data?

iBookmark			dd ?
Bookmark			HEBMK 32 dup(<?>)

	.code

SetBookmark proc uses edi hWin:DWORD,nLine:DWORD

	mov edi,offset Bookmark
	xor ecx,ecx
	.while ecx < sizeof Bookmark/sizeof HEBMK
		.if ![edi+ecx*sizeof HEBMK].HEBMK.hWin
			mov eax,hWin
			mov [edi+ecx*sizeof HEBMK].HEBMK.hWin,eax
			mov eax,nLine
			mov [edi+ecx*sizeof HEBMK].HEBMK.nLine,eax
			jmp Ex
		.endif
		inc ecx
	.endw
	xor ecx,ecx
	dec ecx
  Ex:
	mov eax,ecx
	ret
	align 4

SetBookmark endp

ClrBookmark proc uses edi nInx:DWORD

	mov edi,offset Bookmark
	mov ecx,nInx
	mov edx,ecx
	inc edx
	.while edx < sizeof Bookmark/sizeof HEBMK
		mov eax,[edi+edx*sizeof HEBMK].HEBMK.hWin
		mov [edi+ecx*sizeof HEBMK].HEBMK.hWin,eax
		mov eax,[edi+edx*sizeof HEBMK].HEBMK.nLine
		mov [edi+ecx*sizeof HEBMK].HEBMK.nLine,eax
		inc edx
		inc ecx
	.endw
	xor eax,eax
	mov [edi+ecx*sizeof HEBMK].HEBMK.hWin,eax
	mov [edi+ecx*sizeof HEBMK].HEBMK.nLine,eax
	ret
	align 4

ClrBookmark endp

FindBookmark proc uses edi hWin:DWORD,nLine:DWORD

	mov edi,offset Bookmark
	mov eax,hWin
	mov edx,nLine
	xor ecx,ecx
	.while ecx < sizeof Bookmark/sizeof HEBMK
		cmp eax,[edi+ecx*sizeof HEBMK].HEBMK.hWin
		jne @f
		cmp edx,[edi+ecx*sizeof HEBMK].HEBMK.nLine
		je Ex
	  @@:
		inc ecx
	.endw
	xor ecx,ecx
	dec ecx
  Ex:
	mov eax,ecx
	ret
	align 4

FindBookmark endp

ToggleBookmark proc hWin:DWORD,nLine:DWORD

	invoke FindBookmark,hWin,nLine
	.if eax==-1
		invoke SetBookmark,hWin,nLine
		.if eax!=-1
			mov 	iBookmark,eax
		.endif
	.else
		invoke ClrBookmark,eax
	.endif
	ret
	align 4

ToggleBookmark endp

NxtBookmark proc uses edi lpHEBMK:DWORD

	mov edi,offset Bookmark
	mov ecx,iBookmark
	inc ecx
	.if ecx >= sizeof Bookmark/sizeof HEBMK
		xor ecx,ecx
	.endif
	.if ![edi+ecx*sizeof HEBMK].HEBMK.hWin
		xor ecx,ecx
	.endif
	mov iBookmark,ecx
	mov edx,lpHEBMK
	mov eax,[edi+ecx*sizeof HEBMK].HEBMK.nLine
	mov [edx].HEBMK.nLine,eax
	mov eax,[edi+ecx*sizeof HEBMK].HEBMK.hWin
	mov [edx].HEBMK.hWin,eax
	.if eax
		xor eax,eax
		inc eax
	.endif
	ret
	align 4

NxtBookmark endp

PrvBookmark proc uses edi lpHEBMK:DWORD

	mov edi,offset Bookmark
	mov ecx,iBookmark
	dec ecx
	jns @f
	mov ecx,sizeof Bookmark/sizeof HEBMK-1
  @@:
	.while ecx
		.break .if [edi+ecx*sizeof HEBMK].HEBMK.hWin
		dec ecx
	.endw
	mov iBookmark,ecx
	mov edx,lpHEBMK
	mov eax,[edi+ecx*sizeof HEBMK].HEBMK.nLine
	mov [edx].HEBMK.nLine,eax
	mov eax,[edi+ecx*sizeof HEBMK].HEBMK.hWin
	mov [edx].HEBMK.hWin,eax
	.if eax
		xor eax,eax
		inc eax
	.endif
	ret
	align 4

PrvBookmark endp

ClearBookmarks proc uses esi edi

	mov edi,offset Bookmark
	xor ecx,ecx
	.while ecx < sizeof Bookmark/sizeof HEBMK
		mov eax,[edi+ecx*sizeof HEBMK].HEBMK.hWin
		.if eax
			mov esi,[edi+ecx*sizeof HEBMK].HEBMK.nLine
			push ecx
			invoke GetWindowLong,eax,MYWINDATA
			mov ebx,eax
			shl esi,5
			invoke InvalidateLine,m_edta.hwnd,esi
			invoke InvalidateLine,m_edtb.hwnd,esi
			pop ecx
			xor eax,eax
			mov [edi+ecx*sizeof HEBMK].HEBMK.hWin,eax
			mov [edi+ecx*sizeof HEBMK].HEBMK.nLine,eax
		.endif
		inc ecx
	.endw
	ret
	align 4

ClearBookmarks endp

ClearWinBookmarks proc uses edi hWin:DWORD

	mov edi,offset Bookmark
	xor ecx,ecx
	.while ecx < sizeof Bookmark/sizeof HEBMK
		mov eax,[edi+ecx*sizeof HEBMK].HEBMK.hWin
		.if eax==hWin
			push ecx
			invoke ClrBookmark,ecx
			pop ecx
		.endif
		inc ecx
	.endw
	ret
	align 4

ClearWinBookmarks endp

endif

;--- Create a windowclass for the user control


CreateClass proc hInst:HINSTANCE
	LOCAL wc:WNDCLASSEX

	mov wc.cbSize,sizeof WNDCLASSEX
;	mov wc.style,CS_GLOBALCLASS or CS_HREDRAW or CS_VREDRAW
	mov wc.style,CS_GLOBALCLASS
	mov wc.lpfnWndProc,offset RAWndProc
	mov wc.cbClsExtra, NULL
if MYWINDATA eq GWL_USERDATA
	mov wc.cbWndExtra,0
else
	mov wc.cbWndExtra,4 	;Holds memory handle
endif
	push hInst
	pop wc.hInstance
	mov wc.hbrBackground,NULL
	mov wc.lpszMenuName,NULL
	mov wc.lpszClassName,offset szClassName
	mov eax,NULL
	mov wc.hIcon,eax
	mov wc.hIconSm,eax
	invoke LoadCursor,NULL,IDC_ARROW
	mov wc.hCursor,eax
	invoke RegisterClassEx,addr wc

	mov wc.cbSize,sizeof WNDCLASSEX
;	mov wc.style,CS_DBLCLKS or CS_HREDRAW or CS_VREDRAW
	mov wc.style,CS_DBLCLKS
	mov wc.lpfnWndProc,offset RAHexEdProc
	mov wc.cbClsExtra,NULL
if MYWINDATA eq GWL_USERDATA
	mov wc.cbWndExtra,0
else
	mov wc.cbWndExtra,4 	;Holds memory handle
endif
	push hInst
	pop wc.hInstance
	mov wc.hbrBackground,NULL
	mov wc.lpszMenuName,NULL
	mov wc.lpszClassName,offset szEditClassName
	mov eax,NULL
	mov wc.hIcon,eax
	mov wc.hIconSm,eax
	invoke LoadCursor,NULL,IDC_IBEAM
	mov wc.hCursor,eax
	invoke RegisterClassEx,addr wc
	ret
	align 4

CreateClass endp

TimerProc proc hWin:DWORD,uMsg:DWORD,idEvent:DWORD,dwTime:DWORD

	.if tmr1.hwnd
		invoke PostMessage,tmr1.hwnd,tmr1.umsg,tmr1.wparam,tmr1.lparam
	.endif
	.if tmr2.hwnd
		invoke PostMessage,tmr2.hwnd,tmr2.umsg,tmr2.wparam,tmr2.lparam
	.endif
	ret
	align 4

TimerProc endp

SplittBtnProc proc uses ebx esi edi hWin:HWND,uMsg:UINT,wParam:WPARAM,lParam:LPARAM

local	rect:RECT

	mov eax,uMsg
	.if eax==WM_SETCURSOR
		invoke SetCursor, g_hHSCur
		mov eax, 1
		jmp done
	.elseif eax==WM_LBUTTONDOWN
		invoke GetParent,hWin
		invoke GetWindowLong,eax,MYWINDATA
		mov ebx,eax
		mov m_fresize,1

		invoke SetCapture,m_hwnd

		invoke GetWindowRect, m_hwnd, addr rect
		mov ecx, rect.bottom
		sub ecx, rect.top
		sub ecx, MINSIZE
		.if (!SIGN?)
			sub rect.bottom, MINSIZE
		.endif
		invoke ClipCursor, addr rect

	.elseif eax==WM_LBUTTONDBLCLK

		invoke GetParent,hWin
		invoke GetWindowLong,eax,MYWINDATA
		mov ebx,eax
		.if m_fsplitt
			mov m_fsplitt,0
		.else
			mov m_fsplitt,512
		.endif
		invoke SendMessage,m_hwnd, WM_SIZE, 0, 0
	.endif
	invoke CallWindowProc,OldSplittBtnProc,hWin,uMsg,wParam,lParam
done:
	ret
	align 4

SplittBtnProc endp

if ?STATE

StateProc proc uses ebx hWin:HWND,uMsg:UINT,wParam:WPARAM,lParam:LPARAM
	LOCAL ps:PAINTSTRUCT
	LOCAL hBr:DWORD

	.if uMsg==WM_PAINT
		invoke BeginPaint,hWin,addr ps
		invoke GetWindowLong,hWin,GWL_USERDATA
		mov ebx,eax
		test m_fstyle,STYLE_READONLY
		.if !ZERO?
			invoke CreateSolidBrush,0FFh
			mov hBr,eax
		.elseif m_fChanged
			invoke CreateSolidBrush,0B000h
			mov hBr,eax
		.else
			mov hBr,0
			mov eax,m_br.hBrSelBar
		.endif
		invoke FillRect,ps.hdc,addr ps.rcPaint,eax
		.if hBr
			invoke DeleteObject,hBr
		.endif
		invoke EndPaint,hWin,addr ps
		xor eax,eax
		ret
	.endif
	invoke CallWindowProc,OldStateProc,hWin,uMsg,wParam,lParam
	ret
	align 4

StateProc endp

endif

DoMove proc hWin:HWND, fSel:DWORD, fShift:DWORD, cpMin:DWORD, cpMax:DWORD

	.if !fShift
		mov eax,m_cpMin
		mov m_cpMax,eax
	.endif
	invoke ScrollCaret, hWin
	mov eax,m_cpMin
	.if eax==m_cpMax
		.if fSel
			invoke InvalidateSelection, 0, cpMin, cpMax
			invoke InvalidateSelection, 1, cpMin, cpMax
		.endif
	.else
		push eax
		invoke InvalidateLine, m_edta.hwnd, eax
		pop eax
		invoke InvalidateLine, m_edtb.hwnd, eax
		invoke InvalidateLine, m_edta.hwnd, cpMin
		invoke InvalidateLine, m_edtb.hwnd, cpMin
		invoke InvalidateLine, m_edta.hwnd, cpMax
		invoke InvalidateLine, m_edtb.hwnd, cpMax
	.endif
	invoke SelChange, SEL_TEXT
	ret
	align 4
DoMove endp

OnKeyDown proc uses ebx hWin:HWND, wParam:WPARAM, lParam:LPARAM

	LOCAL	sinf:SCROLLINFO
	LOCAL	cpMin:DWORD
	LOCAL	cpMax:DWORD
	LOCAL	fControl:BYTE
	LOCAL	fShift:BYTE
	LOCAL	fSel:DWORD

	mov ebx, hWin
	mov eax,m_cpMin
	mov ecx,m_cpMax
	.if eax > ecx
		xchg eax,ecx
	.endif
	mov cpMin,eax
	mov cpMax,ecx
	sub ecx,eax
	mov fSel,ecx
	mov sinf.cbSize,sizeof sinf
	mov sinf.fMask,SIF_ALL
	invoke GetScrollInfo,[esi].HEEDT.hvscroll,SB_CTL,addr sinf
	invoke GetKeyState,VK_CONTROL
	and al,80h
	mov fControl,al

	invoke GetKeyState,VK_SHIFT
	and al,80h
	mov fShift,al

	mov edx,wParam
	mov eax,lParam
	shr eax,16
	and eax,3FFh		;scan code in al, ah[0] = extended key

;--- VK_DOWN=28h, VK_UP=26h, VK_RIGHT=27h, VK_LEFT=25h

	DebugOut "HexEdit, OnKeyDown, wParam=%X, lParam=%X, shift=%X", wParam, lParam, fShift

	.if edx==VK_DOWN	;; && (eax==150h || eax==50h)

		.if fControl
			;Ctrl+Down
			mov eax,SB_LINEDOWN
			call Scroll
		.else
			;Down
			mov ecx,m_nbytes
			shr ecx,4
			mov eax,m_cpMin
			mov edx,eax
			shr edx,5
			.if edx < ecx
				add eax,32
				mov edx,eax
				shr edx,1
				.if edx > m_nbytes
					mov eax,m_nbytes
					shl eax,1
				.endif
				mov m_cpMin,eax
			.endif
			.if fShift
				and m_cpMin,0FFFFFFFEh
				and m_cpMax,0FFFFFFFEh
			.elseif fSel
				mov eax,cpMax
				mov m_cpMin,eax
			.endif
			invoke DoMove, ebx, fSel, fShift, cpMin, cpMax
		.endif

	.elseif edx==VK_UP		;; && (eax==148h || eax==48h)

		.if fControl
			;Ctrl+Up
			mov eax,SB_LINEUP
			call Scroll
		.else
			;Up
			mov eax,m_cpMin
			sub eax,32
			jnb @f
			add eax,32
		@@:
			mov m_cpMin,eax
			.if fShift
				and m_cpMin,0FFFFFFFEh
				and m_cpMax,0FFFFFFFEh
			.elseif fSel
				mov eax,cpMin
				mov m_cpMin,eax
			.endif
			invoke DoMove, ebx, fSel, fShift, cpMin, cpMax
		.endif

	.elseif edx==VK_RIGHT	;; && (eax==14Dh || eax==4Dh)

		.if fControl || fShift
			;Ctrl+Right
			mov eax,m_cpMin
			and eax,0FFFFFFFEh
			inc eax
			inc eax
			mov edx,eax
			inc edx
			shr edx,1
			.if edx > m_nbytes
				dec eax
			.endif
			mov m_cpMin,eax
			and m_cpMin,0FFFFFFFEh
		.else
			;Right
			mov eax,m_cpMin
			inc eax
			mov edx,eax
			inc edx
			shr edx,1
			.if edx > m_nbytes
				dec eax
			.endif
			mov m_cpMin,eax
		.endif
		.if !fShift && fSel
			mov eax,cpMax
			mov m_cpMin,eax
		.endif
		invoke DoMove, ebx, fSel, fShift, cpMin, cpMax

	.elseif edx==VK_LEFT	;; && (eax==14Bh || eax==4Bh)

		.if fControl || fShift
			;Ctrl+Left
			mov eax,m_cpMin
			and eax,0FFFFFFFEh
			.if eax
				dec eax
				and eax,0FFFFFFFEh
				mov m_cpMin,eax
			.endif
			mov m_cpMin,eax
			and m_cpMin,0FFFFFFFEh
		.else
			;Left
			mov eax,m_cpMin
			.if eax
				dec eax
				mov m_cpMin,eax
			.endif
		.endif
		.if !fShift && fSel
			mov eax,cpMin
			mov m_cpMin,eax
		.endif
		invoke DoMove, ebx, fSel, fShift, cpMin, cpMax

;--- VK_PGDN=22h, VK_PGUP=21h, VK_END=23h, VK_HOME=24h

	.elseif edx==VK_NEXT	;; && (eax==151h || eax==51h)

		.if fControl
			;Ctrl+PgDn
			mov eax,SB_PAGEDOWN
			call Scroll
		.else
			;PgDn
			invoke SendMessage,ebx,WM_VSCROLL,SB_PAGEDOWN,0
			mov ecx,m_nbytes
			shr ecx,4
			mov eax,sinf.nPage
			shl eax,5
			add eax,m_cpMin
			mov edx,eax
			shr edx,5
			.if edx > ecx
				mov edx,ecx
				shl edx,5
				and eax,1Fh
				add eax,edx
				mov edx,eax
				shr edx,1
				.if edx > m_nbytes
					mov eax,m_nbytes
					shl eax,1
				.endif
			.endif
			mov edx,m_nbytes
			shl edx,1
			.if eax > edx
				mov eax,edx
			.endif
			mov m_cpMin,eax
			.if fShift
				and m_cpMin,0FFFFFFFEh
				and m_cpMax,0FFFFFFFEh
				invoke DoMove, ebx, fSel, fShift, cpMin, cpMax
				invoke InvalidateRect,hWin,NULL,FALSE
			.else
				invoke DoMove, ebx, fSel, fShift, cpMin, cpMax
			.endif
		.endif

	.elseif edx==VK_PRIOR	;; && (eax==149h || eax==49h)

		.if fControl
			;Ctrl+PgUp
			mov eax,SB_PAGEUP
			call Scroll
		.else
			;PgUp
			invoke SendMessage,ebx,WM_VSCROLL,SB_PAGEUP,0
			mov edx,sinf.nPage
			shl edx,5
			mov eax,m_cpMin
			sub eax,edx
			jnb @f
			and eax,1Fh
@@:
			mov m_cpMin,eax
			.if fShift
				and m_cpMin,0FFFFFFFEh
				and m_cpMax,0FFFFFFFEh
				invoke DoMove, ebx, fSel, fShift, cpMin, cpMax
				invoke InvalidateRect,ebx,NULL,FALSE
			.else
				invoke DoMove, ebx, fSel, fShift, cpMin, cpMax
			.endif
		.endif

	.elseif edx==VK_END	;; && (eax==14Fh || eax==4Fh)

		.if fControl
			;Ctrl+End
			mov eax,m_nbytes
			shl eax,1
			mov m_cpMin,eax
			.if !fShift
				mov m_cpMax,eax
			.else
				and m_cpMin,0FFFFFFFEh
				and m_cpMax,0FFFFFFFEh
			.endif
			mov eax,SB_BOTTOM
			call Scroll
			invoke DoMove, ebx, fSel, fShift, cpMin, cpMax
			invoke InvalidateRect,ebx,NULL,FALSE
		.else
			;End
			mov edx,m_nbytes
			shl edx,1
			mov eax,m_cpMin
			or eax,1Fh
			.if fShift
				inc eax
			.endif
			.if eax > edx
				mov eax,edx
			.endif
			mov m_cpMin,eax
			.if fShift
				and m_cpMin,0FFFFFFFEh
				and m_cpMax,0FFFFFFFEh
			.endif
			invoke DoMove, ebx, fSel, fShift, cpMin, cpMax
		.endif

	.elseif edx==VK_HOME	;; && (eax==147h || eax==47h)

		.if fControl
			;Ctrl+Home
			xor eax,eax
			mov m_cpMin,eax
			.if !fShift
				mov m_cpMax,eax
			.else
				and m_cpMin,0FFFFFFFEh
				and m_cpMax,0FFFFFFFEh
			.endif
			mov eax,SB_TOP
			call Scroll
			invoke DoMove, ebx, fSel, fShift, cpMin, cpMax
			invoke InvalidateRect,ebx,NULL,FALSE
		.else
			;Home
			mov eax,m_cpMin
			and eax,0FFFFFFE0h
			mov m_cpMin,eax
			invoke DoMove, ebx, fSel, fShift, cpMin, cpMax
			.if fShift
				and m_cpMin,0FFFFFFFEh
				and m_cpMax,0FFFFFFFEh
			.endif
		.endif

;--- VK_INSERT=2Dh, VK_DELETE=2Eh, VK_BACK

	.elseif edx==VK_INSERT	;; && (eax==152h || eax==52h)

		;Insert
		xor m_fOvr,TRUE
		invoke SetCaret,ebx

	.elseif edx==VK_DELETE	;; && (eax==153h || eax==53h)

		;Delete
		invoke EditDelete

	.elseif edx==VK_BACK && !fShift

		.if (m_cpMin)
			invoke OnKeyDown, ebx, VK_LEFT, lParam
			invoke OnKeyDown, ebx, VK_DELETE, lParam
		.else
			invoke MessageBeep, MB_OK
		.endif

;--- Control

	.elseif (fControl)

		.if edx==43h					;Ctrl+C, Copy

			invoke EditCopy

		.elseif edx==58h				;Ctrl+X, Cut

			invoke EditCopy
			invoke SendMessage, m_hwnd, WM_CLEAR, 0, 0

		.elseif edx==56h				;Ctrl+V, Paste

			invoke EditPaste

		.elseif edx==41h				;Ctrl+A, Select all

			invoke EditSelectAll, ebx

		.elseif edx==5Ah				;Ctrl+Z, Undo

			invoke EditUndo, ebx

		.elseif edx==59h				;Ctrl+Y, Redo

			invoke EditRedo, ebx
    	.endif

	.endif
	xor eax,eax
	ret
Scroll:
	invoke SendMessage, ebx, WM_VSCROLL, eax, 0
	retn
	align 4

OnKeyDown endp

OnChar	proc hWin:HWND, wParam:WPARAM, lParam:LPARAM

	LOCAL	cpMin:DWORD
	LOCAL	cpMax:DWORD
	LOCAL	nChar:DWORD
	LOCAL	fShift:BYTE
	LOCAL	fControl:BYTE
	LOCAL	fIns:DWORD
	LOCAL	fSel:DWORD

		mov fIns,FALSE
		mov eax,m_cpMin
		mov ecx,m_cpMax
		.if eax > ecx
			xchg eax,ecx
		.endif
		mov cpMin,eax
		mov cpMax,ecx
		sub ecx,eax
		mov fSel,ecx
		invoke GetKeyState,VK_CONTROL
		and al,80h
		mov fControl,al
		invoke GetKeyState,VK_SHIFT
		and al,80h
		mov fShift,al
		mov eax,wParam
		.if (eax >= '0' && eax <= '9') || (eax >= 'A' && eax <= 'F') || (eax >= 'a' && eax <= 'f')
			.if (m_fstyle & STYLE_READONLY)
				invoke MessageBeep,MB_ICONHAND
				jmp Exit
			.endif
			.if fSel
				push eax
				invoke DeleteSelection, cpMin, cpMax
				pop eax
			.endif
			.if !m_fOvr || ecx==m_nbytes
				mov ecx,m_cpMin
				shr ecx,1
				.if !CARRY?
					push eax
					invoke InsertChars, ecx, 1
					mov fIns,TRUE
					pop eax
				.endif
			.endif
			.if eax >= 'a'
				and eax,5Fh
			.endif
			.if eax >= 'A'
				sub eax,'A'-10
			.else
				sub eax,'0'
			.endif
			mov edx,0F0h
			mov ecx,m_cpMin
			shr ecx,1
			.if !CARRY?
				shl eax,4
				shr edx,4
			.endif
			push eax
			push edx
			invoke GetChar, ecx
			push eax
			.if !fIns
				invoke SaveUndo, UNDO_CHAROVERWRITE, m_cpMin, eax, 1
			.endif
			pop eax
			pop edx
			pop ecx
			and eax,edx
			or eax,ecx
			mov nChar,eax
			mov ecx,m_cpMin
			shr ecx,1
			invoke PutChar, ecx, addr nChar, 1
			.if fIns
				invoke SaveUndo, UNDO_CHARINSERT, m_cpMin, nChar, 1
			.endif
			mov ecx,m_cpMin
			inc ecx
			inc ecx
			shr ecx,1
			.if ecx > m_nbytes
				mov m_nbytes,ecx
			.endif
			invoke InvalidateLine,m_edta.hwnd,m_cpMin
			invoke InvalidateLine,m_edtb.hwnd,m_cpMin
			mov ecx,m_cpMin
			inc ecx
			mov m_cpMin,ecx
			mov m_cpMax,ecx
			add ecx,32
			.if fIns
				mov edx,ecx
				add edx,4096
				push ecx
				push edx
				invoke InvalidateSelection, 0, ecx, edx
				pop edx
				pop ecx
				invoke InvalidateSelection, 1, ecx, edx
			.else
				push ecx
				invoke InvalidateLine,m_edta.hwnd,ecx
				pop ecx
				invoke InvalidateLine,m_edtb.hwnd,ecx
			.endif
			mov m_fChanged,TRUE
			inc m_nchange
if ?STATE
			invoke InvalidateRect,m_hsta,NULL,TRUE
endif
			invoke DoMove, hWin, fSel, fShift, cpMin, cpMax
			push word ptr EN_CHANGE
			push word ptr m_ID
			pop ecx
			invoke SendMessage, m_hwndParent, WM_COMMAND, ecx, m_hwnd
		.endif
Exit:
		xor eax, eax
		ret
	align 4

OnChar endp

RAHexEdProc proc uses ebx esi edi,hWin:HWND,uMsg:UINT,wParam:WPARAM,lParam:LPARAM

	LOCAL sinf:SCROLLINFO
	LOCAL cpMin:DWORD
	LOCAL cpMax:DWORD
	LOCAL rect:RECT
	LOCAL pt:POINT
	LOCAL fSel:DWORD
	LOCAL fShift:BYTE
	LOCAL fControl:BYTE

	invoke GetWindowLong,hWin,MYWINDATA
	mov ebx,eax
	.if eax
		mov eax,hWin
		.if eax==m_edta.hwnd
			lea esi,m_edta
		.else
			lea esi,m_edtb
		.endif
	.else
		.if uMsg!=WM_CREATE
			jmp ExDef
		.endif
	.endif
	mov eax,uMsg

	.if eax==WM_PAINT

		call SetScroll
		invoke OnPaint, hWin

	.elseif eax==WM_SIZE

		mov fSize,TRUE
		mov eax,m_edta.hwnd
		.if eax==hWin
			invoke GetClientRect,hWin,addr m_edta.rc
			mov eax,m_fstyle
			and eax,STYLE_NOSPLITT or STYLE_NOVSCROLL
			.if eax!=STYLE_NOSPLITT or STYLE_NOVSCROLL
				mov ecx, g_sbwt
				sub m_edta.rc.right,ecx
				invoke MoveWindow,m_edta.hvscroll,m_edta.rc.right,0,g_sbwt,m_edta.rc.bottom,TRUE
			.else
				invoke MoveWindow,m_edta.hvscroll,0,0,0,0,TRUE
			.endif
		.else
			invoke GetClientRect,hWin,addr m_edtb.rc
			mov eax,m_fstyle
			and eax,STYLE_NOSPLITT or STYLE_NOVSCROLL
			.if eax!=STYLE_NOSPLITT or STYLE_NOVSCROLL
				mov ecx, g_sbwt
				sub m_edtb.rc.right,ecx
				mov ecx,m_edtb.rc.bottom
				xor edx,edx
				.if !m_nsplitt
					test m_fstyle,STYLE_NOSPLITT
					.if ZERO?
						mov edx,BTNHT
						sub ecx,edx
					.endif
				.endif
				invoke MoveWindow,m_edtb.hvscroll,m_edtb.rc.right,edx,g_sbwt,ecx,TRUE
			.else
				invoke MoveWindow,m_edtb.hvscroll,0,0,0,0,TRUE
			.endif
		.endif
		invoke SetCaret,hWin

	.elseif eax==WM_KEYDOWN

		invoke OnKeyDown, hWin, wParam, lParam
		jmp Ex

	.elseif eax==WM_CHAR

		invoke OnChar, hWin, wParam, lParam
		jmp Ex

	.elseif eax==WM_VSCROLL

		mov sinf.cbSize,sizeof sinf
		mov sinf.fMask,SIF_ALL
		invoke GetScrollInfo,[esi].HEEDT.hvscroll,SB_CTL,addr sinf
		mov eax,[esi].HEEDT.nline
		mov edx,wParam
		movzx edx,dx
		.if edx==SB_THUMBTRACK || edx==SB_THUMBPOSITION
			mov eax,sinf.nTrackPos
		.elseif edx==SB_LINEDOWN
			inc eax
		.elseif edx==SB_LINEUP
			.if eax
				dec eax
			.endif
		.elseif edx==SB_PAGEDOWN
			add eax,sinf.nPage
		.elseif edx==SB_PAGEUP
			sub eax,sinf.nPage
			jnb @f
			xor eax,eax
		  @@:
		.elseif edx==SB_BOTTOM
			mov eax,sinf.nMax
		.elseif edx==SB_TOP
			xor eax,eax
		.endif
		invoke DoScroll,hWin,m_cpx,eax

	.elseif eax==WM_MOUSEWHEEL

		mov eax,[esi].HEEDT.nline
		.if sdword ptr wParam>0
			sub eax,3
			jnb @f
			xor eax,eax
		  @@:
		.else
			add eax,3
		.endif
		invoke DoScroll,hWin,m_cpx,eax

	.elseif eax==WM_LBUTTONDOWN

		mov eax,m_cpMin
		mov ecx,m_cpMax
		.if eax > ecx
			xchg eax,ecx
		.endif
		mov cpMin,eax
		mov cpMax,ecx
		sub ecx,eax
		mov fSel,ecx
		invoke GetKeyState,VK_SHIFT
		and al,80h
		mov fShift,al
		.if fSel
			invoke InvalidateSelection, 0, cpMin, cpMax
			invoke InvalidateSelection, 1, cpMin, cpMax
		.endif
		invoke SetFocus,hWin
		invoke SetCapture,hWin
		;Get mouse position
		mov eax,lParam
		shr eax,16
		cwde
		mov edx,eax
		mov eax,lParam
		cwde
		mov ecx,m_linenrwt
		add ecx,m_selbarwt
		mov fSelState,1
		.if eax < ecx
			mov fSelState,2
			push edx
;;			invoke SetCursor,hSelCur
			pop edx
			.if !fShift
				push edx
				invoke GetCharFromPos, hWin, 0, edx
				mov m_cpMax,eax
				pop edx
			.endif
			mov eax,[esi].HEEDT.rc.right
			mov fShift,TRUE
		.endif
		invoke GetCharFromPos, hWin, eax, edx
		mov m_cpMin,eax
		mov ecx,m_cpMax
		.if !fShift
			mov m_cpMax,eax
		.elseif eax!=ecx
			inc eax
			and eax,0FFFFFFFEh
			and ecx,0FFFFFFFEh
			mov m_cpMin,eax
			mov m_cpMax,ecx
			.if eax > ecx
				xchg eax,ecx
			.endif
			push eax
			push ecx
			invoke InvalidateSelection, 0, eax, ecx
			pop ecx
			pop eax
			invoke InvalidateSelection, 1, eax, ecx
		.endif

		invoke ScrollCaret, hWin
		invoke SelChange, SEL_TEXT

	.elseif eax==WM_RBUTTONDOWN

		invoke SetFocus,hWin

	.elseif eax==WM_LBUTTONUP

		invoke ReleaseCapture
		xor eax,eax
		mov fSelState,eax
		mov tmr1.hwnd,eax
		mov tmr2.hwnd,eax

		.if (TimerID)
			invoke KillTimer, NULL, TimerID
			mov TimerID, NULL
		.endif

	.elseif eax==WM_LBUTTONDBLCLK

		mov eax,m_cpx
		mov ecx,m_fntinfo.fntwt
		mul ecx
		mov ecx,eax
		mov eax,m_linenrwt
		add eax,m_selbarwt
		mov edx,eax
		add eax,m_dataxp
		sub eax,ecx
		mov ecx,eax
		mov eax,lParam
		cwde
		.if eax < ecx && eax > edx
			mov eax,m_cpMin
			shr eax,1
			.if eax < m_nbytes
				shr eax,4
				shl eax,5
				mov m_cpMin,eax
				add eax,32
				shr eax,1
				.if eax > m_nbytes
					mov eax,m_nbytes
				.endif
				shl eax,1
				mov m_cpMax,eax
				invoke InvalidateSelection, 0, m_cpMin, m_cpMax
				invoke InvalidateSelection, 1, m_cpMin, m_cpMax
				invoke ScrollCaret, hWin
				invoke SelChange, SEL_TEXT
			.endif
		.elseif eax > edx
			mov eax,m_cpMin
			shr eax,1
			.if eax < m_nbytes
				shl eax,1
				mov m_cpMin,eax
				add eax,2
				mov m_cpMax,eax
				invoke InvalidateSelection, 0, m_cpMin, m_cpMax
				invoke InvalidateSelection, 1, m_cpMin, m_cpMax
				invoke ScrollCaret, hWin
				invoke SelChange, SEL_TEXT
			.endif
		.endif

	.elseif eax==WM_MOUSEMOVE

		invoke SendMessage,m_htt,TTM_ACTIVATE,FALSE,0
		invoke SendMessage,m_htt,TTM_ACTIVATE,TRUE,0
		;Get mouse position
		mov eax,lParam
		shr eax,16
		cwde
		mov edx,eax
		mov eax,lParam
		cwde
		.if fSelState

			;Is selection on selbar
			.if fSelState==2
				mov eax,[esi].HEEDT.rc.right
			.endif
			push eax
			push edx
			mov ecx, -1
			.if (sdword ptr edx > [esi].HEEDT.rc.bottom)
				mov ecx, SB_LINEDOWN
				mov edx, WM_VSCROLL
			.elseif (sdword ptr edx < 0)
				mov ecx, SB_LINEUP
				mov edx, WM_VSCROLL
			.elseif (sdword ptr eax > [esi].HEEDT.rc.right)
				mov ecx, SB_LINEDOWN
				mov edx, WM_HSCROLL
			.elseif (sdword ptr eax < 0)
				mov ecx, SB_LINEUP
				mov edx, WM_HSCROLL
			.endif
			.if (ecx != -1)
				mov eax,hWin
				mov tmr1.hwnd,eax
				mov tmr1.umsg,edx
				mov tmr1.wparam,ecx
				mov tmr1.lparam,0
				mov tmr2.hwnd,eax
				mov tmr2.umsg,WM_MOUSEMOVE
				mov eax,wParam
				mov tmr2.wparam,eax
				mov eax,lParam
				mov tmr2.lparam,eax
				.if (!TimerID)
					invoke SetTimer, NULL, 0, 50, offset TimerProc
					mov TimerID,eax
				.endif
			.else
				xor eax,eax
				mov tmr1.hwnd,eax
				mov tmr2.hwnd,eax
			.endif

			pop edx
			pop eax
			invoke GetCharFromPos, hWin, eax, edx
			inc eax
			and eax,0FFFFFFFEh
			mov edx,m_cpMin
			mov m_cpMin,eax
			.if ((eax != edx) || TimerID)
				push eax
				push edx
				invoke InvalidateSelection, 0, edx, eax
				pop edx
				pop eax
				invoke InvalidateSelection, 1, edx, eax
				invoke SetCaret,hWin
				invoke SelChange, SEL_TEXT
			.endif
		.else
			mov ecx,m_linenrwt
			add ecx,m_selbarwt
			.if eax < ecx
;;				invoke SetCursor,hSelCur
			.endif
		.endif

	.elseif eax==WM_SETFOCUS

		mov m_fHasFocus, TRUE
		invoke InvalidateRect,m_edta.hwnd,NULL,FALSE
		invoke InvalidateRect,m_edtb.hwnd,NULL,FALSE

		mov eax,hWin
		mov m_hwndFocus,eax
;;		invoke SetFocus,hWin
		push word ptr EN_SETFOCUS
		push word ptr m_ID
		pop ecx
		invoke SendMessage, m_hwndParent, WM_COMMAND, ecx, m_hwnd

		call SetScroll
		invoke SetCaret,hWin
		invoke SelChange, SEL_TEXT

	.elseif eax==WM_KILLFOCUS

		mov m_fHasFocus, FALSE
		invoke InvalidateRect,m_edta.hwnd,NULL,FALSE
		invoke InvalidateRect,m_edtb.hwnd,NULL,FALSE

		.if !m_fCaretHide
			mov m_fCaretHide,TRUE
			invoke HideCaret, hWin
		.endif
		push word ptr EN_KILLFOCUS
		push word ptr m_ID
		pop ecx
		invoke SendMessage, m_hwndParent, WM_COMMAND, ecx, m_hwnd

	.elseif eax==WM_CREATE

		invoke GetParent,hWin
		invoke GetWindowLong,eax,MYWINDATA
		invoke SetWindowLong,hWin,MYWINDATA,eax

	.elseif eax == WM_GETDLGCODE

 		mov eax,DLGC_WANTARROWS or DLGC_WANTCHARS
		jmp Ex
	.endif
  ExDef:
	invoke DefWindowProc,hWin,uMsg,wParam,lParam
  Ex:
	ret

SetScroll:
ifdef @StackBase
if @ProcStatus and 4
@StackBase = @StackBase + 4
endif
endif
	invoke GetClientRect,hWin,addr rect
	mov sinf.cbSize,sizeof sinf
	mov sinf.fMask,SIF_ALL or SIF_DISABLENOSCROLL
	mov sinf.nMin,0
	mov eax,m_nbytes
	shr eax,4
	inc eax
	mov ecx,m_fntinfo.fntht
	mul ecx
	.if eax < rect.bottom
		mov eax,rect.bottom
	.endif
	xor edx,edx
	div ecx
	dec eax
	mov sinf.nMax,eax
	mov eax,rect.bottom
	xor edx,edx
	div ecx
	mov sinf.nPage,eax
	mov eax,[esi].HEEDT.nline
	mov sinf.nPos,eax
	invoke SetScrollInfo,[esi].HEEDT.hvscroll,SB_CTL,addr sinf,TRUE
	mov sinf.cbSize,sizeof sinf
	mov sinf.fMask,SIF_ALL or SIF_DISABLENOSCROLL
	mov sinf.nMin,0
	mov eax,rect.right
	sub eax,g_sbwt
	sub eax,m_selbarwt
	sub eax,m_linenrwt
	mov ecx,m_fntinfo.fntwt
	xor edx,edx
	div ecx
	mov sinf.nPage,eax
	mov eax,m_fntinfo.fntwt
	shl eax,4
	add eax,m_asciixp
	mov ecx,m_fntinfo.fntwt

	mov eax,m_asciixp
	add eax,m_asciiwt

	xor edx,edx
	div ecx
	mov sinf.nMax,eax
	mov eax,m_cpx
	mov sinf.nPos,eax
	invoke SetScrollInfo,m_hhscroll,SB_CTL,addr sinf,TRUE
	retn
ifdef @StackBase
if @ProcStatus and 4
@StackBase = @StackBase - 4
endif
endif
	align 4

RAHexEdProc endp

ResetMem proc
	.if (m_hundo)
		invoke free, m_hundo
	.endif
	invoke	malloc, MAXUNDOMEM
	mov 	m_hundo,eax
	mov 	m_rpundo,0
	mov 	m_cbundo,MAXUNDOMEM
	.if (m_hmem)
if ?USEGLOBAL
		invoke	GlobalFree, m_hmem
else
		invoke free, m_hmem
endif
	.endif
if ?USEGLOBAL
	invoke	GlobalAlloc, GMEM_MOVEABLE or GMEM_ZEROINIT,MAXCHARMEM
else
	invoke	malloc,MAXCHARMEM
endif
	mov 	m_hmem,eax
	mov 	m_nbytes,0
	mov 	m_nsize,MAXCHARMEM
	mov 	m_edta.nline,0
	mov 	m_edtb.nline,0
	mov 	m_cpx,0
	mov 	m_cpMin,0
	mov 	m_cpMax,0
	mov 	m_fChanged,FALSE
	ret
	align 4
ResetMem endp

;--- WM_CREATE

OnCreate proc hWin:HWND

	LOCAL	ti:TOOLINFO

;---------------------------------------- Allocate memory for the hex edit control
		invoke malloc,sizeof CHexEdit
		mov ebx,eax
;---------------------------------------- Save the pointer
		invoke SetWindowLong,hWin,MYWINDATA,ebx
		mov eax,hWin
		mov m_hwnd,eax
		invoke GetParent,eax
		mov m_hwndParent,eax
		invoke GetWindowLong,m_hwnd,GWL_ID
		mov m_ID,eax
		invoke ResetMem
		mov m_fOvr,TRUE
if ?LINENUMBER
		mov m_nlinenrwt, LNRWT
endif
		mov m_selbarwt, SELWT
		mov m_fntinfo.fntht,10

		invoke GetSysColor, COLOR_WINDOW
		mov m_clr.bckcol,eax

		mov m_clr.adrtxtcol,DEFADRTXTCOLOR

;		mov m_clr.dtatxtcol,DEFDTATXTCOLOR
		invoke GetSysColor, COLOR_WINDOWTEXT 
		mov m_clr.dtatxtcol,eax

		mov m_clr.asctxtcol,DEFASCTXTCOLOR

;		mov m_clr.selbckcol,DEFSELBCKCOLOR
		invoke GetSysColor, COLOR_HIGHLIGHT 
		mov m_clr.selbckcol, eax

;		mov m_clr.seltxtcol,DEFSELTXTCOLOR
		invoke GetSysColor, COLOR_HIGHLIGHTTEXT 
		mov m_clr.seltxtcol,eax

		mov m_clr.selbarbck,DEFSELBARCOLOR
		mov m_clr.selbarpen,DEFSELBARPEN
		mov m_clr.lnrcol,DEFLNRCOLOR
		invoke CreateBrushes

		invoke CreateWindowEx,NULL,addr szToolTips,NULL,TTS_ALWAYSTIP,CW_USEDEFAULT,CW_USEDEFAULT,CW_USEDEFAULT,CW_USEDEFAULT,NULL,NULL,hInstance,NULL
		mov m_htt,eax

		invoke CreateWindowEx,1,addr szStatic,NULL,SS_NOTIFY or WS_CHILD or WS_VISIBLE,0,0,0,0,hWin,-1,hInstance,0
		mov m_hsbtn,eax
		mov edx,offset szSplitterBar
		call SetToolTip
		invoke SetWindowLong,m_hsbtn,GWL_WNDPROC,addr SplittBtnProc
		mov OldSplittBtnProc,eax

		invoke CreateWindowEx,NULL,addr szEditClassName,NULL,WS_CLIPSIBLINGS or WS_CHILD or WS_VISIBLE,0,0,0,0,hWin,NULL,hInstance,0
		mov m_edta.hwnd,eax
		invoke CreateWindowEx,NULL,addr szScrollBar,NULL,WS_CHILD or WS_VISIBLE or SBS_VERT,0,0,0,0,m_edta.hwnd,NULL,hInstance,0
		mov m_edta.hvscroll,eax

		invoke CreateWindowEx,NULL,addr szEditClassName,NULL,WS_CLIPSIBLINGS or WS_CHILD or WS_VISIBLE,0,0,0,0,hWin,NULL,hInstance,0
		mov m_edtb.hwnd,eax
		invoke CreateWindowEx,NULL,addr szScrollBar,NULL,WS_CHILD or WS_VISIBLE or SBS_VERT,0,0,0,0,m_edtb.hwnd,NULL,hInstance,0
		mov m_edtb.hvscroll,eax

		invoke CreateWindowEx,NULL,addr szScrollBar,NULL,WS_CHILD or WS_VISIBLE or SBS_HORZ,0,0,0,0,hWin,NULL,hInstance,0
		mov m_hhscroll,eax
		invoke CreateWindowEx,NULL,addr szScrollBar,NULL,WS_CHILD or WS_VISIBLE or SBS_SIZEGRIP,0,0,0,0,hWin,NULL,hInstance,0
		mov m_hgrip,eax

		invoke CreateWindowEx,NULL,addr szStatic,NULL,WS_CHILD or WS_VISIBLE,0,0,0,0,hWin,NULL,hInstance,0
		mov m_hnogrip,eax

if ?STATE
		invoke CreateWindowEx,NULL,addr szStatic,NULL,SS_NOTIFY or WS_CHILD or WS_VISIBLE,0,0,0,0,hWin,NULL,hInstance,0
		mov m_hsta,eax
		mov edx,offset szChanged
		call SetToolTip
		invoke SetWindowLong,m_hsta,GWL_USERDATA, ebx
		invoke SetWindowLong,m_hsta,GWL_WNDPROC, addr StateProc
		mov OldStateProc,eax
endif
if ?LINENUMBER
		invoke CreateWindowEx,NULL,addr szButton,NULL,BS_BITMAP or BS_PUSHLIKE or BS_AUTOCHECKBOX or WS_CHILD or WS_VISIBLE,0,0,0,0,hWin,-2,hInstance,0
		mov m_hlin,eax
		mov edx,offset szLineNumber
		call SetToolTip
		invoke SendMessage,m_hlin,BM_SETIMAGE,IMAGE_BITMAP,hBmpLnr
endif
if ?STATE
		invoke SetWindowPos,m_hsta,HWND_TOP,0,0,0,0,SWP_NOSIZE or SWP_NOSIZE;SWP_NOREPOSITION
endif
		invoke SetWindowPos,m_hsbtn,HWND_TOP,0,0,0,0,SWP_NOSIZE or SWP_NOSIZE;SWP_NOREPOSITION
		ret
SetToolTip:
ifdef @StackBase
	if @ProcStatus and 4
		@StackBase = @StackBase + 4
	endif
endif
	mov ti.cbSize,sizeof TOOLINFO
	mov ti.uFlags,TTF_IDISHWND or TTF_SUBCLASS
	mov ti.hwnd,0
	mov ti.uId,eax
	mov ti.hinst,0
	mov ti.lpszText,edx
	invoke SendMessage,m_htt,TTM_ADDTOOL,NULL,addr ti
ifdef @StackBase
	if @ProcStatus and 4
		@StackBase = @StackBase - 4
	endif
endif
	retn
	align 4

OnCreate endp

MyCreateFont proc wParam:WPARAM

local hDC:HDC
local logfont:LOGFONT
local tm:TEXTMETRIC

if 0
		invoke GetDC, m_hwnd
		mov hDC, eax
		invoke SelectObject, hDC, wParam
		push eax
		invoke GetTextMetrics, hDC, addr tm
		pop eax
		invoke SelectObject, hDC, eax

		sub esp, 128
		mov edx, esp
		invoke wsprintf, edx, CStr("tm.height=%u, tm.avewidth=%u",13,10), tm.tmHeight, tm.tmAveCharWidth
		invoke OutputDebugString, esp
		add esp, 128

		invoke ReleaseDC, m_hwnd, hDC
endif

	mov ecx, g_dwFontWidth
	.if (!g_hFont || (ecx != g_OldFontWidth))
		.if (g_hFont)
			invoke DeleteObject, g_hFont
			mov g_hFont, NULL
		.endif
		invoke RtlZeroMemory, addr logfont, sizeof LOGFONT
		mov eax, g_dwFontWidth
		mov g_OldFontWidth, eax
		mov logfont.lfWidth, eax
		add eax, eax
		mov logfont.lfHeight, eax
		mov logfont.lfWeight, FW_NORMAL 
		mov logfont.lfCharSet, DEFAULT_CHARSET
		mov logfont.lfPitchAndFamily, FIXED_PITCH
		invoke CreateFontIndirect, addr logfont
		mov g_hFont, eax
	.endif
	mov eax,g_hFont
	ret
	align 4

MyCreateFont endp

;--- The hex edit controls callback (WndProc).

RAWndProc proc uses ebx esi edi,hWin:HWND,uMsg:UINT,wParam:WPARAM,lParam:LPARAM

	LOCAL rect:RECT
	LOCAL pt:POINT
	LOCAL sinf:SCROLLINFO
	LOCAL chrg:CHARRANGE
	LOCAL cpMin:DWORD
	LOCAL cpMax:DWORD
	LOCAL fSel:DWORD
ifdef _DEBUG
	LOCAL this_:ptr CHexEdit
endif

	;Get memory pointer
	invoke GetWindowLong,hWin,MYWINDATA
	.if eax
		mov ebx,eax
ifdef _DEBUG
		mov this_, eax
endif
		mov eax,m_hwndFocus
		.if eax==m_edta.hwnd
			lea esi,m_edta
		.else
			lea esi,m_edtb
		.endif
	.else
		.if uMsg!=WM_CREATE
			jmp ExDef
		.endif
	.endif
	mov eax,uMsg
	.if eax==WM_CREATE

		invoke OnCreate, hWin

	.elseif eax==WM_COMMAND
		movzx eax,word ptr wParam+0
if ?LINENUMBER
		.if eax==-2
			.if m_linenrwt
				mov m_linenrwt,0
			.else
				mov eax,m_nlinenrwt
				mov m_linenrwt,eax
			.endif
			invoke SetFocus,hWin
;;			invoke InvalidateRect,m_edta.hwnd,NULL,FALSE
;;			invoke InvalidateRect,m_edtb.hwnd,NULL,FALSE
		.endif
endif
	.elseif eax==HEM_SETFONT
		;wParam=nLineSpacing
		;lParam=lpRAFONT
		mov eax,wParam
		mov m_fntinfo.linespace,eax
		invoke SetFont, lParam
		invoke GetFocus
		.if eax==m_hwndFocus && eax
			invoke SetFocus,m_hwndFocus
			invoke ScrollCaret, m_hwndFocus
		.endif
		invoke InvalidateRect,m_edta.hwnd,NULL,FALSE
		invoke InvalidateRect,m_edtb.hwnd,NULL,FALSE
		xor eax,eax
		jmp Ex
	.elseif eax==HEM_GETFONT
		;wParam=0
		;lParam=lpRAFONT
		mov edx,lParam
		mov eax,m_fnt.hFont
		mov [edx].HEFONT.hFont,eax
		mov eax,m_fnt.hLnrFont
		mov [edx].HEFONT.hLnrFont,eax
		mov eax,m_fntinfo.linespace
		jmp Ex
if ?COLOR
	.elseif eax==HEM_SETCOLOR
		;wParam=0
		;lParam=lpRACOLOR
		mov edx,lParam
		mov eax,[edx].HECOLOR.bckcol
		mov m_clr.bckcol,eax
		mov eax,[edx].HECOLOR.adrtxtcol
		mov m_clr.adrtxtcol,eax
		mov eax,[edx].HECOLOR.dtatxtcol
		mov m_clr.dtatxtcol,eax
		mov eax,[edx].HECOLOR.asctxtcol
		mov m_clr.asctxtcol,eax
		mov eax,[edx].HECOLOR.selbckcol
		.if eax==[edx].HECOLOR.bckcol
			xor eax,03F3F3Fh
		.endif
		mov m_clr.selbckcol,eax
		mov eax,[edx].HECOLOR.seltxtcol
		mov m_clr.seltxtcol,eax
		mov eax,[edx].HECOLOR.selbarbck
		mov m_clr.selbarbck,eax
		mov eax,[edx].HECOLOR.selbarpen
		mov m_clr.selbarpen,eax
		mov eax,[edx].HECOLOR.lnrcol
		mov m_clr.lnrcol,eax
		invoke CreateBrushes
		invoke InvalidateRect,m_edta.hwnd,NULL,FALSE
		invoke InvalidateRect,m_edtb.hwnd,NULL,FALSE
if ?STATE
		invoke InvalidateRect,m_hsta,NULL,FALSE
endif
		xor eax,eax
		jmp Ex
	.elseif eax==HEM_GETCOLOR
		;wParam=0
		;lParam=lpRACOLOR
		mov edx,lParam
		mov eax,m_clr.bckcol
		mov [edx].HECOLOR.bckcol,eax
		mov eax,m_clr.adrtxtcol
		mov [edx].HECOLOR.adrtxtcol,eax
		mov eax,m_clr.dtatxtcol
		mov [edx].HECOLOR.dtatxtcol,eax
		mov eax,m_clr.asctxtcol
		mov [edx].HECOLOR.asctxtcol,eax
		mov eax,m_clr.selbckcol
		mov [edx].HECOLOR.selbckcol,eax
		mov eax,m_clr.seltxtcol
		mov [edx].HECOLOR.seltxtcol,eax
		mov eax,m_clr.selbarbck
		mov [edx].HECOLOR.selbarbck,eax
		mov eax,m_clr.selbarpen
		mov [edx].HECOLOR.selbarpen,eax
		mov eax,m_clr.lnrcol
		mov [edx].HECOLOR.lnrcol,eax
		xor eax,eax
		jmp Ex
endif
if ?VCENTER
	.elseif eax==HEM_VCENTER
		;wParam=0
		;lParam=0
		mov eax,[esi].HEEDT.rc.bottom
		mov ecx,m_fntinfo.fntht
		xor edx,edx
		div ecx
		shr eax,1
		mov ecx,m_cpMin
		shr ecx,5
		sub ecx,eax
		jnb @f
		xor ecx,ecx
	  @@:
		mov edx,ecx
		add edx,eax
		add edx,eax
		shl edx,4
		.if edx > m_nbytes
			mov edx,m_nbytes
			shr edx,4
			inc edx
			shl eax,1
			sub edx,eax
			jnb @f
			xor edx,edx
		  @@:
			mov ecx,edx
		.endif
		mov [esi].HEEDT.nline,ecx
		invoke InvalidateRect,[esi].HEEDT.hwnd,NULL,FALSE
		invoke SetCaret,[esi].HEEDT.hwnd
		xor eax,eax
		jmp Ex
	.elseif eax==HEM_REPAINT
		;wParam=0
		;lParam=0
		invoke InvalidateRect,m_edta.hwnd,NULL,FALSE
		invoke InvalidateRect,m_edtb.hwnd,NULL,FALSE
		invoke SetCaret,[esi].HEEDT.hwnd
		xor eax,eax
		jmp Ex
endif
if ?BOOKMARK
	.elseif eax==HEM_ANYBOOKMARKS
		;wParam=0
		;lParam=0
		mov eax,Bookmark.hWin
		.if eax
			mov eax,TRUE
		.endif
		jmp Ex
	.elseif eax==HEM_TOGGLEBOOKMARK
		;wParam=nLine
		;lParam=0
		mov edi,wParam
		invoke ToggleBookmark,hWin,edi
		shl edi,5
		invoke InvalidateLine,m_edta.hwnd,edi
		invoke InvalidateLine,m_edtb.hwnd,edi
		xor eax,eax
		jmp Ex
	.elseif eax==HEM_CLEARBOOKMARKS
		;wParam=0
		;lParam=0
		invoke ClearBookmarks
		xor eax,eax
		jmp Ex
	.elseif eax==HEM_NEXTBOOKMARK
		;wParam=0
		;lParam=lpHEBMK
		invoke NxtBookmark,lParam
		jmp Ex
	.elseif eax==HEM_PREVIOUSBOOKMARK
		;wParam=0
		;lParam=lpHEBMK
		invoke PrvBookmark,lParam
		jmp Ex
endif
	.elseif eax==HEM_SELBARWIDTH
		;wParam=nWidth
		;lParam=0
		mov eax,wParam
		mov m_selbarwt,eax
		invoke InvalidateRect,m_edta.hwnd,NULL,FALSE
		invoke InvalidateRect,m_edtb.hwnd,NULL,FALSE
		invoke SetCaret,m_hwndFocus
		xor eax,eax
		jmp Ex
if ?LINENUMBER
	.elseif eax==HEM_LINENUMBERWIDTH
		;wParam=nWidth
		;lParam=0
		mov eax,wParam
		mov m_nlinenrwt,eax
		.if m_linenrwt
			mov m_linenrwt,eax
			invoke InvalidateRect,m_edta.hwnd,NULL,FALSE
			invoke InvalidateRect,m_edtb.hwnd,NULL,FALSE
			invoke SetCaret,m_focus
		.endif
		xor eax,eax
		jmp Ex
endif
	.elseif eax==HEM_GETSIZE

		mov eax, m_nbytes
		jmp Ex

	.elseif eax == HEM_RESET

		invoke ResetMem
		jmp streamin1

	.elseif eax==EM_STREAMIN
		;wParam=SF_TEXT
		;lParam=lpStream

		invoke SetCursor, g_hCsrWait
		push eax
		invoke StreamIn, lParam
		pop eax
		invoke SetCursor,eax
streamin1:
		xor eax,eax
		mov m_edta.nline,eax
		mov m_edtb.nline,eax
		mov m_cpMin,eax
		mov m_cpMax,eax
		mov m_cpx,eax
		invoke InvalidateRect,m_edta.hwnd,NULL,FALSE
		invoke InvalidateRect,m_edtb.hwnd,NULL,FALSE
		invoke SelChange, SEL_TEXT
		xor eax,eax
		jmp Ex

	.elseif eax==EM_STREAMOUT
		;wParam=SF_TEXT
		;lParam=lpStream

		invoke SetCursor, g_hCsrWait
		push eax
		invoke StreamOut, lParam
		pop eax
		invoke SetCursor,eax
		xor eax,eax
		jmp Ex

	.elseif eax==EM_GETMODIFY
		;wParam=0
		;lParam=0
		movzx eax,m_fChanged
		jmp Ex
	.elseif eax==EM_SETMODIFY
		;wParam=TRUE/FALSE
		;lParam=0
		mov eax,wParam
		mov m_fChanged,al
if ?STATE
		invoke InvalidateRect,m_hsta,NULL,TRUE
endif
		xor eax,eax
		jmp Ex
	.elseif eax==EM_EXGETSEL
		;wParam=0
		;lParam=lpCHARRANGE
		mov edx,lParam
		mov eax,m_cpMin
		mov ecx,m_cpMax
		.if eax > ecx
			xchg eax,ecx
		.endif
		mov [edx].CHARRANGE.cpMin,eax
		mov [edx].CHARRANGE.cpMax,ecx
		xor eax,eax
		jmp Ex
	.elseif eax==EM_EXSETSEL
		;wParam=0
		;lParam=lpCHARRANGE
		mov edx,lParam
		mov eax,[edx].CHARRANGE.cpMin
		mov ecx,[edx].CHARRANGE.cpMax
		.if eax > ecx
			xchg eax,ecx
		.endif
		mov edx,m_nbytes
		shl edx,1
		.if eax > edx
			mov eax,edx
		.endif
		.if ecx > edx
			mov ecx,edx
		.endif
		mov m_cpMin,eax
		mov m_cpMax,ecx
		invoke SetCaret,m_hwndFocus
		invoke InvalidateRect,m_edta.hwnd,NULL,FALSE
		invoke InvalidateRect,m_edtb.hwnd,NULL,FALSE
		invoke SelChange, SEL_TEXT
		xor eax,eax
		jmp Ex
	.elseif eax==EM_SCROLLCARET
		;wParam=0
		;lParam=0
		invoke ScrollCaret, m_hwndFocus
		xor eax,eax
		jmp Ex
	.elseif eax==EM_REPLACESEL
		;wParam=fCanUndo
		;lParam=lpHexStr
		mov eax,m_cpMin
		mov ecx,m_cpMax
		.if eax > ecx
			xchg eax,ecx
		.endif
		mov cpMin,eax
		mov cpMax,ecx
		sub ecx,eax
		mov fSel,ecx
		.if fSel
			test m_fstyle,STYLE_READONLY
			jne ErrBeep
			invoke DeleteSelection, cpMin, cpMax
		.endif
		mov edx,lParam
		.if edx
			test m_fstyle,STYLE_READONLY
			jne ErrBeep
			mov eax,cpMin
			shr eax,1
			test wParam,FR_HEX
			.if ZERO?
				invoke InsertAsciiString, eax, edx
			.else
				invoke InsertHexString, eax, edx
			.endif
			mov m_fChanged,TRUE
			inc m_nchange
		.endif
if ?STATE
		invoke InvalidateRect,m_hsta,NULL,TRUE
endif
		xor eax,eax
		jmp Ex
if ?FINDTEXT
	.elseif eax==EM_FINDTEXTEX
		;wParam=fuFlags
		;lParam=lpFINDTEXTEX 
		invoke FindTextEx, wParam, lParam
		jmp Ex
endif
	.elseif eax==EM_CANPASTE
		;wParam=CF_TEXT
		;lParam=0
		invoke IsClipboardFormatAvailable,wParam
		jmp Ex
	.elseif eax==EM_CANREDO
		;wParam=0
		;lParam=0
		mov edx,m_hundo
		add edx,m_rpundo
		mov eax,[edx].HEUNDO.cb
		.if eax
			mov eax,TRUE
		.endif
		jmp Ex
	.elseif eax==EM_REDO
		;wParam=0
		;lParam=0
		invoke EditRedo, [esi].HEEDT.hwnd
		jmp Ex
	.elseif eax==EM_CANUNDO
		;wParam=0
		;lParam=0
		mov eax,m_rpundo
		.if eax
			mov eax,TRUE
		.endif
		jmp Ex
	.elseif eax==EM_UNDO
		;wParam=0
		;lParam=0
		invoke EditUndo, [esi].HEEDT.hwnd
		jmp Ex
	.elseif eax==EM_EMPTYUNDOBUFFER
		;wParam=0
		;lParam=0
		mov edi,m_hundo
		mov ecx,m_cbundo
		xor eax,eax
		mov m_rpundo,eax
		rep stosb
		jmp Ex
	.elseif eax==WM_CLEAR
		;wParam=0
		;lParam=0
		mov eax,m_cpMin
		mov ecx,m_cpMax
		.if eax > ecx
			xchg eax,ecx
		.endif
		mov cpMin,eax
		mov cpMax,ecx
		sub ecx,eax
		mov fSel,ecx
		.if fSel
			test m_fstyle,STYLE_READONLY
			jne ErrBeep
			invoke DeleteSelection, cpMin, cpMax
		.endif
		xor eax,eax
		jmp Ex
	.elseif eax==WM_CUT
		;wParam=0
		;lParam=0
		invoke SendMessage,hWin,WM_COPY,0,0
		invoke SendMessage,hWin,WM_CLEAR,0,0
		xor eax,eax
		jmp Ex
	.elseif eax==WM_COPY
		;wParam=0
		;lParam=0
		invoke EditCopy
		xor eax,eax
		jmp Ex
	.elseif eax==WM_PASTE
		;wParam=0
		;lParam=0
		invoke EditPaste
		xor eax,eax
		jmp Ex
	.elseif eax==WM_MOUSEMOVE
		;Get mouse position
		mov eax,lParam
		mov edx,eax
		and eax,0FFFFh
		shr edx,16
		.if m_fresize==1
			mov m_fresize,2
		.elseif m_fresize==2
			mov eax,lParam
			shr eax,16
			cwde
			sub eax,2
			mov edx,m_rc.bottom
			sub edx, MINSIZE
			.if sdword ptr eax<0
				xor eax,eax
			.elseif sdword ptr eax>edx
				mov eax,edx
			.endif
			.if eax
				shl eax,10
				mov ecx,m_rc.bottom
				xor edx,edx
				div ecx
			.endif
			mov m_fsplitt,eax
			call SizeIt
			invoke UpdateWindow,m_hsbtn
			invoke UpdateWindow,m_edta.hwnd
			invoke UpdateWindow,m_edtb.hwnd
		.endif
		xor eax,eax
		jmp Ex
	.elseif eax==WM_LBUTTONUP
		.if m_fresize
			mov m_fresize,0
			invoke ReleaseCapture
			invoke ClipCursor, NULL
			call SizeIt
		.endif
		xor eax,eax
		jmp Ex
	.elseif eax==WM_SIZE
		invoke GetWindowLong,hWin,GWL_STYLE
		mov m_fstyle,eax
		invoke GetClientRect,hWin,addr m_rc
		mov eax,m_fstyle
		and eax,STYLE_NOLINENUMBER or STYLE_NOHSCROLL
		.if eax!=STYLE_NOLINENUMBER or STYLE_NOHSCROLL
			mov ecx, g_sbwt
			sub m_rc.bottom,ecx
		.endif
		call SizeIt
		push edi
		xor edi,edi
if ?LINENUMBER
		test m_fstyle,STYLE_NOLINENUMBER
		.if ZERO?
			invoke MoveWindow,m_hlin,0,m_rc.bottom,BTNWT,g_sbwt,TRUE
			add edi,BTNWT
		.else
			invoke MoveWindow,m_hlin,0,0,0,0,TRUE
		.endif
endif
		mov eax,m_fstyle
		and eax,STYLE_NOLINENUMBER or STYLE_NOHSCROLL
		.if eax!=STYLE_NOLINENUMBER or STYLE_NOHSCROLL
			mov ecx,m_rc.right
			sub ecx,edi
			mov eax,m_fstyle
			and eax,STYLE_NOSPLITT or STYLE_NOVSCROLL
			.if eax != STYLE_NOSPLITT or STYLE_NOVSCROLL
				sub ecx,g_sbwt
				invoke MoveWindow,m_hhscroll,edi,m_rc.bottom,ecx,g_sbwt,TRUE
				test m_fstyle,STYLE_NOSIZEGRIP
				jne @f
				invoke GetWindowLong,m_hwndParent,GWL_STYLE
				test eax,WS_MAXIMIZE
				jne @f
				invoke GetClientRect,m_hwndParent,addr rect
				invoke ClientToScreen,m_hwndParent,addr rect.right
				mov eax,m_rc.bottom
				add eax,g_sbwt
				mov rect.top,eax
				invoke ClientToScreen,hWin,addr rect
				mov eax,rect.bottom
				sub eax,rect.top
				.if eax < 5
					mov eax,m_rc.right
					sub eax,g_sbwt
					invoke MoveWindow,m_hgrip,eax,m_rc.bottom,g_sbwt,g_sbwt,TRUE
					invoke MoveWindow,m_hnogrip,0,0,0,0,TRUE
				.else
				  @@:
					mov eax,m_rc.right
					sub eax,g_sbwt
					invoke MoveWindow,m_hnogrip,eax,m_rc.bottom,g_sbwt,g_sbwt,TRUE
					invoke MoveWindow,m_hgrip,0,0,0,0,TRUE
				.endif
			.else
				invoke MoveWindow,m_hhscroll,edi,m_rc.bottom,ecx,g_sbwt,TRUE
				invoke MoveWindow,m_hnogrip,0,0,0,0,TRUE
				invoke MoveWindow,m_hgrip,0,0,0,0,TRUE
			.endif
		.else
			invoke MoveWindow,m_hhscroll,0,0,0,0,TRUE
			invoke MoveWindow,m_hgrip,0,0,0,0,TRUE
			invoke MoveWindow,m_hnogrip,0,0,0,0,TRUE
		.endif
		pop edi
	.elseif eax==WM_HSCROLL
		mov sinf.cbSize,sizeof sinf
		mov sinf.fMask,SIF_ALL
		invoke GetScrollInfo,lParam,SB_CTL,addr sinf
		mov edi,m_cpx
		mov eax,wParam
		movzx eax,ax
		.if eax==SB_THUMBTRACK || eax==SB_THUMBPOSITION
			mov edi,sinf.nTrackPos
		.elseif eax==SB_LINELEFT
			dec edi
			jns @f
			xor edi,edi
		  @@:
		.elseif eax==SB_LINERIGHT
			add edi,1
		.elseif eax==SB_PAGELEFT
			sub edi,sinf.nPage
			jnb @f
			xor edi,edi
		  @@:
		.elseif eax==SB_PAGERIGHT
			add edi,sinf.nPage
		.endif
		invoke DoScroll,m_edtb.hwnd,edi,m_edtb.nline
	.elseif eax==WM_SETFOCUS

		mov eax,m_hwndFocus
		.if eax==m_edta.hwnd
			invoke SetFocus,eax
		.else
			invoke SetFocus,m_edtb.hwnd
		.endif

	.elseif eax==WM_DESTROY

if ?BOOKMARK
		invoke ClearWinBookmarks,hWin
endif
		invoke SetWindowLong,hWin,MYWINDATA,0
		invoke DestroyWindow,m_htt
		invoke DestroyWindow,m_edta.hwnd
		invoke DestroyWindow,m_edtb.hwnd
		invoke DestroyBrushes
		;Free memory
if ?USEGLOBAL
		invoke GlobalFree,m_hmem
else
		invoke free, m_hmem
endif
		invoke free, m_hundo
		invoke	free, __this

	.elseif eax==WM_SETFONT

		invoke MyCreateFont, wParam
		mov m_fnt.hFont,eax
		mov m_fnt.hLnrFont,eax
		invoke SetFont, addr m_fnt

	.elseif eax==WM_STYLECHANGED

		invoke GetWindowLong, hWin, GWL_STYLE
		mov m_fstyle,eax
		invoke SetFont, addr m_fnt

	.elseif eax==WM_ENABLE
		mov ecx, wParam
		mov m_bEnabled, cl

	.elseif eax==WM_CONTEXTMENU
		invoke ShowContextMenu
	.endif

  ExDef:
	invoke DefWindowProc,hWin,uMsg,wParam,lParam
  Ex:
	ret

ErrBeep:
	invoke MessageBeep,MB_ICONHAND
	xor eax,eax
	jmp Ex


SizeIt:
ifdef @StackBase
	if @ProcStatus and 4
		@StackBase = @StackBase + 4
	endif
endif
	mov eax,m_fsplitt
	.if eax
		mov ecx,eax
		mov eax,m_rc.bottom
		mul ecx
		shr eax,10
	.endif
	mov m_nsplitt,eax
	.if m_nsplitt
		invoke MoveWindow,m_edta.hwnd,0,0,m_rc.right,m_nsplitt,TRUE
		mov ecx,m_nsplitt
		add ecx,BTNHT
		mov eax,m_rc.bottom
		sub eax,ecx
		push eax
		push ecx
		pop ecx
		pop eax
		invoke MoveWindow,m_edtb.hwnd,0,ecx,m_rc.right,eax,TRUE
		invoke MoveWindow,m_hsbtn,0,m_nsplitt,m_rc.right,BTNHT,TRUE
	.else
		invoke MoveWindow,m_edta.hwnd,0,0,0,0,TRUE
		mov ecx,m_rc.right
		.if sdword ptr ecx < 0
			xor ecx,ecx
		.endif
		mov edx,m_rc.bottom
		.if sdword ptr edx < 0
			xor edx,edx
		.endif
		invoke MoveWindow,m_edtb.hwnd,0,0,ecx,edx,TRUE
		mov eax,m_rc.right
		sub eax,g_sbwt
		test m_fstyle,STYLE_NOSPLITT
		.if ZERO?
			invoke MoveWindow,m_hsbtn,eax,0,g_sbwt,BTNHT,TRUE
		.else
			invoke MoveWindow,m_hsbtn,0,0,0,0,TRUE
		.endif
if ?STATE
		test m_fstyle,STYLE_NOSTATE
		.if ZERO?
			invoke MoveWindow,m_hsta,0,0,4,4,TRUE
		.else
			invoke MoveWindow,m_hsta,0,0,0,0,TRUE
		.endif
endif
	.endif
	retn
ifdef @StackBase
	if @ProcStatus and 4
		@StackBase = @StackBase - 4
	endif
endif
	align 4

RAWndProc endp


Init@CHexEdit proc public hInst:HINSTANCE


	.if (!hInstance)
		push hInst
		pop hInstance
		invoke CreateClass, hInstance

		invoke LoadCursor, hInstance, IDC_CURSOR1
		mov g_hHSCur,eax

		invoke LoadMenu, hInstance, IDR_MENU5 
		mov g_hMenu, eax

		invoke GetSystemMetrics, SM_CXVSCROLL
		mov g_sbwt, eax

		invoke RegisterClipboardFormat, CStr("CHexEdit")
		mov g_dwMyCBFormat, eax

	.endif
	ret
	align 4

Init@CHexEdit endp

Deinit@CHexEdit proc public

	.if (g_hFont)
		invoke DeleteObject, g_hFont
	.endif
	ret
	align 4
Deinit@CHexEdit endp

End
