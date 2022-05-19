
;*** implements a splitt button class (subclassed static control)
;*** IMPORTANT: keep this file independant from COMView

	.386
	.model flat,stdcall
	option casemap :none   ; case sensitive
	option proc:private

	.nolist
	.nocref
WIN32_LEAN_AND_MEAN	equ 1
	include windows.inc
	include commctrl.inc
	include windowsx.inc

	include CSplittButton.inc
	include debugout.inc

	include rsrc.inc

	.list
	.cref

?RESIZEATONCE	equ 1
MINSIZE			equ 16

externdef	g_hInstance:HINSTANCE
malloc		proto :DWORD
free		proto :HANDLE

CSplittButton struct
hWndParent HWND ?
hWnd1	HWND ?
hWnd2	HWND ?
CSplittButton ends

	.data

g_OldSplittBtnProc DWORD 0
g_hHSCsr	HCURSOR NULL
;;g_hVSCsr	HCURSOR NULL

	.code

if ?RESIZEATONCE eq 0

DrawResizeLine proc uses esi hWnd:HWND , yPos:DWORD
        
local	hBrushOld:HBRUSH
local	dwGripSize:DWORD
local	hdc:HDC
local	rect:RECT
local	hBitmap:HBITMAP
local	dwYChild:DWORD
local	pattern[8]:WORD
	
	.data
g_hHalftoneBrush	HBRUSH 0
g_iOldResizeLine	DWORD 0
	.code

	invoke GetWindowRect,hWnd,addr rect
	mov eax, rect.bottom
	sub eax, rect.top
	mov dwGripSize, eax
	invoke GetParent,hWnd
	mov hWnd, eax
	invoke ScreenToClient, hWnd, addr rect
	invoke ScreenToClient, hWnd, addr rect.right

	.if (!g_hHalftoneBrush)
		lea ecx, pattern
		xor esi, esi
		mov eax, 5555h
		.while (esi < 8)
			mov [ecx+esi*2], eax
			xor eax, 0FFFFh
			inc esi
		.endw
		invoke CreateBitmap, 8,8,1,1,addr pattern
		mov hBitmap,eax
		invoke CreatePatternBrush, hBitmap
		mov g_hHalftoneBrush, eax
		invoke DeleteObject, hBitmap
	.endif

	invoke GetDC, hWnd
	mov hdc, eax
	invoke SelectObject, hdc, g_hHalftoneBrush
	mov hBrushOld, eax
	mov eax, yPos
	add eax, rect.top
	mov ecx, rect.right
	sub ecx, rect.left
	invoke PatBlt, hdc, rect.left, eax, ecx, dwGripSize, PATINVERT

	invoke SelectObject, hdc, hBrushOld
	invoke ReleaseDC, hWnd, hdc
	mov eax, yPos
	mov g_iOldResizeLine, eax
	ret
	align 4

DrawResizeLine endp

endif

NotifyParent proc hWnd:HWND, lParam:LPARAM

local rect:RECT
local sbn:SBNOTIFY

	movsx eax, word ptr lParam+2
	mov rect.top, eax
	mov rect.left, 0
	invoke ClientToScreen, hWnd, addr rect
	invoke ScreenToClient, [ebx].CSplittButton.hWndParent, addr rect
	invoke GetWindowLong, hWnd, GWL_ID
	mov sbn.hdr.idFrom, eax
	mov ecx, hWnd
	mov sbn.hdr.hwndFrom, ecx
	mov sbn.hdr.code, SBN_SETSIZE
	mov edx, rect.top
	mov sbn.iPos, edx
	lea ecx, sbn
	invoke SendMessage, [ebx].CSplittButton.hWndParent, WM_NOTIFY, eax, ecx
	ret
	align 4
NotifyParent endp

SplittBtnProc proc uses ebx hWnd:HWND, uMsg:UINT, wParam:WPARAM, lParam:LPARAM

local rect:RECT
local rect2:RECT

	mov eax,uMsg
	.if eax==WM_MOUSEMOVE

		invoke GetCapture
		.if (eax == hWnd)
			invoke GetWindowLong, hWnd, GWL_USERDATA
			mov ebx,eax
if ?RESIZEATONCE eq 0
			invoke DrawResizeLine, hWnd, g_iOldResizeLine
			movsx eax, word ptr lParam+2
			invoke DrawResizeLine, hWnd, eax
else
			invoke NotifyParent, hWnd, lParam
endif
		.endif

	.elseif eax==WM_LBUTTONDOWN

		invoke GetWindowLong, hWnd, GWL_USERDATA
		mov ebx,eax

		invoke SetCapture, hWnd

		invoke GetWindowRect, [ebx].CSplittButton.hWnd1, addr rect
		invoke GetWindowRect, [ebx].CSplittButton.hWnd2, addr rect2
		invoke UnionRect, addr rect, addr rect, addr rect2
		mov ecx, rect.bottom
		sub ecx, rect.top
		.if ((!CARRY?) && (ecx > MINSIZE))
			.if (ecx > MINSIZE * 2)
				add rect.top, MINSIZE
			.endif
			sub rect.bottom, MINSIZE
		.endif
		invoke ClipCursor, addr rect
if ?RESIZEATONCE eq 0
		mov g_iOldResizeLine, -1
		invoke DrawResizeLine, hWnd, 0
endif

	.elseif eax==WM_LBUTTONUP

		invoke GetCapture
		.if (eax == hWnd)
			invoke GetWindowLong, hWnd, GWL_USERDATA
			mov ebx,eax
			invoke ClipCursor, NULL
			invoke ReleaseCapture
if ?RESIZEATONCE eq 0
			invoke DrawResizeLine, hWnd, g_iOldResizeLine
			invoke NotifyParent, hWnd, lParam
endif
		.endif

	.elseif eax==WM_SETCURSOR

		invoke SetCursor, g_hHSCsr
		mov eax, 1
		jmp done

	.elseif eax==WM_DESTROY

		invoke GetWindowLong, hWnd, GWL_USERDATA
		invoke free, eax

	.endif
	invoke CallWindowProc, g_OldSplittBtnProc, hWnd, uMsg, wParam, lParam
done:
	ret
	align 4

SplittBtnProc endp

Create@CSplittButton proc public uses ebx hWnd:HWND, hWnd1:HWND, hWnd2:HWND

local	rect:RECT
local	rect2:RECT

	.if (!g_hHSCsr)
		invoke LoadCursor, g_hInstance, IDC_CURSOR1
		mov g_hHSCsr, eax
;;		invoke LoadCursor, g_hInstance, IDC_CURSOR2
;;		mov g_hVSCsr, eax
	.endif

	invoke malloc, sizeof CSplittButton
	.if (eax)
		mov ebx, eax
		invoke SetWindowLong, hWnd, GWL_USERDATA, ebx
		mov eax, hWnd1
		mov [ebx].CSplittButton.hWnd1, eax
		mov eax, hWnd2
		mov [ebx].CSplittButton.hWnd2, eax
		invoke GetParent, hWnd
		mov [ebx].CSplittButton.hWndParent, eax

		invoke SetWindowLong, hWnd, GWL_WNDPROC, addr SplittBtnProc
		mov g_OldSplittBtnProc, eax

		invoke GetWindowRect, hWnd1, addr rect
		invoke ScreenToClient, [ebx].CSplittButton.hWndParent, addr rect
		invoke ScreenToClient, [ebx].CSplittButton.hWndParent, addr rect.right

		invoke GetWindowRect, hWnd2, addr rect2
		invoke ScreenToClient, [ebx].CSplittButton.hWndParent, addr rect2
		mov ecx, rect.right
		sub ecx, rect.left
		mov eax, rect2.top
		sub eax, rect.bottom
		invoke SetWindowPos, hWnd, NULL, rect.left, rect.bottom, ecx, eax, SWP_NOZORDER or SWP_NOACTIVATE

		mov eax, ebx
	.endif
	ret
	align 4
Create@CSplittButton endp

	end
