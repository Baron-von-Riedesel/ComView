
;*** application "class", WinMain ***

	.386
	.model flat,stdcall
	option casemap :none   ; case sensitive
	option proc:private

	include COMView.inc
	include richedit.inc
	include wininet.inc

	include classes.inc
	include rsrc.inc
	include CEditDlg.inc
	include debugout.inc

	includelib wininet.lib


VERSIONSTRING	equ <"2.9.12">

;--- year, month, dayofweek, day, hour, minute

COMPILEDATE		equ <{2011,2,,25,14,00}>		;;GMT (CET - 1/2 hours) !!!
ifdef _DEBUG
HTTPSERVER		equ <"gra3">
else
HTTPSERVER		equ <"www.japheth.de">
endif

	.const

g_szCOMViewHttp	db "http://",HTTPSERVER,"/Download/COMView.zip",0

	.data

g_hInstance HINSTANCE NULL			;COMView module handle
_onexitbegin LPVOID NULL			;for atexit() crt function
g_hCurAccel	HACCEL NULL				;current accelerator
g_pMainDlg	pCDlg 0					;main dialog object
g_hWndDlg	HWND NULL
g_hMenu		HMENU NULL
g_hIconApp	HICON NULL
g_hWndOption	HWND NULL
g_pcdi		pCInterfaceList NULL
g_dwCnt		DWORD 0
g_argc		DWORD 0
g_argv		LPVOID NULL
g_himlLV	HANDLE NULL
g_heap		HANDLE NULL
g_hCsrWait	HCURSOR NULL
g_dwMyCBFormat dword 0
g_pViewObjectDlg pCViewObjectDlg NULL
g_pDataObject	LPDATAOBJECT NULL

if ?HTMLHELP eq 0
g_pszAboutText LPSTR 0
endif

	.const

g_szHint	db "Hint",0
g_szWarning	db "Warning",0
g_szNull	db 0

g_szAboutText	db 13,10,"COMView Version ",VERSIONSTRING
				db 13,10, "Copyright Japheth 2001-2010."
				db 13,10,"http://www.japheth.de"
				db 13,10
				db 13,10, "internal Hex Editor written by KetilO 2003"
				db 13,10
				db 13,10,"COMView is free."
				db 0

ifdef _DEBUG
DEBUGPREFIX LPSTR CStr("COMView:")
;?MALLOCSPY	equ 1
endif

	.code

if ?HTMLHELP eq 0

;*** edit stream callback for about dialog


editstreamcb proc uses esi dwCookie:DWORD, pbBuff:LPBYTE , cb:LONG , pcb:ptr LONG

local	i:DWORD

	invoke PostMessage, dwCookie, EM_SETSEL, -1, 0
	.if (!g_pszAboutText)
		invoke FindResource, g_hInstance,IDR_TEXT1,RT_RCDATA
		.if (eax != 0)
			invoke LoadResource, g_hInstance,eax
			.if (eax)
				invoke LockResource, eax
				mov g_pszAboutText,eax
			.endif
		.endif
	.endif
	mov esi,g_pszAboutText

	.if (esi)
		mov eax, g_dwCnt
		add esi,eax
		invoke lstrlen, esi
		.if (eax > cb)
			mov edx,pcb
			mov ecx,cb
			mov [edx],ecx
			invoke CopyMemory, pbBuff, esi, cb
			mov eax,cb
			add g_dwCnt,eax
			xor eax, eax
			ret
		.else
			mov edx,pcb
			mov ecx,cb
			mov [edx],eax
			invoke CopyMemory, pbBuff, esi, eax
			xor eax, eax
			ret
		.endif
	.endif
	mov eax,1
	ret
	align 4
editstreamcb endp


;*** help dialog proc


helpdialogproc proc public hWnd:HWND,message:dword,wParam:WPARAM,lParam:LPARAM

local	pStr:LPSTR
local	hWndRE:HWND
local	rect:RECT
local	estrm:EDITSTREAM

		mov eax,message
		.if (eax == WM_INITDIALOG)

			mov g_dwCnt,0
			invoke SendMessage, hWnd, WM_SETICON, ICON_SMALL, g_hIconApp
			invoke SendMessage, hWnd, WM_SETICON, ICON_BIG, g_hIconApp
			invoke GetDlgItem, hWnd, IDC_EDIT1
			mov hWndRE, eax

			mov estrm.dwCookie,eax
			mov estrm.dwError,0
			mov estrm.pfnCallback,offset editstreamcb

			invoke SendMessage, hWndRE, EM_STREAMIN, SF_RTF, addr estrm
			invoke CenterWindow,hWnd
			mov eax,0

		.elseif (eax == WM_SIZE)

			invoke GetDlgItem, hWnd, IDC_EDIT1
			mov hWndRE, eax
			invoke GetChildPos, hWndRE
			movzx ecx, word ptr lParam+0
			push eax
			movzx eax, ax
			sub ecx, eax
			sub ecx, eax
			movzx edx, word ptr lParam+2
			pop eax
			shr eax, 15
			sub edx, eax
			invoke SetWindowPos, hWndRE, NULL, 0, 0, ecx, edx, SWP_NOMOVE or SWP_NOZORDER or SWP_NOACTIVATE
			invoke GetClientRect, hWndRE, addr rect
			add rect.left, 10
			sub rect.right, 4
			invoke SendMessage, hWndRE, EM_SETRECT, 0, addr rect

		.elseif (eax == WM_CLOSE)

			invoke EndDialog,hWnd,0

		.elseif (eax == WM_COMMAND)
			movzx eax,word ptr wParam		;use only LOWORD(wParam)
			.if (eax == IDCANCEL)
				invoke PostMessage,hWnd,WM_CLOSE,0,0
			.else
				xor eax,eax
			.endif
		.else
			xor eax,eax
		.endif
		ret
		align 4

helpdialogproc endp

endif

if ?UPDATECHK

ifndef HINTERNET
HINTERNET typedef HANDLE
endif
CONTEXT_ID	equ 1

	.const
compiletime	SYSTEMTIME COMPILEDATE
	.data

protoInternetOpen				typedef proto :DWORD, :DWORD, :DWORD, :DWORD, :DWORD
protoInternetCloseHandle		typedef proto :DWORD
protoInternetOpenUrl			typedef proto :DWORD, :DWORD, :DWORD, :DWORD, :DWORD, :DWORD
protoHttpQueryInfo				typedef proto :DWORD, :DWORD, :DWORD, :DWORD, :DWORD
protoInternetTimeToSystemTime	typedef proto :DWORD, :DWORD, :DWORD

LPINTERNETOPEN	typedef ptr protoInternetOpen
LPINTERNETCLOSEHANDLE typedef ptr protoInternetCloseHandle
LPINTERNETOPENURL typedef ptr protoInternetOpenUrl
LPHTTPQUERYINFO typedef ptr protoHttpQueryInfo
LPINTERNETTIMETOSYSTEMTIME typedef ptr protoInternetTimeToSystemTime

g_pfnInternetOpen			LPINTERNETOPEN			NULL
g_pfnInternetCloseHandle	LPINTERNETCLOSEHANDLE	NULL
g_pfnInternetOpenUrl		LPINTERNETOPENURL		NULL
g_pfnHttpQueryInfo			LPHTTPQUERYINFO			NULL
g_pfnInternetTimeToSystemTime	LPINTERNETTIMETOSYSTEMTIME	NULL

	.code

CheckUpdate proc public hWnd:HWND

local	hInternet:HINTERNET
local	hUrl:HINTERNET
local	ib:INTERNET_BUFFERS
local	estrm:EDITSTREAM
local	dwIndex:DWORD
local	dwBufLength:DWORD
local	systemtime:SYSTEMTIME
local	rc:DWORD
local	hLib:HINSTANCE
local	szHeader[256]:byte

	mov hInternet, NULL

	.if (!g_pfnInternetOpen)
		invoke LoadLibrary, CStr("WININET")
		.if (eax > 32)
			mov hLib, eax
			invoke GetProcAddress, hLib, CStr("InternetOpenA")
			mov g_pfnInternetOpen, eax
			invoke GetProcAddress, hLib, CStr("InternetOpenUrlA")
			mov g_pfnInternetOpenUrl, eax
			invoke GetProcAddress, hLib, CStr("HttpQueryInfoA")
			mov g_pfnHttpQueryInfo, eax
			invoke GetProcAddress, hLib, CStr("InternetTimeToSystemTimeA")
			mov g_pfnInternetTimeToSystemTime, eax
			invoke GetProcAddress, hLib, CStr("InternetCloseHandle")
			mov g_pfnInternetCloseHandle, eax
		.endif
		.if ((!g_pfnInternetOpen) || (!g_pfnInternetOpenUrl) || (!g_pfnHttpQueryInfo))
			jmp error
		.endif
	.endif

	invoke g_pfnInternetOpen, CStr("COMView"),INTERNET_OPEN_TYPE_DIRECT,\
			NULL, NULL, 0
	.if (eax == NULL)
		jmp error
	.endif
	mov hInternet,eax

	invoke g_pfnInternetOpenUrl, hInternet,\
		addr g_szCOMViewHttp, NULL, NULL,\
		INTERNET_FLAG_RELOAD, CONTEXT_ID
	.if (eax == NULL)
		jmp error
	.endif
	mov hUrl, eax

	invoke ZeroMemory, addr szHeader, sizeof szHeader

	mov dwIndex,0
	mov dwBufLength, sizeof szHeader
	invoke g_pfnHttpQueryInfo, hUrl, HTTP_QUERY_LAST_MODIFIED,
			addr szHeader, addr dwBufLength, addr dwIndex

	mov rc, eax
	.if (eax)
		invoke g_pfnInternetTimeToSystemTime, addr szHeader, addr systemtime, 0
		mov systemtime.wDayOfWeek, 0
		mov ecx, sizeof SYSTEMTIME/sizeof WORD
		pushad
		lea edi, systemtime
		lea esi, compiletime
		repz cmpsw
		popad
		sbb eax, eax	;-1 if newer version available
		neg eax
		mov rc, eax
	.endif
	invoke g_pfnInternetCloseHandle, hUrl
	invoke g_pfnInternetCloseHandle, hInternet
	return rc
error:
	.if (hInternet)
		invoke g_pfnInternetCloseHandle, hInternet
	.endif
	invoke wsprintf, addr szHeader, CStr("Cannot connect to %s"), addr g_szCOMViewHttp
	invoke MessageBox, hWnd, addr szHeader, 0, MB_OK
	mov eax, -1
	ret

CheckUpdate endp


endif

aboutdialogproc proc public hWnd:HWND,message:dword,wParam:WPARAM,lParam:LPARAM

local	pStr:LPSTR
local	hWndRE:HWND
local	estrm:EDITSTREAM

		mov eax,message
		.if (eax == WM_INITDIALOG)

			invoke SetDlgItemText, hWnd, IDC_EDIT1, addr g_szAboutText
			
			invoke CenterWindow,hWnd
			mov eax,1
		.elseif (eax == WM_CLOSE)
			invoke EndDialog,hWnd,0
		.elseif (eax == WM_COMMAND)
			movzx eax,word ptr wParam		;use only LOWORD(wParam)
			.if (eax == IDCANCEL)
				invoke PostMessage,hWnd,WM_CLOSE,0,0
			.else
				xor eax,eax
			.endif
		.else
			xor eax,eax
		.endif
		ret
		align 4

aboutdialogproc endp

LoadSave proc bMode:BOOL
	
local szModule[MAX_PATH]:byte

		invoke	GetModuleFileName, NULL, addr szModule,  sizeof szModule
		lea ecx, szModule
		mov dword ptr [ecx+eax-3],"ini"
		.if (bMode)
			invoke	Save@COptions, addr szModule
		.else
			invoke	Load@COptions, addr szModule
		.endif
		ret
LoadSave endp

;*** constructor application object


Create@CApp proc uses esi edi hInstance:HINSTANCE

local iccx:INITCOMMONCONTROLSEX

		mov eax,hInstance
		mov g_hInstance,eax

		invoke GetProcessHeap
		mov g_heap,eax

		mov iccx.dwSize,sizeof INITCOMMONCONTROLSEX
		mov iccx.dwICC,ICC_LISTVIEW_CLASSES or ICC_TREEVIEW_CLASSES or \
			ICC_BAR_CLASSES or ICC_PROGRESS_CLASS or ICC_USEREX_CLASSES
		invoke InitCommonControlsEx,addr iccx

		invoke LoadSave, 0

;		invoke CoInitialize,NULL
		invoke OleInitialize,NULL

ifdef ?MALLOCSPY
		invoke Create@CMallocSpy
		.if (eax)
			invoke CoRegisterMallocSpy, eax
		.endif
endif

		mov g_pcdi,NULL					;init Interface list

		invoke LoadIcon,g_hInstance,IDI_ICON1
		mov g_hIconApp,eax

		invoke LoadMenu,g_hInstance,IDR_MENU2
		mov g_hMenu,eax

		invoke LoadCursor,NULL,IDC_WAIT
		mov g_hCsrWait,eax

		invoke ImageList_LoadImage, g_hInstance, IDB_BITMAP2, 0, 0, CLR_DEFAULT, \
				IMAGE_BITMAP, LR_DEFAULTCOLOR
		mov g_himlLV, eax
if 0
		invoke ImageList_Create, 16, 16, ILC_COLORDDB or ILC_MASK, 1, 0
		mov esi,eax
		mov g_himlLV,esi
		invoke LoadImage, g_hInstance, IDI_CROSS, IMAGE_ICON, 16, 16, LR_DEFAULTCOLOR
		invoke ImageList_AddIcon( esi, eax )
		invoke LoadImage, g_hInstance, IDI_CROSS2, IMAGE_ICON, 16, 16, LR_DEFAULTCOLOR
		invoke ImageList_AddIcon( esi, eax )
endif

		invoke SetArguments

		invoke RegisterClipboardFormat, CStr("COMView")
		mov g_dwMyCBFormat, eax

		ret
		align 4

Create@CApp endp


;*** destructor application object

Destroy@CApp proc


		.if (g_pStorage)
			invoke vf(g_pStorage, IStorage, Release)
			mov g_pStorage, NULL
		.endif
		.if (g_pDataObject)
			invoke OleIsCurrentClipboard, g_pDataObject
			.if (eax == S_OK)
				invoke OleFlushClipboard
			.endif
		.endif

		invoke LoadSave, 1

		.if (g_pcdi != NULL)
			invoke Destroy@CInterfaceList,g_pcdi
			mov g_pcdi,NULL
		.endif
		.if (g_hMenu)
			invoke DestroyMenu, g_hMenu
			mov g_hMenu, NULL
		.endif
		.if (g_hIconApp)
			invoke DestroyIcon, g_hIconApp
			mov g_hIconApp, NULL
		.endif

ifdef ?MALLOCSPY
		invoke CoRevokeMallocSpy
endif
;		invoke CoUninitialize
		invoke OleUninitialize

		invoke free, g_argv

		ret
		align 4

Destroy@CApp endp


;*** WinMain: create CMainDlg, enter message loop, destroy CMainDlg


WinMain proc hInstance:HINSTANCE,hPrevInstance:HINSTANCE,lpszCmdLine:LPSTR,iCmdShow:dword

local msg:MSG

		invoke Create@CApp, hInstance	;create global objects
		invoke Create@CMainDlg			;create main dialog

		.while (1)							;main message loop
			invoke GetMessage, addr msg, NULL, 0, 0
			.break .if (eax == 0)
			.if (g_hCurAccel)
				invoke TranslateAccelerator, g_hWndDlg, g_hCurAccel, addr msg
				.continue .if (eax)
			.elseif (g_pViewObjectDlg)
				invoke TranslateAccelerator@CViewObjectDlg, g_pViewObjectDlg, addr msg
				.continue .if (eax)
			.endif
if 0
			DebugOut "%X received", msg.message
endif
			invoke IsDialogMessage, g_hWndDlg, addr msg
			.continue .if (eax)
			invoke TranslateMessage, addr msg
			invoke DispatchMessage, addr msg
		.endw

		invoke Destroy@CMainDlg, g_pMainDlg		;destroy main dialog
		invoke Destroy@CApp
		xor eax,eax
		ret
		align 4

WinMain endp


;*** simple initialization to get hInstance ***


mainCRTStartup proc public

		invoke GetModuleHandle,0
		invoke WinMain,eax,0,0,0
		push eax
		invoke _cexit
		pop eax
		invoke ExitProcess,eax

mainCRTStartup endp

		end mainCRTStartup
