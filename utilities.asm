
;*** application wide utility functions (static, not class related) ***

	.386
	.model flat,stdcall
	option casemap :none   ; case sensitive
	option proc:private

	include COMView.inc
	.nolist
	.nocref
	include shellapi.inc
	include ExDispId.inc
	include MsHtmDid.inc
	include IDispIds.inc
	.cref
	.list
	include statusbar.inc
	include classes.inc
	include rsrc.inc
	include CEditDlg.inc
	include debugout.inc

_ALLOCDEBUG	equ 0
?SHOWICON	equ 1
?SHOWAMBIENT	equ 0		;show "Ambient" tab
?USEBMP		equ 0			;use header bitmaps instead of imagelist

LVCOMP	struct
hWndLV	HWND ?
dwFlags	dd ?
iSortCol db ?
iSortDir db ?
LVCOMP	ends


	.const

pVARTYPE label BYTE
	dw	VT_EMPTY,	VT_NULL,		VT_I2,		VT_I4,\
		VT_R4,		VT_R8,			VT_CY,		VT_DATE,\
		VT_BSTR,	VT_DISPATCH,	VT_ERROR,	VT_BOOL,\
		VT_VARIANT,	VT_UNKNOWN,		VT_DECIMAL,	VT_I1
	dw	VT_UI1,		VT_UI2,			VT_UI4,		VT_I8,\
		VT_UI8,		VT_INT,			VT_UINT,	VT_VOID,\
		VT_HRESULT,	VT_PTR,			VT_SAFEARRAY,VT_CARRAY,\
		VT_USERDEFINED,VT_LPSTR,	VT_LPWSTR,	VT_FILETIME
	dw	VT_BLOB,	VT_STREAM,		VT_STORAGE,	VT_STREAMED_OBJECT,\
		VT_STORED_OBJECT,VT_BLOB_OBJECT,VT_CF,	VT_CLSID,\
		VT_VECTOR,	VT_ARRAY,		VT_BYREF,	VT_RESERVED,\
		VT_ILLEGAL,	VT_ILLEGALMASKED,VT_TYPEMASK
NUMVT equ ($ - pVARTYPE) / sizeof word

	align 4

pVARTYPEStr label ptr
	dd	CStr("Empty"),	CStr("Null"),	CStr("I2"),		CStr("I4")
	dd	CStr("R4"),		CStr("R8"),		CStr("CY"),		CStr("Date")
	dd	CStr("Bstr"),	CStr("Dispatch"),CStr("Error"),	CStr("Bool")
	dd	CStr("Variant"),CStr("Unknown"),CStr("Decimal"),CStr("I1")
	dd	CStr("UI1"),	CStr("UI2"),	CStr("UI4"),	CStr("I8")
	dd	CStr("UI8"),	CStr("Int"),	CStr("UInt"),	CStr("Void")
	dd	CStr("HResult"),CStr("Ptr"),	CStr("SafeArray"),CStr("CArray")
	dd	CStr("UserDefined"),CStr("LPStr"),CStr("LPWStr"),CStr("FileTime")
	dd	CStr("Blob"),	CStr("Stream"),	CStr("Storage"),CStr("Streamed Object")
	dd	CStr("Stored Object"),CStr("Blob Object"),CStr("CF"), g_szCLSID
	dd	CStr("Vector"),	CStr("Array"),	CStr("ByRef"),	CStr("Reserved")
	dd	CStr("Illegal"),CStr("IllegalMasked"),CStr("TypeMask")
	dd	CStr("?")

;--- ASM types
;--- DATE is REAL8
;--- CY is REAL8

pVARTYPEStrAsm label ptr
	dd	CStr("Empty"),	CStr("Null"),	CStr("SWORD"),	CStr("SDWORD")
	dd	CStr("REAL4"),	CStr("REAL8"),	CStr("CY"),		CStr("DATE")
	dd	CStr("BSTR"),	CStr("LPDISPATCH"),CStr("SCODE"),CStr("BOOL")
	dd	CStr("VARIANT"),CStr("LPUNKNOWN"),CStr("Decimal"),CStr("SBYTE")
	dd	CStr("BYTE"),	CStr("WORD"),	CStr("DWORD"),	CStr("SQWORD")
	dd	CStr("QWORD"),	CStr("SDWORD"),	CStr("DWORD"),	CStr("")
	dd	CStr("HRESULT"),CStr("ptr"),	CStr("SAFEARRAY"),CStr("CArray")
	dd	CStr("UserDefined"),CStr("LPSTR"),CStr("LPWSTR"),CStr("FileTime")
	dd	CStr("Blob"),	CStr("Stream"),	CStr("Storage"),CStr("Streamed Object")
	dd	CStr("Stored Object"),CStr("Blob Object"),CStr("CF"),g_szCLSID
	dd	CStr("Vector"),	CStr("Array"),	CStr("ByRef"),	CStr("Reserved")
	dd	CStr("Illegal"),CStr("IllegalMasked"),CStr("TypeMask")
	dd	CStr("?")

dwTypeKind label dword
	dd TKIND_ENUM
	dd TKIND_RECORD
	dd TKIND_MODULE
	dd TKIND_INTERFACE
	dd TKIND_DISPATCH
	dd TKIND_COCLASS
	dd TKIND_ALIAS
	dd TKIND_UNION
	dd TKIND_MAX
NUMTYPEKIND equ ($ - dwTypeKind) / sizeof dword

pszTypeKind label dword
	dd CStr("ENUM")
	dd CStr("RECORD")
	dd CStr("MODULE")
	dd CStr("INTERFACE")
	dd CStr("DISPATCH")
	dd CStr("COCLASS")
	dd CStr("ALIAS")
	dd CStr("UNION")
	dd CStr("MAX")
	dd CStr("???")

;--- table of MASM reserved names to avoid as function/var names
;--- yes, there surely are missing some names. Feel free to add them here

ReservedNames label dword
	dd CStr("add")
	dd CStr("align")
	dd CStr("and")
	dd CStr("comment")
	dd CStr("cx")
	dd CStr("db")
	dd CStr("dw")
	dd CStr("dd")
	dd CStr("dq")
	dd CStr("echo")
	dd CStr("end")
	dd CStr("enter")
	dd CStr("even")
	dd CStr("goto")
	dd CStr("group")
	dd CStr("invoke")
	dd CStr("label")
	dd CStr("length")
	dd CStr("lock")
	dd CStr("name")
	dd CStr("offset")
	dd CStr("or")
	dd CStr("record")
	dd CStr("repeat")
	dd CStr("rept")
	dd CStr("size")
	dd CStr("type")
	dd CStr("union")
	dd CStr("wait")
	dd CStr("width")
SIZERESERVEDNAMES equ ($ - ReservedNames) / sizeof DWORD

	.data

g_pszMenuHelp LPSTR NULL
g_szLastDir db MAX_PATH dup (0)

	.const

g_szRootCLSID		db "CLSID",0
g_szRootTypeLib		db "TypeLib",0
g_szRootInterface	db "Interface",0
g_szRootAppID		db "AppID",0
g_szRootCompCat		db "Component Categories",0

	.code

strchr proc c public uses edi pStr:ptr SBYTE, dwByte:dword

;;		and eax,3							???
		mov edi,pStr
		invoke lstrlen,edi
		mov ecx,eax
		jecxz @F
		mov al,byte ptr dwByte
		repnz scasb
		jnz @F
		lea eax,[edi-1]
		ret
@@:
		xor eax,eax
		ret
		align 4

strchr endp

;*** malloc: memory is ZEROINIT!

malloc proc public dwBytes:DWORD
		invoke HeapAlloc, g_heap, HEAP_ZERO_MEMORY or HEAP_NO_SERIALIZE, dwBytes
if _ALLOCDEBUG
		.data
g_cntMalloc	DWORD 0
		.code
		pushad
		sub esp, 128
		inc g_cntMalloc
		mov edx, esp
		invoke wsprintf, edx, CStr("%3u malloc(%u) returned %X, called from %X",13,10), g_cntMalloc, dwBytes, eax, dword ptr [ebp+4]
		invoke OutputDebugString, esp
		add esp, 128
		popad
endif
		ret
		align 4
malloc endp

free proc public pMem:ptr byte
		.if (pMem)
if _ALLOCDEBUG
			dec g_cntMalloc
			sub esp, 128
			invoke HeapSize, g_heap, HEAP_NO_SERIALIZE, pMem
			.if (eax == -1)
				mov edx, esp
				invoke wsprintf, edx, CStr("Heap Error with %X at %X"), pMem, dword ptr [ebp+4]
				mov edx, esp
				invoke MessageBox, 0, edx, 0, MB_OK
			.else
				invoke FillMemory, pMem, eax, 0AAh
			.endif
			mov edx, esp
			invoke wsprintf, edx, CStr("%3u free(%X), called from %X",13,10), g_cntMalloc, pMem, dword ptr [ebp+4]
			invoke OutputDebugString, esp
			add esp, 128
endif
			invoke HeapFree, g_heap, HEAP_NO_SERIALIZE, pMem
		.endif
		ret
		align 4
free endp

;*** convert string to number

String2Number proc public uses esi ebx pStr:LPSTR,pDW:ptr dword,radix:dword

		mov ecx,radix
		xor edx,edx
		mov esi,pStr
		mov bl,0
next:
		lodsb
		and al,al
		jz exit
		cmp al,' '
		jz exit
if 1
		cmp al,','
		jz exit
endif
		sub al,'0'
		jc errexit
		cmp al,9
		jbe @F
		cmp cl,16
		jc errexit
		or al,20h
		sub al,27h
		jc errexit
		cmp al,10h
		cmc
		jc errexit
@@:
		movzx eax,al
		push eax
		mov eax,edx
		mul ecx
		pop edx
		add edx,eax
		inc bl
		jmp next
exit:
		cmp bl,1	;C if 0 (error)
		jc errexit
		mov eax,pDW
		mov [eax],edx
		lea eax,[esi-1]
		ret
errexit:
		xor eax,eax
		ret
		align 4
String2Number endp

;*** convert string to number, check radix first

String2DWord proc public uses esi ebx pStr:LPSTR, pDW:ptr dword

		mov bl,00
		mov esi,pStr
		mov ax,[esi]
		cmp al,'-'
		jnz @F
		mov bl, 1
		inc esi
		mov ax,[esi]
@@:
		mov ecx,10
		or ah,20h
		cmp ax,'x0'
		jnz @F
		mov cl,16
		inc esi
		inc esi
@@:
		invoke String2Number,esi,pDW,ecx
		.if (bl && eax)
			mov ecx, pDW
			neg dword ptr [ecx]
		.endif
		ret
		align 4
String2DWord endp


;*** convert string to 2 numbers, separated by '.'


String22DWords proc public uses esi edi pStr:LPSTR,pDW1:ptr dword,pDW2:ptr dword

local	szStr[64]:byte

		mov esi,pStr
		lea edi,szStr
		.while (1)
			lodsb
			stosb
			.break .if (al == 0)
			.if (al == '.')
;------------------------- interpret the minor number as decimal
;------------------------- (no problem since LoadRegTypeLib will always
;------------------------- load the version with the highest minor version#)
				invoke String2Number,esi,pDW2,10
				mov byte ptr [edi-1],0
				.break
			.endif
		.endw
;------------------------- that may be a problem, but currently we have
;------------------------- only a hex major version > 9
		invoke String2Number,addr szStr,pDW1,16
		ret
		align 4
String22DWords endp

String2DWords proc public uses esi edi pStr:LPSTR, iNum:DWORD, pArray:ptr DWORD

		mov esi, 0
		mov edi, pStr
		.while (esi < iNum)
			mov ecx, esi
			shl ecx, 2
			add ecx, pArray
			invoke String2DWord, edi, ecx
			.break .if (!eax)
			mov edi, eax
			.if (byte ptr [edi] == ',')
				inc edi
			.endif
			inc esi
		.endw
		ret
		align 4
String2DWords endp


;*** convert "{xxxxxxxx-xxxx-xxxx-...} to GUID ***


GUIDFromLPSTR proc public pszGUID:LPSTR, pguid:ptr GUID

local	wszGUID[40]:WORD

		invoke MultiByteToWideChar,CP_ACP,MB_PRECOMPOSED,
			pszGUID,40,addr wszGUID, 40 
		invoke CLSIDFromString,addr wszGUID,pguid
		ret
		align 4

GUIDFromLPSTR endp

SysStringFromLPSTR proc public pszString:LPSTR, dwSize:DWORD
	mov eax, dwSize
	.if (!eax)
		invoke lstrlen, pszString
		inc eax
		mov dwSize, eax
	.endif
	invoke SysAllocStringLen, NULL, eax
	push eax
	invoke MultiByteToWideChar, CP_ACP, MB_PRECOMPOSED,
			pszString, -1, eax, dwSize
	pop eax
	ret
	align 4

SysStringFromLPSTR endp

;*** center window on parent window ***

CenterWindow proc public uses edi hWnd:HWND

local rect:RECT
local rect2:RECT
local dwCX:dword
local dwCY:dword

		invoke GetParent, hWnd
		mov edi, eax
		.if (eax)
			invoke IsIconic, eax
			.if (eax)
				xor edi,edi
			.endif
		.endif
		.if (edi == 0)
			mov rect2.top,0
			mov rect2.left,0
			invoke GetSystemMetrics,SM_CXFULLSCREEN
			mov rect2.right,eax
			invoke GetSystemMetrics,SM_CYFULLSCREEN
			mov rect2.bottom,eax
		.else
			invoke GetWindowRect, edi, addr rect2 
		.endif
		mov eax,rect2.right
		sub eax,rect2.left
		mov dwCX,eax
		mov eax,rect2.bottom
		sub eax,rect2.top
		mov dwCY,eax

		invoke GetWindowRect,hWnd,addr rect 

		mov edx,rect.right
		sub edx,rect.left
		shr edx,1			;edx = (rect.right - rect.left)/2

		mov eax,dwCX
		shr eax,1			;eax = dwCX / 2
		sub eax,edx			;Parent(CX)/2 - Child(CX)/2
		add eax,rect2.left
		test eax,eax
		.if SIGN? 
			xor eax,eax
		.endif

		mov edx,rect.bottom
		sub edx,rect.top
		shr edx,1			;edx = (rect.bottom - rect.top)/2

		mov ecx,dwCY
		shr ecx,1			;ecx = dwCY / 2
		sub ecx,edx			;ecx = dwCY / 2 - (rect.bottom - rect.top) / 2
		add ecx,rect2.top
		test ecx,ecx
		.if SIGN? 
			xor ecx,ecx
		.endif

		invoke SetWindowPos,hWnd,0,eax,ecx,0,0,SWP_NOSIZE or SWP_NOZORDER
		ret
		align 4

CenterWindow endp


SetLVColumns proc public uses ebx esi hWndLV:HWND, numCols:dword,pColHdr:ptr CColHdr

local	dwCXLeft:DWORD
local	iWidthTotal:DWORD
local	rect:RECT
local	lvc:LVCOLUMN

		invoke GetClientRect, hWndLV,addr rect
if 1
		invoke GetWindowLong, hWndLV, GWL_STYLE
		.if !(eax & WS_VSCROLL)
			invoke GetSystemMetrics, SM_CXVSCROLL
			inc eax
			sub rect.right,eax
		.endif
else
		invoke GetSystemMetrics, SM_CXVSCROLL
		inc eax
		sub rect.right,eax
endif
		mov eax, rect.right
		dec eax
		mov dwCXLeft, eax

		xor ebx, ebx
		mov iWidthTotal,ebx
		mov esi,pColHdr
		.while (ebx < numCols)
			movzx eax, [esi].CColHdr.wWidth
			.if ([esi].CColHdr.wFlags & FCOLHDR_ABSOLUTE)
				movzx eax, ax
				invoke MulDiv, eax, 100, rect.right
			.endif
			movzx eax,ax
			add iWidthTotal,eax
			inc ebx
			add esi,sizeof CColHdr
		.endw


		mov lvc.mask_,LVCF_TEXT or LVCF_WIDTH or LVCF_FMT

		mov ebx,0
		mov esi,pColHdr
		.while (ebx < numCols)
			movzx eax,[ebx * sizeof CColHdr + esi].CColHdr.wWidth
			invoke MulDiv,rect.right,eax, iWidthTotal
			.if (eax > dwCXLeft)
				mov eax, dwCXLeft
			.endif
			sub dwCXLeft, eax
			mov lvc.cx_,eax
			mov eax,[ebx * sizeof CColHdr + esi].CColHdr.pColText
			mov lvc.pszText,eax
			.if ([ebx * sizeof CColHdr + esi].CColHdr.wFlags & FCOLHDR_RDXMASK)
				mov lvc.fmt, LVCFMT_RIGHT
			.else
				mov lvc.fmt, LVCFMT_LEFT
			.endif
			invoke ListView_InsertColumn( hWndLV, ebx, addr lvc)
			inc ebx
		.endw
		ret
		align 4

SetLVColumns endp


;*** get child position in main client area ***


GetChildPos proc public hWnd:HWND

local	rect:RECT
local	pt:POINT

		invoke GetWindowRect,hWnd,addr rect
		mov eax,rect.left
		mov pt.x,eax
		mov eax,rect.top
		mov pt.y,eax
		invoke GetParent,hWnd
		mov edx,eax
		invoke ScreenToClient,edx,addr pt
		mov eax,pt.y
		shl eax,16
		mov ax,word ptr pt.x
		ret
		align 4
GetChildPos endp


;*** read all GUID keys from entry (CLSID,TypeLib or Interface)
;*** save them in allocated memory block


ReadAllRegSubKeys proc public uses ebx esi edi hWnd:HWND, pszKey:LPSTR, pNumEntries:ptr dword

local	iid:IID
local	hKey:HANDLE
local	dwSize:dword
local	szIID[260]:byte
local	szStr[260]:byte
local	wszIID[40]:word
local	rc:ptr byte

		invoke RegOpenKeyEx,HKEY_CLASSES_ROOT,pszKey,0,KEY_READ,addr hKey
		.if (eax != ERROR_SUCCESS)
			invoke wsprintf,addr szStr,CStr("Key %s does not exist"),pszKey
			invoke MessageBox,hWnd,addr szStr,0,MB_OK
			xor eax,eax
			ret
		.endif

;--------------------------------------- get number of entries

		invoke RegQueryInfoKey,hKey,0,0,0,pNumEntries,0,0,0,0,0,0,0
		mov eax,pNumEntries
		mov eax,[eax]
		.if (eax == 0)
			invoke wsprintf,addr szStr,CStr("No subkeys in %s defined"),pszKey
			invoke MessageBox,hWnd,addr szStr,0,MB_OK
			invoke RegCloseKey,hKey
			xor eax,eax
			ret
		.endif

;--------------------------------------- alloc space for GUID list

		mov ecx,sizeof IID
		mul ecx
		invoke malloc, eax

		.if (eax == 0)
			invoke MessageBox,hWnd,CStr("Memory failure"),0,MB_OK
			invoke RegCloseKey,hKey
			xor eax,eax
			ret
		.endif
		mov rc,eax
		mov esi,eax

		mov eax,pNumEntries
		mov edi,[eax]
		xor ebx,ebx
		.while (ebx < edi)
			mov dwSize,sizeof szIID
			invoke RegEnumKeyEx,hKey,ebx,addr szIID,addr dwSize,0,NULL,0,NULL
			.break .if (eax != ERROR_SUCCESS)
			invoke MultiByteToWideChar,CP_ACP,MB_PRECOMPOSED,\
					addr szIID, -1, addr wszIID, 40 
			invoke IIDFromString,addr wszIID, esi
			add esi, sizeof IID
			inc ebx
		.endw
		invoke RegCloseKey,hKey
		mov eax,rc
		ret
		align 4
ReadAllRegSubKeys endp


;*** get Standard value from entry ***


GetTextFromCLSID proc public pGUID:ptr GUID, pStr:LPSTR, dwSize:dword

local	szStr[128]:byte
local	wszStr[40]:word
local	szGUID[40]:byte
local   hKey:HANDLE
local   dwType:dword

		mov ecx, pStr
		mov byte ptr [ecx],0
		invoke StringFromGUID2, pGUID,addr wszStr,40
		invoke WideCharToMultiByte,CP_ACP,0,addr wszStr,40,addr szGUID, sizeof szGUID,0,0 
		invoke wsprintf, addr szStr, CStr("%s\%s"), addr g_szRootCLSID, addr szGUID
		invoke RegOpenKeyEx,HKEY_CLASSES_ROOT,addr szStr,0,KEY_READ,addr hKey
		.if (eax == S_OK)
			invoke RegQueryValueEx,hKey,addr g_szNull,NULL,addr dwType,pStr,addr dwSize
			invoke RegCloseKey,hKey
		.endif
		ret
		align 4

GetTextFromCLSID endp


;*** a simple dialogproc with 1 edit field (IDC_EDIT1)


inputdlgproc proc public hWnd:HWND, uMsg:DWORD, wParam:WPARAM, lParam:LPARAM

		mov eax,uMsg
		.if (eax == WM_INITDIALOG)
			invoke SetDlgItemText, hWnd, IDC_EDIT1, lParam
			invoke SetWindowLong, hWnd, DWL_USER, lParam
			mov eax,1
		.elseif (eax == WM_CLOSE)
			invoke EndDialog, hWnd, 0
		.elseif (eax == WM_COMMAND)
			movzx eax,word ptr wParam
			.if (eax == IDOK)
				invoke GetWindowLong, hWnd, DWL_USER
				invoke GetDlgItemText, hWnd, IDC_EDIT1, eax, MAXINPUTTEXT
				invoke EndDialog, hWnd, 1
			.elseif (eax == IDCANCEL)
				invoke EndDialog, hWnd, 0
			.endif
		.else
			xor eax,eax
		.endif
		ret
		align 4

inputdlgproc endp


;*** delete registry key with all related subkeys


DeleteKeyWithSubKeys proc public uses ebx hKey:HANDLE,pszKey:LPSTR

local	szKey[260]:byte
local	hSubKey:HANDLE
local	filetime:FILETIME
local	dwSize:dword

		invoke RegOpenKeyEx,hKey,pszKey,NULL,KEY_READ or KEY_WRITE,addr hSubKey
		.if (eax == ERROR_SUCCESS)
			mov ebx,0
			.while (1)
				mov dwSize,sizeof szKey
				invoke RegEnumKeyEx,hSubKey,ebx,addr szKey,addr dwSize,NULL,NULL,NULL,addr filetime
				.break .if (eax != ERROR_SUCCESS)
				invoke DeleteKeyWithSubKeys,hSubKey,addr szKey
;				inc ebx
			.endw
			invoke RegCloseKey,hSubKey
		.endif
		invoke RegDeleteKey,hKey,pszKey		;and delete subkey
		ret
		align 4

DeleteKeyWithSubKeys endp


;*** reset all selected items in a listview


ResetListViewSelection proc public uses ebx hWndLV:HWND

		mov ebx,-1
if 0
		.while (1)
			ListView_GetNextItem hWndLV,ebx,LVNI_ALL
			.break .if (eax == -1)
			mov ebx,eax
			ListView_SetItemState hWndLV,ebx,0,LVIS_SELECTED or LVIS_FOCUSED
		.endw
else
		.while (1)
			invoke ListView_GetNextItem( hWndLV,ebx,LVNI_SELECTED)
			.break .if (eax == -1)
			mov ebx,eax
			ListView_SetItemState hWndLV,ebx,0,LVIS_SELECTED
		.endw
endif
		ret
		align 4

ResetListViewSelection endp


;*** display a messagebox with a FormatMessage message


OutputMessage proc public hWnd:HWND, dwErr:dword, pszCaption:LPSTR,pszPrefix:LPSTR

local	szFormatMsg[256]:byte
local	szStr[256+128]:byte
local	szErrCode[16]:byte
local	dwLen:dword

		.if (!pszPrefix)
			mov pszPrefix, offset g_szNull
		.endif
		invoke FormatMessage,FORMAT_MESSAGE_FROM_SYSTEM,\
				NULL,dwErr,NULL,addr szFormatMsg,sizeof szFormatMsg,0
		.if (!eax)
			invoke lstrcpy, addr szFormatMsg, CStr("<NO SYSTEM MESSAGE DEFINED>")
		.endif
		invoke wsprintf, addr szStr, CStr("%s%s (%08X)"),pszPrefix, addr szFormatMsg, dwErr
		invoke MessageBox, hWnd, addr szStr, pszCaption, MB_OK
		ret
		align 4

OutputMessage endp


;*** set lParam for listview items (preparation for sort) ***


LVSort proc public uses esi ebx hWndLV:HWND,iSortCol:dword,iSortDir:dword,dwFlags:dword

local	dwItems:DWORD
local	hCsrOld:HCURSOR
local	pSortTab:LPVOID
local	lvc:LVCOMP
local	lvi:LVITEM

		mov eax,hWndLV
		mov ecx,iSortCol
		mov edx,iSortDir
		mov lvc.hWndLV,eax
		mov lvc.iSortCol,cl
		mov lvc.iSortDir,dl
		mov eax,dwFlags
		mov lvc.dwFlags,eax

		invoke SetCursor,g_hCsrWait
		mov hCsrOld, eax

		invoke ListView_GetItemCount( lvc.hWndLV)
		.if (!eax)
			jmp exit
		.endif
		mov dwItems, eax

;------------------------------ alloc array to save lParams

		shl eax, 2
		invoke malloc, eax
		mov pSortTab, eax
		.if (!eax)
			jmp exit
		.endif

;------------------------------ save lParams in array, set lParam to index

		mov esi, dwItems
		@mov lvi.iItem, 0
		@mov lvi.iSubItem, 0
		mov lvi.mask_,LVIF_PARAM
		.while (esi)
			invoke ListView_GetItem( lvc.hWndLV, addr lvi)
			mov eax, lvi.lParam
			mov ecx, pSortTab
			mov edx, lvi.iItem
			mov [edx*4+ecx], eax
			mov lvi.lParam, edx
			invoke ListView_SetItem( lvc.hWndLV, addr lvi)
			inc lvi.iItem
			dec	esi
		.endw

;------------------------------ sort listview

		invoke ListView_SortItems( lvc.hWndLV, comparelvproc, addr lvc)

;------------------------------ restore lParams

		mov esi, dwItems
		@mov lvi.iItem, 0 
		@mov lvi.iSubItem, 0
		mov lvi.mask_, LVIF_PARAM
		.while (esi)
			invoke ListView_GetItem( lvc.hWndLV, addr lvi)
			mov eax, lvi.lParam
			mov ecx, pSortTab
			mov eax, [eax*4+ecx]
			mov lvi.lParam, eax
			invoke ListView_SetItem( lvc.hWndLV, addr lvi)
			inc lvi.iItem
			dec esi
		.endw

		invoke free, pSortTab
exit:
if ?HDRBMPS
		invoke ResetHeaderBitmap, hWndLV
		invoke SetHeaderBitmap, hWndLV, iSortCol, iSortDir
endif
		invoke SetCursor, hCsrOld
		ret
		align 4

LVSort endp


;*** compare proc for sorting list view items


comparelvproc proc public uses ebx lParam1:dword,lParam2:dword,pLVC:ptr

local	lvi:LVITEM
local	szStr1[MAX_PATH]:byte
local	szStr2[MAX_PATH]:byte
local	dwTmp1:dword
local	dwTmp2:dword

		mov ebx,pLVC

		movzx eax,[ebx].LVCOMP.iSortCol
		mov lvi.iSubItem,eax

		mov eax,lParam1
		mov lvi.iItem,eax
		mov lvi.mask_,LVIF_TEXT
		lea eax,szStr1
		mov lvi.pszText,eax
		mov lvi.cchTextMax,sizeof szStr1
		invoke ListView_GetItem( [ebx].LVCOMP.hWndLV,addr lvi)

		mov eax,lParam2
		mov lvi.iItem,eax
		lea eax,szStr2
		mov lvi.pszText,eax
		invoke ListView_GetItem( [ebx].LVCOMP.hWndLV,addr lvi)

		.if ([ebx].LVCOMP.dwFlags)
			invoke String2DWord, addr szStr1, addr dwTmp1
			invoke String2DWord, addr szStr2, addr dwTmp2
			mov eax,dwTmp1
			mov ecx,dwTmp2
			.if ([ebx].LVCOMP.iSortDir == 1)
				xchg eax,ecx
			.endif
			sub eax,ecx
		.else
			.if ([ebx].LVCOMP.iSortDir == 0)
				invoke lstrcmp,addr szStr1,addr szStr2
			.else
				invoke lstrcmp,addr szStr2,addr szStr1
			.endif
		.endif
		ret
		align 4

comparelvproc endp


;--- translate TYPEKIND into readable string


GetTypekindStr proc public uses edi tk:dword

		mov eax,tk
		mov ecx,NUMTYPEKIND
		mov edi,offset dwTypeKind
		repnz scasd
		sub ecx,NUMTYPEKIND-1
		neg ecx
		mov eax,[ecx*4+offset pszTypeKind]
		ret
		align 4
GetTypekindStr endp


;--- translate variable type into readable string


GetVarType proc public dwVarType:DWORD
		mov eax,dwVarType
		@mov ecx,0
		.while (ecx < NUMVT)
			.break .if (ax == word ptr [ecx * sizeof word + offset pVARTYPE])
			inc ecx
		.endw
		mov eax,[ecx * sizeof dword + offset pVARTYPEStr]
		ret
		align 4
GetVarType endp


;--- translate variable type into ASM readable string


GetVarTypeAsm proc public dwVarType:DWORD
		mov eax,dwVarType
		@mov ecx,0
		.while (ecx < NUMVT)
			.break .if (ax == word ptr [ecx * sizeof word + offset pVARTYPE])
			inc ecx
		.endw
		mov eax,[ecx * sizeof dword + offset pVARTYPEStrAsm]
		ret
		align 4

GetVarTypeAsm endp


;--- avoid some ASM reserved names in declarations
;--- by simply add a underscore to it

CheckReservedNames	proc public uses esi pStr:ptr byte

		mov esi, offset ReservedNames
		mov ecx, SIZERESERVEDNAMES
		.while ecx
			push ecx
			lodsd
			invoke lstrcmpi, pStr, eax
			pop ecx
			.break .if (eax == 0)
			dec ecx
		.endw
		.if (ecx)
			invoke lstrcat, pStr, CStr("_")
		.endif
		ret

CheckReservedNames	endp


GetUserDefinedSize proc uses esi pTypeInfo: LPTYPEINFO, hRef:HREFTYPE

local	pTI:LPTYPEINFO
local	pTypeAttr:ptr TYPEATTR
local	rc:DWORD

		mov rc,0
		invoke vf(pTypeInfo,ITypeInfo,GetRefTypeInfo), hRef, addr pTI
		.if (eax == S_OK)
			invoke vf(pTI, ITypeInfo, GetTypeAttr), addr pTypeAttr
			.if (eax == S_OK)
				mov eax, pTypeAttr
				mov eax, [eax].TYPEATTR.cbSizeInstance
				mov rc, eax
			.endif
			invoke vf(pTI, ITypeInfo, Release)
		.endif
		return rc
		align 4

GetUserDefinedSize endp

;*** translate type VT_USERDEFINED: search ITypeInfo and get name


GetUserDefinedType proc pTypeInfo: LPTYPEINFO, hRef:HREFTYPE, pStr:LPSTR, iMax:dword

local	pTI:LPTYPEINFO
local	bstr:BSTR
local	szName[64]:byte
local	rc:DWORD

		mov rc,0
		invoke vf(pTypeInfo,ITypeInfo,GetRefTypeInfo), hRef, addr pTI
		.if (eax == S_OK)
			invoke vf(pTI, ITypeInfo, GetDocumentation), MEMBERID_NIL, addr bstr, NULL, NULL, NULL
			.if (eax == S_OK)
				invoke WideCharToMultiByte,CP_ACP,0,bstr,-1,addr szName,sizeof szName,0,0 
				invoke SysFreeString, bstr
				invoke lstrcat, pStr, addr szName
				mov rc,1
			.endif
			invoke vf(pTI, ITypeInfo, Release)
		.endif
		ret
		align 4

GetUserDefinedType endp

GPT_STD		equ 0
GPT_ASM		equ 1
GPT_STUB	equ 2

;*** translate type of parameter into readable string
;*** iMode = GPT_STUB is: we want a VARIANTARG for IDispatch::Invoke

GetParameterTypeEx proc uses esi pTypeInfo:LPTYPEINFO, pTypeDesc:ptr TYPEDESC, pStr:LPSTR, iMax:dword, iMode:dword

local szText[32]:byte

        mov esi,pTypeDesc
		movzx eax,[esi].TYPEDESC.vt
		.if (eax == VT_USERDEFINED)
			mov eax,pStr
			mov byte ptr [eax],0
			.if (iMode == GPT_STUB)
;----------------------------------------- dont know what to make with this
;----------------------------------------- should we set VT_USERDEFINED as
;----------------------------------------- VARIANTARG type?
;;				invoke lstrcpy, pStr, CStr("I4")
				invoke lstrcpy, pStr, CStr("USERDEFINED")
			.else
				invoke GetUserDefinedType, pTypeInfo, [esi].TYPEDESC.hreftype, pStr, iMax
				.if (iMode)
					invoke CheckReservedNames, pStr
				.endif
			.endif
		.elseif ((eax == VT_CARRAY) && (iMode == GPT_ASM))
			mov eax, [esi].TYPEDESC.lpadesc
			invoke GetParameterTypeEx, pTypeInfo, addr [eax].ARRAYDESC.tdescElem, pStr, iMax, iMode
			mov eax, [esi].TYPEDESC.lpadesc
			invoke wsprintf, addr szText, CStr(" %u dup"), [eax].ARRAYDESC.rgbounds.cElements
			invoke lstrcat, pStr, addr szText
		.else
			.if (iMode == GPT_ASM)
				invoke GetVarTypeAsm, eax
			.else
				invoke GetVarType, eax
			.endif
			invoke lstrcpy, pStr, eax
		.endif
		.while ([esi].TYPEDESC.vt == VT_PTR)
			.if (iMode == GPT_STUB)
				invoke lstrcpy, pStr, CStr("BYREF or VT_")
			.else
				invoke lstrcat, pStr, CStr(" ")
			.endif
			mov esi,[esi].TYPEDESC.lptdesc
			movzx eax,[esi].TYPEDESC.vt
			.if (eax == VT_USERDEFINED)
				.if (iMode == GPT_STUB)
					invoke lstrcat, pStr, CStr("USERDEFINED")
				.else
					invoke lstrlen, pStr
					push eax
					invoke GetUserDefinedType, pTypeInfo, [esi].TYPEDESC.hreftype, pStr, iMax
					pop eax
					.if (iMode)
						mov ecx, pStr
						add ecx, eax
						invoke CheckReservedNames, ecx
					.endif
				.endif
			.else
				.if (iMode == GPT_ASM)
					invoke GetVarTypeAsm, eax
				.else
					invoke GetVarType, eax
				.endif
				invoke lstrcat, pStr, eax		;!!!
			.endif
			.break .if (iMode == GPT_STUB)
		.endw
		ret
		align 4

GetParameterTypeEx endp


GetParameterType proc public pTypeInfo:LPTYPEINFO, pTypeDesc:ptr TYPEDESC, pStr:LPSTR, iMax:dword

		invoke GetParameterTypeEx, pTypeInfo, pTypeDesc, pStr, iMax, GPT_STD
		ret
		align 4
GetParameterType endp


GetParameterTypeAsm proc public pTypeInfo:LPTYPEINFO, pTypeDesc:ptr TYPEDESC, pStr:LPSTR, iMax:dword

		invoke GetParameterTypeEx, pTypeInfo, pTypeDesc, pStr, iMax, GPT_ASM
		ret
		align 4
GetParameterTypeAsm endp

GetParameterSize proc public uses esi pTypeInfo:LPTYPEINFO, pTypeDesc:ptr TYPEDESC

local dwElemSize:DWORD

        mov esi,pTypeDesc
		movzx eax,[esi].TYPEDESC.vt
		.if (eax == VT_USERDEFINED)

			invoke GetUserDefinedSize, pTypeInfo, [esi].TYPEDESC.hreftype

		.elseif (eax == VT_CARRAY)

			mov eax, [esi].TYPEDESC.lpadesc
			invoke GetParameterSize, pTypeInfo, addr [eax].ARRAYDESC.tdescElem
			mov dwElemSize, eax
			mov ecx, [esi].TYPEDESC.lpadesc
			mov ecx, [ecx].ARRAYDESC.rgbounds.cElements
			mul ecx

		.elseif (eax == VT_PTR)

			mov eax, 4

		.else

			.if (eax == VT_VARIANT)
				mov eax, sizeof VARIANT
			.elseif (eax == VT_DECIMAL)
				mov eax, 16
			.elseif ((eax == VT_R8) || (eax == VT_CY) || (eax == VT_DATE) || (eax == VT_I8) || (eax == VT_UI8))
				mov eax, 8
			.elseif ((eax == VT_I2) || (eax == VT_UI2) || (eax == VT_BOOL))
				mov eax, 2
			.elseif ((eax == VT_I1) || (eax == VT_UI1))
				mov eax, 1
			.else
				mov eax, 4
			.endif

		.endif
		ret
		align 4

GetParameterSize endp


GetParameterTypeStub proc public pTypeInfo:LPTYPEINFO, pTypeDesc:ptr TYPEDESC, pStr:LPSTR, iMax:dword

		invoke lstrcpy, pStr, CStr("VT_")
		add pStr,3
		sub iMax,3
		invoke GetParameterTypeEx, pTypeInfo, pTypeDesc, pStr, iMax, GPT_STUB
		invoke CharUpper, pStr
		ret
		align 4

GetParameterTypeStub endp

;--- get default interface from a TKIND_COCLASS typeinfo

GetDefaultInterfaceFromCoClass proc public uses esi pTypeInfo:LPTYPEINFO, bSource:BOOL

local	pTypeAttr:ptr TYPEATTR
local	dwIndex:DWORD
local	dwFlags:DWORD
local	hRefType:HREFTYPE
local	pTypeInfoRef:LPTYPEINFO

		mov pTypeInfoRef, NULL
		invoke vf(pTypeInfo, ITypeInfo, GetTypeAttr),addr pTypeAttr
		.if (eax == S_OK)
			mov esi, pTypeAttr
			movzx ecx, [esi].TYPEATTR.cImplTypes
			mov dwIndex,0
			.while (ecx)
				push ecx
				invoke vf(pTypeInfo, ITypeInfo, GetImplTypeFlags), dwIndex, addr dwFlags
				pop ecx
				mov eax, dwFlags
				and eax, IMPLTYPEFLAG_FDEFAULT or IMPLTYPEFLAG_FSOURCE
				.if (bSource)
					.break .if (eax == (IMPLTYPEFLAG_FDEFAULT or IMPLTYPEFLAG_FSOURCE))
				.else
					.break .if (eax == IMPLTYPEFLAG_FDEFAULT)
				.endif
				inc dwIndex
				dec ecx
			.endw
			.if (ecx)
				invoke vf(pTypeInfo, ITypeInfo, GetRefTypeOfImplType), dwIndex, addr hRefType
				.if (eax == S_OK)
					invoke vf(pTypeInfo, ITypeInfo, GetRefTypeInfo), hRefType, addr pTypeInfoRef
				.endif
			.endif
			invoke vf(pTypeInfo, ITypeInfo, ReleaseTypeAttr), pTypeAttr
		.endif
		return pTypeInfoRef

GetDefaultInterfaceFromCoClass endp

if 0
GetTypeInfoFromCLSID proc public pClsId:ptr CLSID

local	pTypeInfo:LPTYPEINFO
local	hKey:HANDLE
local	dwSize:DWORD
local	wszGUID[40]:word
local	szGUID[40]:byte
local	szKey[MAX_PATH]:byte

		invoke StringFromGUID2, pClsId,addr wszGUID,40
		invoke WideCharToMultiByte,CP_ACP,0,addr wszGUID,-1,addr szGUID,sizeof szGUID,0,0
		invoke wsprintf, addr szKey, CStr("%s\%s\%s"), addr g_szRootCLSID, addr szGUID, addr g_szTypeLib
		invoke RegOpenKeyEx, HKEY_CLASSES_ROOT, addr szKey,0,KEY_READ,addr hKey
		.if (eax == ERROR_SUCCESS)
			mov dwSize, sizeof szKey
			invoke RegQueryValueEx, hKey, addr g_szNull, NULL, NULL, addr szKey, addr dwSize
			invoke RegCloseKey, hKey
		.endif
		return pTypeInfo
GetTypeInfoFromCLSID endp
endif

;--- get a LPTYPEINFO from an LPUNKNOWN thru IProvideClassInfo


GetTypeInfoFromIProvideClassInfo proc public uses esi pUnknown:LPUNKNOWN, bSource:BOOL

local	pPCI:LPPROVIDECLASSINFO
local	pTypeInfo:LPTYPEINFO
local	pTypeAttr:ptr TYPEATTR
local	dwIndex:DWORD
local	dwFlags:DWORD
local	hRefType:HREFTYPE
local	pTypeInfoRef:LPTYPEINFO

		mov pTypeInfoRef, NULL
		invoke vf(pUnknown, IUnknown, QueryInterface), addr IID_IProvideClassInfo, addr pPCI
		.if (eax == S_OK)
			invoke vf(pPCI, IProvideClassInfo, GetClassInfo_), addr pTypeInfo
			.if (eax == S_OK)
				invoke GetDefaultInterfaceFromCoClass, pTypeInfo, bSource
				mov pTypeInfoRef, eax
				invoke vf(pTypeInfo, IUnknown, Release)
			.endif
			invoke vf(pPCI, IUnknown, Release)
		.endif
		return pTypeInfoRef
		align 4

GetTypeInfoFromIProvideClassInfo endp


;--- start a typeinfo dialog with info from IProvideClassInfo


TypeInfoDlgFromIProvideClassInfo proc public hWnd:HWND, pUnknown:LPUNKNOWN

local	pTID:ptr CTypeInfoDlg
local	pIPCI:ptr IProvideClassInfo
local	pTypeInfo:LPTYPEINFO
local	dwIndex:DWORD

		mov pTID, NULL
		invoke vf(pUnknown, IUnknown, QueryInterface),addr IID_IProvideClassInfo, addr pIPCI
		.if (eax == S_OK)
			invoke vf(pIPCI,IProvideClassInfo,GetClassInfo_),addr pTypeInfo
			.if (eax == S_OK)
				invoke GetTypeInfoFromIProvideClassInfo, pUnknown, FALSE
				.if (eax)
					push eax
					invoke vf(pTypeInfo, IUnknown, Release)
					pop pTypeInfo
					mov dwIndex, 0
				.else
					mov dwIndex, 2
				.endif
				invoke Create2@CTypeInfoDlg, pTypeInfo
				.if (eax)
					mov pTID,eax
					invoke SetTab@CTypeInfoDlg, pTID, dwIndex
					invoke Show@CTypeInfoDlg, pTID, hWnd, TYPEINFODLG_FACTIVATE
;;					invoke Destroy@CTypeInfoDlg, pTID
				.endif
				invoke vf(pTypeInfo,ITypeInfo,Release)
			.else
				mov ecx,eax
				invoke OutputMessage, hWnd, ecx, CStr("IProvideClassInfo::GetClassInfo failed"),0
			.endif
			invoke vf(pIPCI,IProvideClassInfo,Release)
		.else
			mov ecx, eax
			invoke OutputMessage, hWnd,ecx,CStr("QueryInterface(IProvideClassInfo) failed"),0
		.endif
		return pTID

TypeInfoDlgFromIProvideClassInfo endp


;--- get a LPTYPEINFO from an LPUNKNOWN thru IDispatch


GetTypeInfoFromIDispatch proc public uses esi pUnknown:LPUNKNOWN

local	pDispatch:LPDISPATCH
local	pTypeInfo:LPTYPEINFO

		mov pTypeInfo, NULL
		invoke vf(pUnknown, IUnknown, QueryInterface), addr IID_IDispatch, addr pDispatch
		.if (eax == S_OK)
			invoke vf(pDispatch, IDispatch, GetTypeInfo), NULL, g_LCID, addr pTypeInfo
			invoke vf(pDispatch, IUnknown, Release)
		.endif
		return pTypeInfo

GetTypeInfoFromIDispatch endp


;--- start a typeinfo dialog with info from IDispatch::GetTypeInfo


TypeInfoDlgFromIDispatch proc public hWnd:HWND, pUnknown:LPUNKNOWN

local	cntTI:DWORD
local	pTypeInfo:LPTYPEINFO
local	pTID:ptr CTypeInfoDlg
local	pDispatch:LPDISPATCH
local	dwIndex:DWORD
local	pTypeAttr:ptr TYPEATTR
local	szStr[260]:byte

		mov pTID, NULL
		invoke vf(pUnknown, IUnknown, QueryInterface), addr IID_IDispatch, addr pDispatch
		.if (eax == S_OK)
			invoke vf(pDispatch, IDispatch, GetTypeInfoCount), addr cntTI
			.if ((eax == S_OK) && cntTI)
				invoke vf(pDispatch, IDispatch, GetTypeInfo), 0, g_LCID, addr pTypeInfo
				.if (eax == S_OK)
					invoke Create2@CTypeInfoDlg, pTypeInfo
					.if (eax)
						mov pTID,eax
						@mov dwIndex, 0
						invoke vf(pTypeInfo, ITypeInfo, GetTypeAttr), addr pTypeAttr
						.if (eax == S_OK)
							mov ecx,pTypeAttr
							.if ([ecx].TYPEATTR.typekind == TKIND_COCLASS)
								mov dwIndex,2
							.elseif ([ecx].TYPEATTR.typekind == TKIND_ENUM)
								mov dwIndex,1
							.endif
							invoke vf(pTypeInfo, ITypeInfo, ReleaseTypeAttr), pTypeAttr
						.endif
						invoke SetTab@CTypeInfoDlg, pTID, dwIndex
						invoke Show@CTypeInfoDlg, pTID, hWnd, TYPEINFODLG_FACTIVATE
					.endif
					invoke vf(pTypeInfo,ITypeInfo,Release)
				.else
					mov ecx, eax
					invoke OutputMessage, hWnd, ecx, CStr("IDispatch::GetTypeInfo failed"), 0
				.endif
			.else
				invoke wsprintf, addr szStr, CStr("TypeInfoCount returned %X (cnt=%u)"), eax, cntTI
				invoke MessageBox, hWnd, addr szStr, 0, MB_OK
			.endif
			invoke vf(pDispatch, IDispatch, Release)
		.else
			mov ecx, eax
			invoke OutputMessage, hWnd, ecx, CStr("QueryInterface(IDispatch) failed"), 0
		.endif
		return pTID

TypeInfoDlgFromIDispatch endp


;*** dialog proc for all class dialogs (this ptr saved in DWL_USER)


classdialogproc proc public hWnd:HWND,message:dword,wParam:WPARAM,lParam:LPARAM

ifdef _DEBUG
		push ebx
		push esi
		push edi
endif
		.if (message == WM_INITDIALOG)
			invoke SendMessage, hWnd, WM_SETICON, ICON_SMALL, g_hIconApp
			invoke SendMessage, hWnd, WM_SETICON, ICON_BIG, g_hIconApp
			invoke SetWindowLong, hWnd, DWL_USER, lParam
            mov eax,lParam
			mov ecx,hWnd
			mov [eax].CDlg.hWnd,ecx
		.else
			invoke GetWindowLong,hWnd,DWL_USER
		.endif
        .if (eax != 0)
ifdef _DEBUG
;---------------------- is in a free heap memory item
			.if (eax > 80000000h)
				invoke DebugBreak
			.endif
if 0
			push eax
			invoke HeapValidate, g_heap, HEAP_NO_SERIALIZE, eax
			pop eax
endif
endif
			invoke [eax.CDlg.pDlgProc],eax,message,wParam,lParam
			.if (message == WM_DESTROY)
				push eax
				invoke SetWindowLong, hWnd, DWL_USER, NULL
				pop eax
			.endif
        .endif
ifdef _DEBUG
		.if ((edi != [esp+0]) || (esi != [esp+4]) || (ebx != [esp+8]))
			invoke DebugBreak
		.endif
		pop edi
		pop esi
		pop ebx
endif
		ret
		align 4

classdialogproc endp


;--- setup arguments: will set g_argc and g_argv global vars


SetArguments proc public uses esi edi ebx ebp

;local	argc:dword

		invoke GetCommandLine
		and eax,eax
		jz exit
		mov esi,eax
		xor edi,edi			;EDI will count the number of arguments
		xor edx,edx			;EDX will count the number of bytes
							;needed for the arguments
							;(not including the null terminators)
nextarg:					;<---- get next argument
		.while (1)
			lodsb
			.break .if ((al != ' ') && (al != 9))	;skip spaces and tabs
		.endw
		or al,al
		je donescanX		;done commandline scan
		inc edi 			;Another argument
		xor ebx,ebx 		;EBX will count characters in argument
		dec esi 			;back up to reload character
		push esi 			;save start of argument
		mov cl,00
		.while (1)
			lodsb
			.break .if (!al)
			.if (!cl)
				.if ((al == ' ') || (al == 9))	;white space term. argument
					push	ebx 				;save argument length
					jmp		nextarg
				.endif
				.if ((!ebx) && al == '"')	;starts argument with "?
					or cl,1
										;handle argument beginning with doublequote
					pop eax				;throw away old start
					push esi 			;and set new start
					.continue
				.endif
			.elseif (al == '"')
				and cl,0FEh
				.continue
			.endif

			.if ((al == '\')  && (byte ptr [esi] == '"'))
				inc esi
			.endif
			inc ebx
			inc edx 		;one more space
		.endw
		push ebx 			; save length of last argument
donescanX:
;		mov argc,edi		; Store number of arguments
		mov ebp,edi			; Store number of arguments
		add edx,edi 		; add terminator bytes
		inc edi 			; add one for NULL pointer
		shl edi,2			; every pointer takes 4 bytes
		add edx,edi 		; add that space to space for strings

		invoke malloc, edx
		and eax,eax
		jz exit

		mov g_argv,eax
		add edi,eax 		; edi -> behind vector table (strings)
;		mov ecx,argc
		mov ecx,ebp
		mov g_argc,ecx
		lea ebx,[edi-4]
		mov dword ptr [ebx],0 ;mark end of argv
		sub ebx,4
		mov edx,ecx
		.while (edx)
			pop ecx 		;get length
			pop esi 		;get address
			mov [ebx],edi
			sub ebx,4
			.while (ecx)
				lodsb
				.if (al == '\')
					.continue .if (byte ptr [esi] == '"')
				.endif
				stosb
				dec ecx
			.endw
			xor al,al
			stosb
			dec edx
		.endw
exit:
		ret
		align 4

SetArguments endp

BroadCastMessageStruct struct
uMsg	DWORD	?
wParam	WPARAM	?
lParam	LPARAM	?
BroadCastMessageStruct ends

EnumThreadWindowsCB proc hWnd:HWND, lParam:LPARAM
	
	mov ecx, lParam
	assume ecx:ptr BroadCastMessageStruct
	.if ([ecx].uMsg == WM_CLOSE)
		mov edx, [ecx].wParam
		.if (edx != hWnd)
			invoke IsWindowVisible, hWnd
			.if (eax)
				invoke SendMessage, hWnd, WM_CLOSE, 0, 0
			.endif
		.endif
	.else
		invoke SendMessage, hWnd, [ecx].uMsg, [ecx].wParam, [ecx].lParam
	.endif
	return TRUE
	assume ecx:nothing

EnumThreadWindowsCB endp

;--- broadcast a WM_COMMAND message to all thread windows

BroadCastMessage proc public uMsg:DWORD, wParam:WPARAM, lParam:LPARAM

local bcms:BroadCastMessageStruct

	invoke GetCurrentThreadId
	lea ecx, uMsg
	invoke EnumThreadWindows, eax, offset EnumThreadWindowsCB, ecx
	ret
BroadCastMessage endp

RestoreAndActivateWindow proc public hWnd:HWND

	invoke ShowWindow, hWnd, SW_RESTORE
	invoke SetActiveWindow, hWnd
	ret

RestoreAndActivateWindow endp


;--- clear ParamReturn structure


ParamReturnClear proc public uses esi edi ebx pPR:ptr PARAMRETURN

		mov edi, pPR
		.if ([edi].PARAMRETURN.pVariants)
			.if ([edi].PARAMRETURN.iNumVariants > 100)
				invoke SetCursor, g_hCsrWait
				mov ebx, eax
			.endif
			mov ecx, [edi].PARAMRETURN.iNumVariants
			mov esi, [edi].PARAMRETURN.pVariants
			.while (ecx)
				push ecx
ifdef _DEBUG
				movzx eax, [esi].VARIANT.vt
				DebugOut "ParamReturnClear %X, %X", eax, [esi].VARIANT.byref
endif
				.if ([esi].VARIANT.vt & VT_BYREF)
					invoke free, [esi].VARIANT.byref
					mov [esi].VARIANT.byref, NULL
				.endif
				invoke VariantClear, esi
				add esi, sizeof VARIANT
				pop ecx
				dec ecx
			.endw
			invoke free, [edi].PARAMRETURN.pVariants
			.if ([edi].PARAMRETURN.iNumVariants > 100)
				invoke SetCursor, ebx
			.endif
			mov [edi].PARAMRETURN.pVariants, NULL
			mov [edi].PARAMRETURN.iNumVariants, 0

		.endif
		ret
		align 4

ParamReturnClear endp

	.const

@DefDispId macro x:vararg
local string
	for r,<x>
string textequ @CatStr(!",r,!")
	dd DISPID_&r
	dd CStr(string)
	endm
	endm

	align 4

StdDispIdTab	label dword

;--- -600

	@DefDispId CLICK, DBLCLICK
	@DefDispId KEYDOWN, KEYPRESS, KEYUP
	@DefDispId MOUSEDOWN, MOUSEMOVE, MOUSEUP
	@DefDispId ERROREVENT
	@DefDispId READYSTATECHANGE
	@DefDispId CLICK_VALUE
	@DefDispId RIGHTTOLEFT
	@DefDispId TOPTOBOTTOM
	@DefDispId THIS

;--- -700

AmbientDispIdTab	label dword
	@DefDispId AMBIENT_BACKCOLOR
	@DefDispId AMBIENT_DISPLAYNAME
	@DefDispId AMBIENT_FONT
	@DefDispId AMBIENT_FORECOLOR
	@DefDispId AMBIENT_LOCALEID
	@DefDispId AMBIENT_MESSAGEREFLECT
	@DefDispId AMBIENT_SCALEUNITS
	@DefDispId AMBIENT_TEXTALIGN
	@DefDispId AMBIENT_USERMODE
	@DefDispId AMBIENT_UIDEAD
	@DefDispId AMBIENT_SHOWGRABHANDLES
	@DefDispId AMBIENT_SHOWHATCHING
	@DefDispId AMBIENT_DISPLAYASDEFAULT
	@DefDispId AMBIENT_SUPPORTSMNEMONICS
	@DefDispId AMBIENT_AUTOCLIP
	@DefDispId AMBIENT_APPEARANCE
	@DefDispId AMBIENT_CODEPAGE
	@DefDispId AMBIENT_PALETTE
	@DefDispId AMBIENT_CHARSET
	@DefDispId AMBIENT_TRANSFERPRIORITY
	@DefDispId AMBIENT_RIGHTTOLEFT
	@DefDispId AMBIENT_TOPTOBOTTOM
	@DefDispId AMBIENT_DLCONTROL
	@DefDispId AMBIENT_USERAGENT
	@DefDispId AMBIENT_OFFLINEIFNOTCONNECTED
	@DefDispId AMBIENT_SILENT

;--- -500
if 0
	@DefDispId AUTOSIZE, BACKCOLOR, BACKSTYLE
	@DefDispId BORDERCOLOR, BORDERSTYLE, BORDERWIDTH
	@DefDispId DRAWMODE, DRAWSTYLE, DRAWWIDTH
	@DefDispId FILLCOLOR, FILLSTYLE
	@DefDispId FONT, FORECOLOR
	@DefDispId ENABLED, HWND
	@DefDispId TABSTOP, TEXT
	@DefDispId CAPTION
endif

NUMAMBIENTPROPS equ ($ - offset AmbientDispIdTab) / (sizeof DWORD * 2)

	dd 0

	.code

GetStdDispIdStr	proc public uses esi dispid:sdword

	mov esi,offset StdDispIdTab
	mov	ecx,dispid
@@:
	lodsd
	mov edx,eax
	lodsd
	cmp	edx,ecx
	jz	done
	xor eax,eax
	cmp	edx,eax
	jnz	@B
done:
	ret
GetStdDispIdStr endp

if ?SHOWAMBIENT

GetAmbientDispId proc public dwIndex:DWORD, pDispId:ptr DWORD, pDispIdStr:ptr LPSTR
	mov ecx, dwIndex
	cmp ecx, NUMAMBIENTPROPS
	jae error
	mov eax,[ecx*8 + 0 + offset AmbientDispIdTab]
	mov edx,[ecx*8 + 4 + offset AmbientDispIdTab]
	mov ecx, pDispId
	mov [ecx],eax
	mov ecx, pDispIdStr
	mov [ecx],edx
	ret
error:
	xor eax, eax
	ret
	align 4
GetAmbientDispId endp
endif

;--- display arguments of event logs


GetArgument proc public pVariant:ptr VARIANT, pStrOut:LPSTR

local	szValue[64]:byte

	mov ecx,1
	mov eax,pVariant
	.if ([eax].VARIANT.vt == VT_I4)
		mov edx,[eax].VARIANT.lVal
	.elseif ([eax].VARIANT.vt == VT_UI4)
		mov edx,[eax].VARIANT.ulVal
	.elseif ([eax].VARIANT.vt == VT_I2)
		movsx edx,[eax].VARIANT.iVal
	.elseif ([eax].VARIANT.vt == VT_UI2)
		movzx edx,[eax].VARIANT.uiVal
	.elseif ([eax].VARIANT.vt == VT_I1)
		movsx edx,[eax].VARIANT.cVal
	.elseif ([eax].VARIANT.vt == VT_UI1)
		movzx edx,[eax].VARIANT.bVal
	.elseif ([eax].VARIANT.vt == VT_BOOL)
		movsx edx,[eax].VARIANT.boolVal
	.else
		mov ecx,0
	.endif
	.if (ecx)
		invoke wsprintf, addr szValue, CStr("%d"), edx
		invoke lstrcat, pStrOut, addr szValue
		ret
	.endif
	mov edx,pVariant
	.if ([edx].VARIANT.vt == VT_BSTR)
		invoke WideCharToMultiByte, CP_ACP, 0, [edx].VARIANT.bstrVal, -1, addr szValue + 1, sizeof szValue - 3,0,0 
		mov byte ptr szValue,'"'
		mov word ptr [eax+szValue],'"'
		invoke lstrcat, pStrOut, addr szValue
		ret
	.endif
	invoke wsprintf, addr szValue, CStr("0x%X"),[edx].VARIANT.uintVal
	invoke lstrcat,pStrOut,addr szValue
	ret
GetArgument endp


;--- seems not to be implemented in Win9x

_strcmpW proc public uses esi edi pszwStr1:ptr WORD, pszwStr2:ptr WORD
	mov esi, pszwStr1
	mov edi, pszwStr2
	xor eax, eax
	.while (ZERO?)
		lodsw
		.break .if (!ax)
		scasw
	.endw
	.if (!ax)
		sub ax,[edi]
	.endif
	ret
_strcmpW endp


CopyStringToClipboard proc public hWnd:HWND, pszGUID:LPSTR

local hGlobal:HGLOBAL
local pGlobal:ptr byte
;local pszGUID:LPSTR

		invoke lstrlen, pszGUID
		inc eax
		invoke GlobalAlloc,GMEM_MOVEABLE or GMEM_DDESHARE, eax
		mov hGlobal,eax
		.if (eax != 0)
			invoke GlobalLock,hGlobal
			.if (eax)
				invoke lstrcpy, eax, pszGUID
				invoke GlobalUnlock,hGlobal
				invoke OpenClipboard, hWnd
				invoke EmptyClipboard
				invoke SetClipboardData,CF_TEXT,hGlobal
				invoke CloseClipboard
			.endif
		.endif
		ret
		align 4

CopyStringToClipboard endp

CopyFileToClipboard proc public uses ebx hWnd:HWND, pszFile:LPSTR

local hGlobal:HGLOBAL
local dwSize:DWORD
local dwRead:DWORD
local pGlobal:ptr byte

		invoke CreateFile, pszFile,GENERIC_READ,\
				FILE_SHARE_READ or FILE_SHARE_WRITE,\
				NULL,OPEN_EXISTING,0,0
		.if (eax != -1)
			mov ebx, eax
			invoke GetFileSize, ebx, NULL
			mov dwSize, eax
			inc eax
			invoke GlobalAlloc,GMEM_MOVEABLE or GMEM_DDESHARE, eax
			mov hGlobal,eax
			.if (eax != 0)
				invoke GlobalLock,hGlobal
				.if (eax)
					mov ecx, eax
					mov eax, dwSize
					mov byte ptr [ecx+eax],0
					invoke ReadFile, ebx, ecx, dwSize, addr dwRead, NULL
					invoke GlobalUnlock,hGlobal
					invoke OpenClipboard, hWnd
					invoke EmptyClipboard
					invoke SetClipboardData,CF_TEXT, hGlobal
					invoke CloseClipboard
				.endif
			.endif
			invoke CloseHandle, ebx
			mov eax, ebx
		.endif
		ret
		align 4

CopyFileToClipboard endp

GetItemPosition proc public hWnd:HWND, bMouse:BOOL, pPoint:ptr POINT

local	rect:RECT

		.if (bMouse)
			invoke GetCursorPos, pPoint
		.else
			invoke ListView_GetNextItem( hWnd, -1, LVNI_SELECTED)
			.if (eax == -1)
				inc eax
			.endif
			lea ecx, rect
			invoke ListView_GetItemRect( hWnd, eax, ecx, LVIR_BOUNDS)
			.if (!eax)
				invoke GetClientRect, hWnd, addr rect
			.endif
			mov eax, rect.left
			mov ecx, rect.top
			mov edx, rect.bottom
			sub edx, ecx
			shr edx, 1
			add eax, edx
			add ecx, edx
			mov edx, pPoint
			mov [edx].POINT.x, eax
			mov [edx].POINT.y, ecx
			invoke ClientToScreen, hWnd, pPoint
		.endif
		ret
GetItemPosition endp

SetWindowIcon proc public hWnd:HWND, pguid:ptr GUID
if ?SHOWICON
local	hIcon:HICON
local   hKey:HANDLE
local   dwType:dword
local   dwSize:dword
local	dwIndex:dword
local	wszGUID[40]:word
;;local	szGUID[40]:byte
local	szText[MAX_PATH]:byte
local	szStr[MAX_PATH]:byte

		mov hIcon, NULL
		invoke StringFromGUID2, pguid, addr wszGUID, 40
;;		invoke WideCharToMultiByte,CP_ACP,0,addr wszGUID,-1,addr szGUID,sizeof szGUID,0,0
		invoke wsprintf, addr szStr, CStr("%s\%S\DefaultIcon"), addr g_szRootCLSID, addr wszGUID
		invoke RegOpenKeyEx, HKEY_CLASSES_ROOT, addr szStr,0,KEY_READ,addr hKey
		.if (eax == ERROR_SUCCESS)
			mov dwSize, sizeof szText
			invoke RegQueryValueEx, hKey, addr g_szNull, NULL, addr dwType, addr szText, addr dwSize
			.if (eax == ERROR_SUCCESS)
				.if (dwType == REG_EXPAND_SZ)
					invoke lstrcpy, addr szStr, addr szText
					invoke ExpandEnvironmentStrings, addr szStr, addr szText, MAX_PATH
				.endif
				lea ecx, szText
				mov eax, dwSize
				@mov dwIndex, 0
				.while (eax)
					dec eax
					.if (byte ptr [eax+ecx] == ",")
						lea edx, [eax+ecx]
						mov byte ptr [edx],0
						inc edx
						invoke String2DWord, edx, addr dwIndex
						.break
					.endif
				.endw
				invoke ExtractIcon, g_hInstance, addr szText, dwIndex
				.if (eax)
					mov hIcon, eax
				.endif
			.endif
			invoke RegCloseKey, hKey
		.endif
done:
		.if (hIcon)
			invoke SendMessage, hWnd, WM_SETICON, ICON_BIG, hIcon
			invoke SendMessage, hWnd, WM_SETICON, ICON_SMALL, hIcon
		.endif
		mov eax, hIcon
endif
		ret
SetWindowIcon endp

;--- set statusbar parts

ifdef @StackBase
	option stackbase:ebp
endif

SetSBParts proc public uses edi esi ebx hWndSB:HWND, pParts:ptr DWORD, iNumParts:DWORD

local	rect:RECT

		invoke GetClientRect, hWndSB, addr rect
		mov esi, pParts 
		mov eax, iNumParts
		mov ecx, eax
		dec ecx
		shl eax, 2
		sub esp, eax
		mov edi, esp
		xor ebx, ebx
		.while (ecx)
			push ecx
			lodsd
			invoke MulDiv, rect.right, eax, 100
			add eax, ebx
			stosd
			mov ebx, eax
			pop ecx
			dec ecx
		.endw
		mov eax, -1
		stosd

		StatusBar_SetParts hWndSB, iNumParts, esp

		mov esp, edi
		ret
		align 4

SetSBParts endp

ifdef @StackBase
	option stackbase:esp
endif

GetDirOnly proc pszDir:LPSTR, pszFile:LPSTR
		invoke lstrcpy, pszDir, pszFile
		invoke lstrlen, pszDir
		mov ecx, eax
		mov eax, pszDir
		.while (ecx)
			.if (byte ptr [eax+ecx-1] == '\')
				mov byte ptr [eax+ecx-1], 0
				.break
			.endif
			dec ecx
		.endw
		mov eax, ecx
		ret
GetDirOnly endp


;*** get a file name


MyGetFileName proc public hWnd:HWND, pszFileName:LPSTR, iMax:DWORD, pszCustom:LPSTR, iMaxCustom:DWORD, iFlags:DWORD, pszTitle:LPSTR

local	ofn:OPENFILENAME
local	szFilter[128]:byte
local	szPath[MAX_PATH]:byte
local	hWndDlg:HWND

		invoke ZeroMemory,addr ofn,sizeof OPENFILENAME
		mov ofn.lStructSize,sizeof OPENFILENAME
		mov eax, hWnd
		mov ofn.hwndOwner,eax
;------------------------------- set filter property

		invoke ZeroMemory, addr szFilter, sizeof szFilter
		invoke lstrcpy, addr szFilter, CStr("All files (*.*)")
		invoke lstrlen, addr szFilter
		inc eax
		lea ecx,szFilter
		mov ofn.lpstrFilter,ecx
		add ecx,eax
		invoke lstrcpy,ecx,CStr("*.*")
		mov ofn.nFilterIndex,0

;------------------------------- set custom filter property

		mov eax, pszCustom
		mov ecx, iMaxCustom
		mov ofn.lpstrCustomFilter,eax
		mov ofn.nMaxCustFilter,ecx

;------------------------------- set filename

		mov eax, pszFileName
		mov ecx, iMax
		mov ofn.lpstrFile,eax
		mov ofn.nMaxFile,ecx

		mov eax, pszTitle
		mov ofn.lpstrTitle, eax

if 1										;bug in Windows XP?
		invoke GetDirOnly, addr szPath, pszFileName
		.if (eax)
			lea eax, szPath
			mov ofn.lpstrInitialDir, eax
		.elseif (g_szLastDir)
			mov ofn.lpstrInitialDir, offset g_szLastDir
		.endif
endif
		mov eax, iFlags
		.if (!(eax & 1))
			mov ofn.Flags, OFN_PATHMUSTEXIST
			.if (eax & 2)
				or ofn.Flags, OFN_NOVALIDATE
			.endif
			invoke GetOpenFileName,addr ofn
		.else
			mov ofn.Flags, OFN_PATHMUSTEXIST or OFN_OVERWRITEPROMPT
			invoke GetSaveFileName,addr ofn
		.endif
if 1
		.if (eax)
			invoke GetDirOnly, addr g_szLastDir, pszFileName
		.endif
endif
		ret
		align 4

MyGetFileName endp

;*** assign a COM pointer (with AddRef + Release)

ComPtrAssign proc public uses ebx pp:ptr LPUNKNOWN, lp:LPUNKNOWN
	mov ebx,pp
	.if (dword ptr [ebx])
		invoke vf([ebx],IUnknown,Release)
	.endif
	mov eax,lp
	mov [ebx],eax
	.if (eax)
		invoke vf(eax,IUnknown,AddRef)
	.endif
	ret
	align 4

ComPtrAssign endp

_cexit proc c public
									  ;first do the atexit/_onexit terms
		xor 	eax,eax
		xchg	eax,_onexitbegin
		.while (eax)
			push	[eax+0]
			push	eax
			call	dword ptr [eax+4]
			pop 	eax
			invoke	free,eax
			pop 	eax
		.endw
;;		invoke  _initterm,offset __xp_a,offset __xp_z
;;		invoke  _initterm,offset __xt_a,offset __xt_z
		ret
_cexit endp

atexit proc c public pExitProc:LPVOID

		invoke malloc, 8
		mov ecx, _onexitbegin
		mov edx, pExitProc
		.if (eax)
			mov [eax+0], ecx
			mov [eax+4], edx
			xchg eax, _onexitbegin
		.endif
		ret
atexit endp

IsInterfaceSupported proc public uses ebx esi edi pReqIF:ptr IID, pIFTab:ptr ptr IID, dwEntries:dword, pThis:LPUNKNOWN, ppReturn:ptr LPUNKNOWN
	
	mov ecx,dwEntries
	mov esi,pIFTab
	mov ebx,0
	.while (ecx)
		lodsd
		mov edi,eax
		lodsd
		mov edx,eax
		mov eax,esi
		mov esi,pReqIF
		push ecx
		mov ecx,4
		repz cmpsd
		pop ecx
		.if (ZERO?)
			mov ebx,edx
			add ebx,pThis
			.break
		.endif
		mov esi,eax
		dec ecx
	.endw
	mov ecx,ppReturn
	mov [ecx],ebx

	.if (ebx)
		invoke vf(ebx,IUnknown,AddRef)
		mov eax,S_OK
	.else
		mov eax,E_NOINTERFACE
	.endif
	ret

IsInterfaceSupported endp

;*** translate VARKIND into readable string

GetVarKind proc public varkind:DWORD

		mov ecx,varkind
		.if (ecx == VAR_CONST)
			mov eax,CStr("CONST")
		.elseif (ecx == VAR_STATIC)
			mov eax,CStr("STATIC")
		.elseif (ecx == VAR_PERINSTANCE)
			mov eax,CStr("PERINSTANCE")
		.elseif (ecx == VAR_DISPATCH)
			mov eax,CStr("DISPATCH")
		.else
			mov eax,CStr("???")
		.endif
		ret
		align 4

GetVarKind endp

;*** translate FUNCKIND into readable string

GetFuncKind proc public funckind:DWORD

		mov ecx,funckind
		.if (ecx == FUNC_VIRTUAL)
			mov eax,CStr("virtual")
		.elseif (ecx == FUNC_PUREVIRTUAL)
			mov eax,CStr("pure virtual")
		.elseif (ecx == FUNC_NONVIRTUAL)
			mov eax,CStr("nonvirtual")
		.elseif (ecx == FUNC_STATIC)
			mov eax,CStr("static")
		.elseif (ecx == FUNC_DISPATCH)
			mov eax,CStr("dispatch")
		.else
			mov eax,CStr("???")
		.endif
		ret
		align 4

GetFuncKind endp

;*** translate INVOKEKIND into readable string

GetInvokeKind proc public invkind:DWORD

		mov ecx,invkind
		.if (ecx == INVOKE_FUNC)
			mov eax,CStr("func")
		.elseif (ecx == INVOKE_PROPERTYGET)
			mov eax,CStr("propertyget")
		.elseif (ecx == INVOKE_PROPERTYPUT)
			mov eax,CStr("propertyput")
		.elseif (ecx == INVOKE_PROPERTYPUTREF)
			mov eax,CStr("propertyputref")
		.else
			mov eax,CStr("???")
		.endif
		ret
		align 4

GetInvokeKind endp

	.const

pCallConvValue label dword
	dd CC_FASTCALL
	dd CC_CDECL
	dd CC_MSCPASCAL
	dd CC_MACPASCAL
	dd CC_STDCALL
	dd CC_FPFASTCALL
	dd CC_SYSCALL
	dd CC_MPWCDECL
	dd CC_MPWPASCAL
	dd CC_MAX
NUMCALLCONV equ ($ - pCallConvValue) / sizeof dword

pCallConvStr label ptr ptr
	dd CStr("fastcall")
	dd CStr("cdecl")
	dd CStr("mscpascal")
	dd CStr("macpascal")
	dd CStr("stdcall")
	dd CStr("fpfastcall")
	dd CStr("syscall")
	dd CStr("mpwcdecl")
	dd CStr("mpwpascal")
	dd CStr("max")
	dd CStr("?")

	.code

GetCallConv proc public uses edi callconv:DWORD
		mov eax,callconv
		mov ecx,NUMCALLCONV
		mov edi,offset pCallConvValue
		repnz scasd
		sub ecx,NUMCALLCONV-1
		neg ecx
if 0
		mov eax,callconv
		mov ecx,0
		.while (ecx < NUMCALLCONV)
			.break .if (eax == [ecx*4 + offset pCallConvValue])
			inc ecx
		.endw
endif
		mov eax,[ecx*4+offset pCallConvStr]
		ret
		align 4
GetCallConv endp

;--- return implmentation flags

GetImplTypeFlags_ proc public impltypeflags:DWORD, pStr:LPSTR, iMax:dword

		mov eax,pStr
		mov byte ptr [eax],0
		mov ecx,iMax
		.if (impltypeflags & IMPLTYPEFLAG_FDEFAULT)
			invoke lstrcat, pStr, CStr("default,")
		.endif
		.if (impltypeflags & IMPLTYPEFLAG_FSOURCE)
			invoke lstrcat, pStr, CStr("source,")
		.endif
		.if (impltypeflags & IMPLTYPEFLAG_FRESTRICTED)
			invoke lstrcat, pStr, CStr("restricted,")
		.endif
		.if (impltypeflags & IMPLTYPEFLAG_FDEFAULTVTABLE)
			invoke lstrcat, pStr, CStr("defaultvtable,")
		.endif
		invoke lstrlen, pStr
		.if (eax)
			mov ecx,pStr
			lea ecx,[ecx+eax-1]
			mov byte ptr [ecx],0
		.endif
		ret
		align 4

GetImplTypeFlags_ endp

	.const

dwTypeFlags label dword
	dd TYPEFLAG_FAPPOBJECT
	dd TYPEFLAG_FCANCREATE
	dd TYPEFLAG_FLICENSED
	dd TYPEFLAG_FPREDECLID
	dd TYPEFLAG_FHIDDEN
	dd TYPEFLAG_FCONTROL
	dd TYPEFLAG_FDUAL
	dd TYPEFLAG_FNONEXTENSIBLE
	dd TYPEFLAG_FOLEAUTOMATION
	dd TYPEFLAG_FRESTRICTED
	dd TYPEFLAG_FAGGREGATABLE
	dd TYPEFLAG_FREPLACEABLE
	dd TYPEFLAG_FDISPATCHABLE
	dd TYPEFLAG_FREVERSEBIND
NUMTYPEFLAGS equ ($ - dwTypeFlags) / sizeof dword
pszTypeFlag label dword
	dd CStr("AppObject")
	dd CStr("CanCreate")
	dd CStr("Licensed")
	dd CStr("PredeClId")
	dd CStr("Hidden")
	dd CStr("Control")
	dd CStr("Dual")
	dd CStr("NonExtensible")
	dd CStr("OleAutomation")
	dd CStr("Restricted")
	dd CStr("Aggregatable")
	dd CStr("Replacable")
	dd CStr("Dispatchable")
	dd CStr("ReverseBind")

g_szRestricted       db "restricted",0
g_szSource           db "source",0
g_szBindable         db "bindable",0
g_szRequestEdit      db "requestedit",0
g_szDisplayBind      db "displaybind",0
g_szDefaultBind      db "defaultbind",0
g_szHidden           db "hidden",0
g_szUsesGetLastError db "usesgetlasterror",0
g_szDefaultCollElem  db "defaultcollelem",0
g_szUIDefault        db "uidefault",0
g_szNonBrowsable     db "nonbrowsable",0
g_szReplaceable      db "replaceable",0
g_szImmediateBind    db "immediatebind",0

	align 4

pFuncFlagValue label dword
	dd FUNCFLAG_FRESTRICTED
	dd FUNCFLAG_FSOURCE
	dd FUNCFLAG_FBINDABLE
	dd FUNCFLAG_FREQUESTEDIT
	dd FUNCFLAG_FDISPLAYBIND
	dd FUNCFLAG_FDEFAULTBIND
	dd FUNCFLAG_FHIDDEN
	dd FUNCFLAG_FUSESGETLASTERROR
	dd FUNCFLAG_FDEFAULTCOLLELEM
	dd FUNCFLAG_FUIDEFAULT
	dd FUNCFLAG_FNONBROWSABLE
	dd FUNCFLAG_FREPLACEABLE
	dd FUNCFLAG_FIMMEDIATEBIND
NUMFUNCFLAGVALUE equ ($ - pFuncFlagValue) / sizeof DWORD
	dd g_szRestricted
	dd g_szSource
	dd g_szBindable
	dd g_szRequestEdit
	dd g_szDisplayBind
	dd g_szDefaultBind
	dd g_szHidden
	dd g_szUsesGetLastError
	dd g_szDefaultCollElem
	dd g_szUIDefault
	dd g_szNonBrowsable
	dd g_szReplaceable
	dd g_szImmediateBind

pVarFlagValue label dword
	dd VARFLAG_FREADONLY
	dd VARFLAG_FSOURCE
	dd VARFLAG_FBINDABLE
	dd VARFLAG_FREQUESTEDIT
	dd VARFLAG_FDISPLAYBIND
	dd VARFLAG_FDEFAULTBIND
	dd VARFLAG_FHIDDEN
	dd VARFLAG_FRESTRICTED
	dd VARFLAG_FDEFAULTCOLLELEM
	dd VARFLAG_FUIDEFAULT
	dd VARFLAG_FNONBROWSABLE
	dd VARFLAG_FREPLACEABLE
	dd VARFLAG_FIMMEDIATEBIND
NUMVARFLAGVALUE equ ($ - pVarFlagValue) / sizeof DWORD
	dd CStr("readonly")
	dd g_szSource
	dd g_szBindable
	dd g_szRequestEdit
	dd g_szDisplayBind
	dd g_szDefaultBind
	dd g_szHidden
	dd g_szRestricted
	dd g_szDefaultCollElem
	dd g_szUIDefault
	dd g_szNonBrowsable
	dd g_szReplaceable
	dd g_szImmediateBind

pParamFlagValue label dword
	dd PARAMFLAG_NONE
	dd PARAMFLAG_FIN
	dd PARAMFLAG_FOUT
	dd PARAMFLAG_FLCID
	dd PARAMFLAG_FRETVAL
	dd PARAMFLAG_FOPT
	dd PARAMFLAG_FHASDEFAULT
NUMPARAMFLAGVALUE equ ($ - pParamFlagValue) / sizeof DWORD
	dd CStr("none")
	dd CStr("in")
	dd CStr("out")
	dd CStr("lcid")
	dd CStr("retval")
	dd CStr("optional")
	dd CStr("hasdefault")

	.code

GetTypeFlags proc public uses esi edi ebx dwFlags:dword,pStrOut:LPSTR
	mov esi, offset dwTypeFlags
	mov ecx, NUMTYPEFLAGS
	jmp entryFlags
GetTypeFlags endp

GetVarFlags proc public uses esi edi ebx dwFlags:DWORD, pszOut:LPSTR
	mov esi, offset pVarFlagValue
	mov ecx, NUMVARFLAGVALUE
	jmp entryFlags
GetVarFlags endp

GetParamFlags proc public uses esi edi ebx dwFlags:DWORD, pszOut:LPSTR
	mov esi, offset pParamFlagValue
	mov ecx, NUMPARAMFLAGVALUE
	jmp entryFlags
GetParamFlags endp

GetFuncFlags proc public uses esi edi ebx dwFlags:DWORD, pszOut:LPSTR

	mov esi, offset pFuncFlagValue
	mov ecx, NUMFUNCFLAGVALUE
entryFlags::
	mov ebx, ecx
	dec ebx

	mov edi, pszOut
	.while (ecx)
		lodsd
		test eax,dwFlags
		.if (!ZERO?)
			.if (edi != pszOut)
				mov ax," ,"
				stosw
			.endif
			mov eax,[esi+ebx*4]
			push esi
			mov esi, eax
			lodsb
			.while (al)
				stosb
				lodsb
			.endw
			pop esi
		.endif
		dec ecx		
	.endw	
	mov byte ptr [edi],0
	ret
GetFuncFlags endp

;--- set header bitmap

if ?HDRBMPS
	.data
g_hbmpUpArrow		HBITMAP NULL
g_hbmpDownArrow		HBITMAP	NULL
	.code

SetImageListHdr proc public hWndLV:HWND

		.if (!g_hbmpUpArrow)
if ?USEBMP
			invoke LoadImage, g_hInstance, IDB_BITMAP2, IMAGE_BITMAP, 0, 0, LR_LOADMAP3DCOLORS
			mov g_hbmpUpArrow, eax
			invoke LoadImage, g_hInstance, IDB_BITMAP3, IMAGE_BITMAP, 0, 0, LR_LOADMAP3DCOLORS
			mov g_hbmpDownArrow, eax
else
			invoke ImageList_LoadImage, g_hInstance, IDB_BITMAP1, 0, 0, CLR_DEFAULT, \
				IMAGE_BITMAP, LR_LOADMAP3DCOLORS
			mov g_hbmpUpArrow, eax
endif
		.endif
		invoke ListView_GetHeader( hWndLV)
		invoke Header_SetImageList( eax, g_hbmpUpArrow)
		ret

SetImageListHdr endp

SetHeaderBitmap proc public hWndLV:HWND, iColumn:DWORD, iSortDir:BOOL

local	hWndHdr:HWND
local	hdi:HD_ITEM

		invoke SetImageListHdr, hWndLV

		invoke ListView_GetHeader( hWndLV)
		mov hWndHdr, eax

		mov hdi.mask_, HDI_FORMAT
		invoke Header_GetItem( hWndHdr, iColumn, ADDR hdi)
if ?USEBMP
		mov hdi.imask, HDI_FORMAT or HDI_BITMAP
		or hdi.fmt, HDF_BITMAP or HDF_BITMAP_ON_RIGHT
		.if (!iSortDir)
			mov ecx, g_hbmpUpArrow
		.else
			mov ecx, g_hbmpDownArrow
		.endif
		mov hdi.hbm, ecx
else
		mov hdi.mask_, HDI_FORMAT or HDI_IMAGE
		or hdi.fmt, HDF_IMAGE or HDF_BITMAP_ON_RIGHT
		mov eax, iSortDir
		mov hdi.iImage, eax
endif
		invoke Header_SetItem( hWndHdr, iColumn, ADDR hdi)
		ret

SetHeaderBitmap endp

;--- reset header bitmap

ResetHeaderBitmap proc public uses esi hWndLV:HWND

local	hWndHdr:HWND
local	dwItems:DWORD
local	hdi:HD_ITEM

		invoke ListView_GetHeader( hWndLV)
		mov hWndHdr, eax
		invoke Header_GetItemCount( eax)
		mov dwItems, eax
		xor esi, esi
		mov hdi.mask_, HDI_FORMAT
		.while (esi < dwItems)
			invoke Header_GetItem( hWndHdr, esi, ADDR hdi)
if ?USEBMP
			.if (hdi.fmt & HDF_BITMAP)
				and hdi.fmt, NOT HDF_BITMAP
else
			.if (hdi.fmt & HDF_IMAGE)
				and hdi.fmt, NOT HDF_IMAGE
endif
				invoke Header_SetItem( hWndHdr, esi, ADDR hdi)
			.endif
			inc esi
		.endw
		ret

ResetHeaderBitmap endp

endif

ExcepInfo	struct
pEIP		LPVOID ?
pszCaption	LPSTR ?
dwType		DWORD ?
ExcepInfo	ends

excdlgproc proc uses ebx hWnd:HWND, msg:DWORD, wParam:WPARAM, lParam:LPARAM

local	szText[256]:byte

		mov eax, msg
		.if (eax == WM_INITDIALOG)

			mov ebx, lParam
			mov eax, [ebx].ExcepInfo.pEIP
			mov eax, [eax].EXCEPTION_POINTERS.ExceptionRecord
			invoke printf@CLogWindow, CStr("%s: Exception 0x%08X occured at 0x%08X",10),
				[ebx].ExcepInfo.pszCaption, [eax].EXCEPTION_RECORD.ExceptionCode, [eax].EXCEPTION_RECORD.ExceptionAddress
			invoke wsprintf, addr szText, CStr("Exception 0x%08X occured at 0x%08X.",10,"Function: %s"),
				[eax].EXCEPTION_RECORD.ExceptionCode, [eax].EXCEPTION_RECORD.ExceptionAddress, [ebx].ExcepInfo.pszCaption
			invoke GetDlgItem, hWnd, IDC_STATIC1
			lea ecx, szText
			invoke SetWindowText, eax, ecx
			.if ([ebx].ExcepInfo.dwType == EXCEPTION_EXECUTE_HANDLER)
				invoke CheckRadioButton, hWnd, IDC_RADIO1, IDC_RADIO3, IDC_RADIO3
			.endif
			mov eax, 1

		.elseif (eax == WM_CLOSE)

			invoke IsDlgButtonChecked, hWnd, IDC_RADIO1
			.if (!eax)
				invoke IsDlgButtonChecked, hWnd, IDC_RADIO2
				.if (eax)
					mov eax, EXCEPTION_CONTINUE_SEARCH
				.else
					mov eax, EXCEPTION_EXECUTE_HANDLER
				.endif
			.else
				mov eax, EXCEPTION_CONTINUE_EXECUTION
			.endif
			invoke EndDialog, hWnd, eax

		.elseif (eax == WM_COMMAND)

			movzx eax, word ptr wParam
			.if (eax == IDCANCEL)
				invoke PostMessage, hWnd, WM_CLOSE, 0, 0
			.endif

		.else
			xor eax, eax
		.endif
		ret
excdlgproc endp

DisplayExceptionInfo proc public hWnd:HWND, pEIP:ptr EXCEPTION_INFO_PTRS, pszCaption:LPSTR, dwType:DWORD

		invoke DialogBoxParam, g_hInstance, IDD_EXCEPTIONDLG, hWnd, excdlgproc, addr pEIP
		ret
DisplayExceptionInfo endp

;--- display a string resource in statusbar simple mode

	.const

DefStr	macro id, string
local sym
	dw id
CONST$2 segment
    sym db string,0
CONST$2 ends
CONST$3 segment
	dd offset sym
CONST$3 ends
	endm

CONST$3 segment dword public 'CONST'
StringPointers label dword
CONST$3	ends

StringIDs label word
DefStr	IDM_CHECKTYPELIB		,"scans all items for valid TypeLib link. Items with invalid links will be marked."
DefStr	IDM_VIEW				,"view item in more detail"
DefStr	IDM_LOADTYPELIB 		,"Load type library from a file"
DefStr	IDM_OBJECT				,"shows objects created with CoCreateInstance and not released yet"
DefStr	IDM_OBJECTDLG			,"scans object for supported interfaces and shows them in a dialog"
DefStr	IDM_RENAME				,"rename current item"
DefStr	IDM_DELETE				,"delete current item"
DefStr	IDM_CREATE				,"create an instance with ITypeInfo::CreateInstance"
DefStr	IDM_HKCR				,"display all entries in HKEY_CLASSES_ROOT"
DefStr	IDM_FIND				,"Search text in current list"
DefStr	IDM_CLSID				,"show all entries in HKEY_CLASSES_ROOT\\CLSID"
DefStr	IDM_TYPELIB 			,"show all entries in HKEY_CLASSES_ROOT\\Typelib"
DefStr	IDM_INTERFACE			,"show all entries in HKEY_CLASSES_ROOT\\Interface"
DefStr	IDM_EXIT				,"exit COMView"
DefStr	IDM_CHECKFILE			,"scans all items for file references. If file cannot be opened item will be marked."
DefStr	IDM_REFRESH 			,"updates display"
DefStr	IDM_CREATEINSTANCE		,"calls CoCreateInstance with selected CLSID"
DefStr	IDM_EDIT				,"call internal registry editor for selected items"
DefStr	IDM_APPID				,"show all entries in HKEY_CLASSES_ROOT\\AppId"
DefStr	IDM_TYPELIBDLG			,"opens dialog to view type library information"
DefStr	IDM_OPTIONS 			,"view and set some parameters"
DefStr	IDM_COPY				,"copy selected lines to clipboard"
DefStr	IDM_SAVEAS				,"save listview content in a text file"
DefStr	IDM_ABOUT				,"shows very important information about COMView itself"
DefStr	IDM_COMPCAT 			,"show all entries in HKEY_CLASSES_ROOT\\Component Categories"
DefStr	IDM_CHECKCLSID			,"scans all items for valid CLSID references. Items with invalid links will be marked."
DefStr	IDM_CHECKPROGID 		,"scans all items for valid ProgID links. Items with invalid links will be marked."
DefStr	IDM_COPYGUID			,"copy GUID to clipboard"
DefStr	IDM_CREATEINSTON		,"create object on remote machine"
DefStr	IDM_CHECKAPPID			,"scans all items for valid AppID links. Items with invalid links will be marked."
DefStr	IDM_PROPERTIES			,"open properties dialog for selected file"
DefStr	IDM_USERMODE			,"toggle between user mode and design mode"
DefStr	IDM_HIDE				,"call IOleObject::DoVerb(OLEIVERB_HIDE)"
DefStr	IDM_PRIMARY 			,"call IOleObject::DoVerb(OLEIVERB_PRIMARY)"
DefStr	IDM_SHOW				,"call IOleObject::DoVerb(OLEIVERB_SHOW)"
DefStr	IDM_OPEN				,"call IOleObject::DoVerb(OLEIVERB_OPEN)"
DefStr	IDM_SAVESTREAM			,"save object in a temporary stream created with CreateStreamOnHGlobal()"
DefStr	IDM_LOADSTREAM			,"load previously saved object from stream"
DefStr	IDM_PROPERTIESDLG		,"show/edit properties of object"
DefStr	IDM_TYPEINFODLG 		,"open a type information dialog"
DefStr	IDM_LOGWINDOW		 	,"show logs written by COMView"
DefStr	IDM_UNLOCK				,"Unlocks the object. If object isn't used otherwise, it may be destroyed by this command"
DefStr	IDM_EDITITEM			,"edit property or execute method"
DefStr	IDM_TYPEINFO			,"show typeinfo dialog"
DefStr	IDM_HELP				,"displays help"
DefStr	IDM_SHOWALL 			,"add restricted members to list"
DefStr	IDM_CLOSEOBJECT 		,"call IOleObject::Close"
DefStr	IDM_INPLACEACTIVATE 	,"call IOleObject::DoVerb(OLEIVERB_INPLACEACTIVATE)"
DefStr	IDM_UIACTIVATE			,"call IOleObject::DoVerb(OLEIVERB_UIACTIVATE)"
DefStr	IDM_EXPLORE 			,"open explorer window at file location"
DefStr	IDM_FORCETYPEINFO		,"select typeinfo from type library"
DefStr	IDM_VIEWINTERFACE		,"call interface viewer (external or internal)"
DefStr	IDM_VIEWOBJECT			,"open a view control dialog"
DefStr	IDM_ADVISE				,"call IOleObject::Advise/Unadvise"
DefStr	IDM_INPLACEDEACTIVATE	,"call IOleInPlaceObject::InPlaceDeactivate"
DefStr	IDM_SHDESKTOPFOLDER 	,"view shell's namespace"
DefStr	IDM_VERBPROPERTIES		,"call IOleObject::DoVerb(OLEIVERB_PROPERTIES)"
DefStr	IDM_UNREGISTER			,"For dlls exported function DllUnregisterServer is called. Executables are called with parameter /UnregServer"
DefStr	IDM_REGISTER			,"For dlls exported function DllRegisterServer is called. Executables are called with parameter /RegServer."
DefStr	IDM_EXITVIEWDLG 		,"exit object container app"
DefStr	IDM_ROT 				,"shows objects currently registered as running on this local machine"
DefStr	IDM_VIEWVTBL			,"display vtable of interface"
DefStr	IDM_REMOVEITEM			,"remove selected lines from current view"
DefStr	IDM_SELECTALL			,"select all items"
DefStr	IDM_INVERT				,"invert current selection"
DefStr	IDM_OLEREG				,"edit general OLE settings"
DefStr	IDM_PASTE				,"Get object from clipboard"
DefStr	IDM_UIDEACTIVATE		,"call IOleInPlaceObject::UIDeactivate"
DefStr	IDM_UIDEAD				,"set ambient property UIDead"
DefStr	IDM_SAVESTORAGE 		,"save object in a (temporary) storage file"
DefStr	IDM_LOADSTORAGE 		,"load object from previously saved storage"
DefStr	IDM_UPDATE				,"call IOleObject::Update method"
DefStr	IDM_GETCLASSFACT		,"creates IClassFactory object with CoGetClassObject"
DefStr	IDM_SAVEFILE			,"save object into a file by IPersistFile::Save"
DefStr	IDM_LOADFILE			,"Create a File Moniker and call IMoniker::BindToObject"
DefStr	IDM_CREATELINK			,"creates a link to a file with OleCreateLinkToFile"
DefStr	IDM_VIEWSTORAGE 		,"open a View Storage dialog to show temporary storage object"
DefStr	IDM_VIEWMONIKER 		,"view moniker in an object dialog"
DefStr	IDM_FILENAME			,"Add filename from IPersistFile::GetCurFile to window caption"
DefStr	IDM_STG2FILE			,"save (temporary) storage object into a file"
DefStr	IDM_OPENSTORAGE 		,"create a storage object by opening a file with StgOpenStorage"
DefStr	IDM_LOADOBJECT			,"create an object and initialize it with IPersistStorage::Load"
DefStr	IDM_UNDO				,"discards last change"
DefStr	IDM_VIEWSTREAM			,"view temporary Stream (created with CreateStreamOnHGlobal)"
DefStr	IDM_CONNECT 			,"connect to a source interface"
DefStr	IDM_DISCONNECT			,"disconnect from a source interface"
DefStr	IDM_SAVEPROPBAG 		,"save object in a (temporary) property bag"
DefStr	IDM_VIEWPROP			,"open a dialog to view a property set"
DefStr	IDM_OPENSTREAM			,"create a temporary stream object from a file"
DefStr	IDM_VIEWSTORAGEDLG		,"open a View Storage dialog to show current storage object"
DefStr	IDM_VIEWADVISE			,"call IViewObject::SetAdvise"
DefStr	IDM_DATAADVISE			,"call IDataObject::DAdvise/DUnadvise"
DefStr	IDM_CHECKUPD			,"checks if a newer version of COMView is available"
DefStr	IDM_USETIINVOKE 		,"use ITypeInfo::Invoke instead of IDispatch::Invoke"
DefStr	IDM_COPYOBJECT			,"copies object to clipboard"
DefStr	IDM_SECURITY			,"query security settings from IClientSecurity"
DefStr	IDM_CREATEPROXY 		,"create an interface proxy and show it in an object dialog"
DefStr	IDM_FINDALL 			,"Search all occurences of a text in current list. Remove non-matching lines from view. "
DefStr	IDS_GETMOPS 			,"display information from ITypeInfo::GetMops"
DefStr	IDS_GETHELPCONTEXT		,"display help context string from ITypeInfo::GetDocumentation"
DefStr	IDS_GETHELPFILE 		,"display full path of help file from ITypeInfo::GetDocumentation"
DefStr	IDS_GETSIZEINST 		,"display field cbSizeInstance from TypeAttr"
DefStr	IDS_GETSIZEVFT			,"display field cbSizeVft from TypeAttr"
DefStr	IDS_GETALIGNMENT		,"display cbAlignment from TypeAttr"
DefStr	IDS_GETLCID 			,"display lcid from TypeAttr"
DefStr	IDS_CONSTRUCTOR 		,"display memid of constructor (-1 == none)"
DefStr	IDS_DESTRUCTOR			,"display memid of destructor (-1 == none)"
DefStr	IDS_IDLFLAGS			,"display IDLFlags from TypeAttr"
DefStr	IDM_CONTEXTHELP 		,"display context sensitive help from TypeInfo"
DefStr	IDM_GETCLASS			,"call CoLoadLibrary, call DllGetClassObject for this coclass, then call IClassFactory::CreateInstance"
DefStr	IDM_AMBIENTPROP 		,"Show and edit ambient properties of control container"
DefStr	IDM_LOADFROMFILE		,"Reinitialize Object with IPersistFile::Load() - may work or not"
DefStr	IDM_HEXADECIMAL			,"display values of properties in hexadecimal (VB-like style)"
DefStr	IDM_COPYVALUE			,"copy value of property to the clipboard"
STRINGS equ ($ - offset StringIDs) / sizeof WORD

	.code

DisplayStatusBarString proc public uses edi hWnd:HWND, dwID:DWORD

if 0
local	szText[256]:BYTE

		lea eax, szText
		mov byte ptr [eax],0
		.if (dwID)
			invoke LoadString, g_hInstance, dwID, eax, sizeof szText
		.endif
		lea edx, szText
		.if (!eax)
			mov edx, g_pszMenuHelp
		.endif
		StatusBar_SetText hWnd, 255 or SBT_NOBORDERS, edx
endif
		mov eax, dwID
		mov edi, offset StringIDs
		mov ecx, STRINGS
		repnz scasw
		.if (ZERO?)
			sub edi, offset StringIDs + sizeof WORD
			mov eax, [edi*2 + offset StringPointers]
		.else
			.if (eax)
				mov eax, g_pszMenuHelp
			.endif
			.if (!eax)
				mov eax, offset g_szNull
			.endif
		.endif
		StatusBar_SetText hWnd, 255 or SBT_NOBORDERS, eax
		ret

DisplayStatusBarString endp

SaveNormalWindowPos proc public hWnd:HWND, pRect:ptr RECT

local	wp:WINDOWPLACEMENT

		mov wp.length_, sizeof WINDOWPLACEMENT
		invoke GetWindowPlacement, hWnd, addr wp
		invoke CopyRect, pRect, addr wp.rcNormalPosition
		mov ecx, pRect
		mov eax, [ecx].RECT.left
		sub [ecx].RECT.right, eax
		mov eax, [ecx].RECT.top
		sub [ecx].RECT.bottom, eax
		ret

SaveNormalWindowPos endp


MySetWindowPos proc public uses esi hWnd:HWND, pRect:ptr RECT

		invoke GetSystemMetrics, SM_CXFULLSCREEN
		push eax
		invoke GetSystemMetrics, SM_CYFULLSCREEN
		pop ecx
		mov esi, pRect
		mov edx, [esi].RECT.left
		add edx, [esi].RECT.right
		sub edx, ecx
		jc @F
		sub [esi].RECT.left, edx
@@:
		mov edx, [esi].RECT.top
		add edx, [esi].RECT.bottom
		sub edx, eax
		jc @F
		sub [esi].RECT.top, edx
@@:
		.if ([esi].RECT.right)
			invoke SetWindowPos, hWnd, NULL, [esi].RECT.left, [esi].RECT.top,\
				[esi].RECT.right, [esi].RECT.bottom, SWP_NOZORDER
		.else
			invoke SetWindowPos, hWnd, NULL, [esi].RECT.left, [esi].RECT.top,\
				0, 0, SWP_NOSIZE or SWP_NOZORDER
		.endif
		ret
MySetWindowPos endp

		end
