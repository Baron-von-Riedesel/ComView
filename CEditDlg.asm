
;*** a registry editor dialog class ***
;*** values + keys can be inserted, renamed + deleted *** 

	.386
	.MODEL FLAT, STDCALL
	option casemap:none
	option proc:private

	include COMView.inc
	include richedit.inc

	include classes.inc
	include debugout.inc
INSIDE_CEditDlg equ 1
	include CEditDlg.inc
	include CHexEdit.inc
	include CSplittButton.inc

;*** resource ids

	include rsrc.inc

ifndef TVINSERTSTRUCT
TVINSERTSTRUCT typedef TV_INSERTSTRUCT
endif

Init@CEditDlg		proto
CreateImageLists	proto

if 0
IDD_EDITREGDLG 	EQU 108t

IDI_BINARY		EQU 121t
IDI_FOLDER		EQU 122t
IDI_FOLDEROPEN	EQU 123t
IDI_STRING		EQU 124t
IDR_MENU3		EQU 125t

IDC_EDIT1		EQU 1001t	;ID of edit control
IDC_TREE1		EQU 1002t	;ID of treeview control

IDM_REFRESH 	EQU 40006t	;internal "refresh" command
IDM_DELETE		EQU 40024t	;delete key or value
IDM_RENAME		EQU 40025t	;rename key or value
IDM_KEY 		EQU 40026t
IDM_STRING		EQU 40027t
IDM_BINARY		EQU 40028t
IDM_DWORD		EQU 40029t
endif

;*** other equates

KEYXOFFS		equ 4	;*** num chars between end of value name and value itself
MAXBYTES		equ 16
MAXBYTES$		textequ <"16">

CMD_EDIT		equ 0
CMD_DELVALUE	equ 1
CMD_DELKEY		equ 2
CMD_RENVALUE	equ 3
CMD_RENKEY		equ 4
CMD_NEWVALUE	equ 5
CMD_NEWKEY		equ 6

IDX_BINARY		equ 0
IDX_STRING		equ 1
IDX_FOLDER		equ 2
IDX_FOLDEROPEN	equ 3

;*** typedefs + structures

HKEY	typedef HANDLE

CEditDlg struct
hWnd			HWND	?
hWndTV			HWND	?
hWndEdit		HWND	?
hWndHE			HWND	?
hWndSplit		HWND	?
hWndParent		HWND	?
hRoot			HANDLE	?
hSelItem		HTREEITEM ?
hTmpItem		HTREEITEM ?	;temporary item for rclick selection
EditWndProc		WNDPROC	?
pKeyPair		LPKEYPAIR ?
pMem			LPVOID	?
dwSize			DWORD	?
dwPos			DWORD	?
dwNumKeyPair	DWORD	?
dwCookie		DWORD	?
dwXPosTV		DWORD	?	;X pos of treeview in dialog
dwYPosTV		DWORD	?	;Y pos of treeview in dialog
dwHeightEdit	DWORD	?	;size of edit area
dwHeightBtns	DWORD	?	;size of lower area	
dwExpandMode	DWORD	?
bChanged		BOOLEAN	?
bConfirmDelete	BOOLEAN	?
bRC				BOOLEAN	?
bModeless		BOOLEAN	?
bHexEdHasFocus	BOOLEAN	?
bMBDisplayed	BOOLEAN	?	;was a message box displayed?
CEditDlg ends

__this	textequ <edi>
_this	textequ <[__this].CEditDlg>
thisarg equ <this@:ptr CEditDlg>

	MEMBER hWnd, hWndTV, hWndEdit, hWndParent, hRoot
	MEMBER hSelItem, hTmpItem, EditWndProc, pKeyPair, dwCookie
	MEMBER dwNumKeyPair, bChanged, bConfirmDelete, bRC, bModeless
	MEMBER dwXPosTV, dwYPosTV, dwHeightEdit, dwHeightBtns
	MEMBER dwExpandMode, hWndHE, hWndSplit
	MEMBER pMem, dwSize, dwPos, bHexEdHasFocus, bMBDisplayed

;*** static methods (dont have a this pointer)

String2Number		proto pStr:LPSTR,pDW:ptr dword,radix:dword
String2DWord		proto pStr:LPSTR,pDW:ptr dword
String22DWords		proto pStr:LPSTR,pDW1:ptr dword,pDW2:ptr dword
String2Binary		proto pStr:LPSTR,pBinary:ptr byte
SetValue			proto pszName:LPSTR,pszValue:LPSTR,dwSize:dword,iType:dword, pszStrOut:LPSTR
ReadAllSubItems 	proto hWnd:HWND, hKey:HANDLE, hParent:HANDLE, pszKey:LPSTR, bExpand:dword, iMaxItems:dword
RecalcSize			proto dwYPos:DWORD

;*** private methods, "this" pointer is omitted

UpdateRegistry proto :HTREEITEM, :dword, :LPSTR
FindKey		proto :HTREEITEM, :LPSTR
OnRefresh	proto
OnNotify	proto pNMHDR:ptr NMHDR
OnCommand	proto wParam:WPARAM, lParam:LPARAM

;--------------------------------------------------------------

	.const

IconTab dd IDI_BINARY, IDI_STRING, IDI_FOLDER, IDI_FOLDEROPEN
NUMICONS textequ %($ - offset IconTab) / sizeof DWORD

BtnTab dd IDC_EXPAND, IDOK, IDCANCEL
NUMBUTTONS textequ %($ - BtnTab) / sizeof DWORD

g_pszStandard	LPSTR CStr("(Standard)")

	.data

g_himlLarge 	HANDLE 0
g_himlSmall 	HANDLE 0
g_hMenuEdit		HMENU 0
g_pObject		LPVOID 0
g_rect			RECT <>
g_dwHeightEdit	DWORD 0
	.code

if 0		;deactivated to avoid duplicate definition

String2Number proc uses esi ebx pStr:LPSTR,pDW:ptr dword,radix:dword

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



String2DWord proc uses esi ebx pStr:LPSTR,pDW:ptr dword
	mov esi,pStr
	mov ax,[esi]
	mov ecx,10
	or ah,20h
	cmp ax,'x0'
	jnz @F
	mov cl,16
	inc esi
	inc esi
@@:
	invoke String2Number,esi,pDW,ecx
	ret
	align 4

String2DWord endp


String22DWords proc uses esi edi pStr:LPSTR,pDW1:ptr dword,pDW2:ptr dword

local	szStr[64]:byte

	mov esi,pStr
	lea edi,szStr
	.while (1)
		lodsb
		stosb
		.break .if (al == 0)
		.if (al == '.')
			invoke String2Number,esi,pDW2,10
			mov byte ptr [edi-1],0
			.break
		.endif
	.endw
	invoke String2Number,addr szStr,pDW1,10
	ret
	align 4

String22DWords endp

endif

String2Binary proc uses ebx esi edi pStr:LPSTR,pBinary:ptr byte

local	dwTmp:dword

	mov esi,pStr
	mov edi,pBinary
	mov ebx,0
	cmp byte ptr [esi],0
	jz	exit
next:
	invoke String2Number,esi,addr dwTmp,16
	and eax,eax
	jz error
	mov esi,eax
	mov eax,dwTmp
	test eax,0ffffff00h
	jnz error
	stosb
	inc ebx
	cmp bl,9
	jnc error
@@:
	lodsb
	cmp al,00
	jz exit
	cmp al,' '
	jz @B
	dec esi
	jmp next
exit:
	mov eax,ebx
	ret
error:
	mov eax,-1
	ret
	align 4

String2Binary endp


;*** get a line with format "name = value" 

ifdef @StackBase
	option stackbase:ebp
endif

SetValue proc uses ebx esi edi pszName:LPSTR, pszValue:LPSTR, dwSize:dword, iType:dword, pszStrOut:LPSTR

local	dwESP:DWORD
local	szValue[MAXBYTES * 3 + 8]:byte
local	szFStr[MAXBYTES * 5 + 8]:byte

	mov eax,pszName 			;Name of Value
	.if (eax == NULL)
		mov eax, g_pszStandard
		mov pszName,eax
	.else
		mov al,[eax]
		.if (al == 0)
			mov eax, g_pszStandard
			mov pszName,eax
		.endif
	.endif

	.if ((iType == REG_DWORD) || (iType == REG_DWORD_LITTLE_ENDIAN))
		mov eax,pszValue
		invoke wsprintf, addr szValue, CStr("0x%X"), dword ptr [eax]
		lea eax, szValue
		mov pszValue, eax
		invoke lstrcpy, addr szFStr, CStr("%s  = %s")
	.elseif ((iType == REG_SZ) || (iType == REG_EXPAND_SZ) || (iType == REG_MULTI_SZ))
		.if (iType == REG_MULTI_SZ)
			mov ecx, pszValue
			.while (1)
				mov al,[ecx]
				.if (al == 0)
					mov byte ptr [ecx],13
					.break .if (byte ptr [ecx+1] == 0)
				.endif
				inc ecx
			.endw
		.endif
		invoke lstrcpy, addr szFStr, CStr("%s = ""%s""")
;;	.elseif (iType == REG_BINARY && dwSize <= MAXBYTES)
	.elseif (iType == REG_BINARY)
		push edi
		mov dwESP, esp
		mov ecx,dwSize
		.if (ecx > MAXBYTES)
			mov ecx, MAXBYTES
		.endif
		mov esi,pszValue
		add esi,ecx
		lea edi,szFStr
		.while (ecx > 0)
			movzx eax,byte ptr [esi-1]
			push eax
			dec esi
			mov eax,"X20%"
			stosd
			.if (ecx > 1)
				mov al," "
				stosb
			.endif
			dec ecx
		.endw
		.if (dwSize > MAXBYTES)
			mov eax, "... "
			stosd
		.endif
		mov [edi],cl
		invoke wsprintf, addr szValue, addr szFStr
		mov esp, dwESP
		pop edi

		invoke lstrcpy, addr szFStr, CStr("%s  = %s")
		lea eax, szValue
		mov pszValue, eax
	.else
		invoke lstrcpy, addr szFStr, CStr("%s  = ?")
		mov eax,pszValue
		mov byte ptr [eax],0
	.endif

	invoke wsprintf, pszStrOut, addr szFStr, pszName, pszValue
	invoke lstrlen, pszValue
	push eax
	invoke lstrlen, pszName
	pop ecx
	add eax,KEYXOFFS	;get pos of value
	shl eax,16
	mov ax,cx			;PARAM = LOWORD=size, HIWORD=position
	ret
	align 4

SetValue endp

ifdef @StackBase
	option stackbase:esp
endif

;*** read registry entries and insert them in treeview control
;*** this proc calls recursive


ReadAllSubItems proc uses ebx esi edi hWnd:HWND,hKey:HANDLE, hParent:HANDLE, pszKey:LPSTR, bExpand:dword, iMaxItems:dword

local	hSubKey:HANDLE
local	dwSize:dword
local	hTreeItem:HANDLE
local	iType:dword
local	iMax:dword
local	iKeyLen:dword
local	pszSubKey:LPSTR
local	pszName:LPSTR
local	pszValue:LPSTR
local	pszOutText:LPSTR
local	dwMaxKeySize:DWORD
local	dwMaxNameSize:DWORD
local	dwMaxValueSize:DWORD
local	dwNumKeys:DWORD
local	dwNumValues:DWORD
local	tvi:TVINSERTSTRUCT

;    DebugOut "CEditDlg, ReadAllSubItems(%X, %s, %u)",hParent, pszKey, bExpand

	invoke lstrlen,pszKey
	mov iKeyLen,eax

	mov tvi.hInsertAfter,TVI_LAST
	mov tvi.item.mask_,TVIF_TEXT or TVIF_PARAM or TVIF_IMAGE or TVIF_SELECTEDIMAGE

	mov eax,hParent
	mov tvi.hParent,eax

;----------------------------------- insert the key entry

	mov tvi.item.lParam,0
	.if (iKeyLen)
		mov eax,pszKey
	.else
		mov eax,CStr("HKEY_CLASSES_ROOT")
	.endif
	mov tvi.item.pszText,eax
	mov tvi.item.iImage,IDX_FOLDER
	mov tvi.item.iSelectedImage,IDX_FOLDER
	invoke TreeView_InsertItem( hWnd,addr tvi)
	mov hTreeItem,eax
	mov tvi.hParent,eax

	invoke RegQueryInfoKey, hKey, NULL, NULL, NULL, addr dwNumKeys, addr dwMaxKeySize, NULL,\
			addr dwNumValues, addr dwMaxNameSize, addr dwMaxValueSize, NULL, NULL
	.if (eax == ERROR_SUCCESS)
		mov edx, dwMaxKeySize
		inc edx
		mov ecx, dwMaxNameSize
		inc ecx
		mov eax, dwMaxValueSize
		inc eax
	.else
		mov edx, 384
		mov ecx, 384
		mov eax, 384
	.endif
	mov dwMaxKeySize, edx
	mov dwMaxNameSize, ecx
	mov dwMaxValueSize, eax
	invoke malloc, dwMaxNameSize
	mov pszName, eax
	invoke malloc, dwMaxValueSize
	mov pszValue, eax
	mov eax, dwMaxNameSize
	mov ecx, dwMaxValueSize
	.if (ecx < MAXBYTES * 3 + 4)
		mov ecx, MAXBYTES * 3 + 4
	.endif
	add eax, ecx
	add eax, 16
	invoke malloc, eax
	mov pszOutText, eax
	mov tvi.item.pszText,eax
	xor ebx,ebx
	.if ((!eax) || (!pszName) || (!pszValue))
		invoke MessageBox,hWnd,CStr("Memory failure"),0,MB_OK
		dec ebx
	.endif

	.while (1)
		mov ecx, dwMaxNameSize
		mov iMax, ecx
		mov eax, dwMaxValueSize
		mov dwSize, eax
		invoke RegEnumValue, hKey, ebx, pszName, addr iMax, NULL,\
				addr iType, pszValue, addr dwSize
		.break .if (eax != ERROR_SUCCESS)
if 0; def _DEBUG
		.if ((iType == REG_SZ) || (iType == REG_EXPAND_SZ) || (iType == REG_MULTI_SZ))
			DebugOut "ValueName=%s, Value=%s", addr szName, addr szValue
		.else
			DebugOut "ValueName=%s", addr szName
		.endif
endif
		invoke SetValue, pszName, pszValue, dwSize, iType, pszOutText
			
		mov tvi.item.lParam,eax
		.if ((iType == REG_SZ) || (iType == REG_EXPAND_SZ) || (iType == REG_MULTI_SZ))
			mov tvi.item.iImage,IDX_STRING
			mov tvi.item.iSelectedImage,IDX_STRING
		.else
			mov tvi.item.iImage,IDX_BINARY
			mov tvi.item.iSelectedImage,IDX_BINARY
		.endif
		invoke TreeView_InsertItem( hWnd,addr tvi)
		inc ebx
	.endw
	invoke free, pszName
	invoke free, pszValue
	invoke free, pszOutText

	invoke malloc, dwMaxKeySize
	mov pszSubKey, eax

	xor ebx,ebx
	.while (iMaxItems)
		mov eax, dwMaxKeySize
		mov dwSize, eax
		invoke RegEnumKeyEx, hKey, ebx, pszSubKey, addr dwSize, 0, NULL, 0, NULL
		.break .if (eax != ERROR_SUCCESS)

		invoke RegOpenKeyEx, hKey, pszSubKey, 0, KEY_READ, addr hSubKey
		.if (eax == ERROR_SUCCESS)
			invoke ReadAllSubItems, hWnd, hSubKey, hTreeItem, pszSubKey, bExpand, -1
			invoke RegCloseKey, hSubKey
		.endif
		inc ebx
		dec iMaxItems
	.endw
	.if (bExpand == TRUE)
		invoke TreeView_Expand( hWnd, hTreeItem, TVE_EXPAND)
	.endif

	invoke free, pszSubKey

	mov eax,hTreeItem
	ret
	align 4

ReadAllSubItems endp


;--- find a key in treeview 
;--- this will avoid multiple CLSID root keys for example


FindKey proc uses ebx hItem:HTREEITEM, pszKey:LPSTR

local tvi:TVITEM
local szText[128]:byte

	lea eax,szText
	mov tvi.pszText,eax
	mov tvi.cchTextMax,sizeof szText
	mov tvi.mask_,TVIF_TEXT or TVIF_PARAM		;init constant values for GetItem

	mov eax, hItem
	.if (!eax)
		invoke TreeView_GetRoot( m_hWndTV)
	.else
		invoke TreeView_GetChild( m_hWndTV, eax)
	.endif
	.if (!eax)
		ret
	.endif

	.while (1)
		mov ebx, eax
		.break .if (!ebx)
		mov tvi.hItem,ebx
		invoke TreeView_GetItem( m_hWndTV,addr tvi)		;get partial key
		invoke lstrcmpi, addr szText, pszKey
		.if (!eax)
			return ebx
		.endif
		invoke TreeView_GetNextSibling( m_hWndTV, ebx)
	.endw
	return 0

FindKey endp


;*** this function doesnt exist, so we have to emulate it


RegRenameKey proc uses ebx esi hOldKey:HKEY, pszOldKey:LPSTR, hNewKey:HKEY, pszNewKey:LPSTR

local	szKey[MAX_PATH]:byte
local	szName[MAX_PATH]:byte
local	hSubOldKey:HKEY
local	hSubNewKey:HKEY
local	filetime:FILETIME
local	iType:dword
local	iMax:dword
local	dwSize:dword
local	dwDis:dword
local	hr:DWORD

		invoke RegOpenKeyEx,hOldKey,pszOldKey,NULL,KEY_ALL_ACCESS,addr hSubOldKey
		mov hr,eax
		.if (eax == ERROR_SUCCESS)
			invoke RegCreateKeyEx,hNewKey,pszNewKey,NULL,0,\
				REG_OPTION_NON_VOLATILE, KEY_WRITE,NULL, addr hSubNewKey,addr dwDis
			mov hr,eax
			.if (eax == ERROR_SUCCESS)
				.if (dwDis != REG_CREATED_NEW_KEY)
					mov hr,ERROR_ALREADY_EXISTS
				.endif
				xor ebx, ebx
				.while (hr == ERROR_SUCCESS)
					mov dwSize,sizeof szKey
					invoke RegEnumKeyEx,hSubOldKey,ebx,addr szKey,addr dwSize,NULL,NULL,NULL,addr filetime
					.break .if (eax != ERROR_SUCCESS)
					invoke RegRenameKey,hSubOldKey,addr szKey, hSubNewKey, addr szKey
					mov hr,eax
;----------------------------- dont increment ebx, since subkey has been deleted
;;					inc ebx
				.endw
				xor ebx, ebx
				.while (hr == ERROR_SUCCESS)
					mov iMax,sizeof szName
					mov dwSize,0
					invoke RegEnumValue,hSubOldKey,ebx,addr szName,addr iMax,NULL,\
							addr iType,0,addr dwSize
					.break .if (eax != ERROR_SUCCESS)
					invoke malloc, dwSize
					.if (eax)
						mov esi,eax
						invoke RegQueryValueEx, hSubOldKey, addr szName, NULL,\
								addr iType, esi, addr dwSize
						.if (eax == ERROR_SUCCESS)
							invoke RegSetValueEx, hSubNewKey, addr szName, NULL,\
								iType, esi, dwSize
						.endif
						invoke free,esi
					.else
						mov hr,ERROR_OUTOFMEMORY
						.break
					.endif
					inc ebx
				.endw
				invoke RegCloseKey,hSubNewKey
			.else
				invoke MessageBox, m_hWnd, CStr("RegCreateKeyEx failed"), 0, MB_OK
			.endif
			invoke RegCloseKey,hSubOldKey
			.if (hr == ERROR_SUCCESS)
				invoke RegDeleteKey,hOldKey,pszOldKey		;and delete subkey
			.endif
		.else
			invoke MessageBox, m_hWnd, CStr("RegOpenKeyEx failed"), 0, MB_OK
		.endif
		return hr
		align 4

RegRenameKey endp

;*** HKEY_CLASSES_ROOT (=""), CLSID, Interface + TypeLib keys cannot be deleted
;*** return 0 if this is such a key

CheckForMainKeys proc pStr:LPSTR

	mov eax,pStr
	movzx eax,byte ptr [eax]
	.if (eax)
		mov edx, m_pKeyPair
		mov ecx, m_dwNumKeyPair
		.while (ecx)
			.if ([edx].KEYPAIR.pszRoot && [edx].KEYPAIR.pszKey)
				push edx
				push ecx
				invoke lstrcmpi, pStr, [edx].KEYPAIR.pszRoot 
				pop ecx
				pop edx
				.if (!eax)
					ret
				.endif
			.endif
			add edx, sizeof KEYPAIR
			dec ecx
		.endw
		invoke lstrcmpi, pStr, CStr("CLSID")
		.if (eax != 0)
			invoke lstrcmpi, pStr, CStr("INTERFACE")
			.if (eax != 0)
				invoke lstrcmpi, pStr, CStr("TYPELIB")
				.if (eax != 0)
					invoke lstrcmpi, pStr, CStr("APPID")
				.endif
			.endif
		.endif
	.endif
	ret
	align 4

CheckForMainKeys endp

ifdef @StackBase
	option stackbase:ebp
endif

GetFullKey proc uses esi ebx hTreeItem:HANDLE, pszKey:LPSTR, iMax:DWORD

local tvi:TVITEM
local szKey[MAX_PATH]:byte

	xor esi,esi
	.while (1)
		invoke TreeView_GetParent( m_hWndTV, hTreeItem)
		.break .if (eax == 0)
		push hTreeItem							;save them on stack
		mov hTreeItem,eax
		inc esi
	.endw

	mov ebx, pszKey
	mov byte ptr [ebx],0

;-------------------- now we read the list (in reverse order) and
;-------------------- concat all names to get the full name in szFullKey

	lea eax,szKey
	mov tvi.pszText,eax
	mov tvi.mask_,TVIF_TEXT or TVIF_PARAM		;init constant values for GetItem

	.while (hTreeItem)
		mov eax,hTreeItem
		mov tvi.hItem,eax
		mov tvi.cchTextMax,sizeof szKey
		invoke TreeView_GetItem( m_hWndTV,addr tvi)		;get partial key
		mov eax,tvi.lParam						;get "value" part 
		.if (eax)								;is this a "value" entry?
			movzx ecx,ax						;ecx = size of "value"
			shr eax,16							;eax = pos of value
			sub eax,KEYXOFFS
			lea edx,szKey
			mov byte ptr [edx+eax],0
		.else
			.if (byte ptr [ebx])
				invoke lstrcat, ebx, CStr("\")
			.endif
			invoke lstrcat, ebx, addr szKey
		.endif
		.break .if (esi == 0)					;we are done
		pop hTreeItem							;get next item from stack
		dec esi
	.endw
	ret
	align 4

GetFullKey endp

;*** user has edited/deleted/renamed value or key of selected registry entry.
;*** new value is found in IDC_EDIT1
;*** now change/delete/rename values/keys in registry
;*** and update treeview control
;*** on errors display message and return FALSE, else TRUE

UpdateRegistry proc uses ebx esi hTreeItem:HTREEITEM, iCmd:dword, pszNewName:LPSTR

local	tvi:TVITEM
local	tvis:TVINSERTSTRUCT
local	estrm:EDITSTREAM
local	dwSize:dword
local	dwSize2:dword
local	dwDis:dword
local	hKey:HANDLE
local	hSubKey:HANDLE
local	pszValueName:LPSTR
local	pszValue:LPSTR
local	dwValue:DWORD
local	dwValueLength:DWORD
local	iType:dword
local	iType2:dword
local	pszError:LPSTR
local	dwError:DWORD
local	eid:EDITITEMDESC
local	szKey[MAX_PATH]:byte
local	szFullKey[MAX_PATH]:byte

	mov pszError,NULL
	mov dwError, ERROR_SUCCESS

	.if (hTreeItem == 0)
		.if ((iCmd == CMD_NEWKEY) || (iCmd == CMD_NEWVALUE))
			;
		.else
			mov eax, m_hSelItem
			.if (eax == 0)						;this should be "impossible"	
				ret 
			.endif
			mov hTreeItem,eax
		.endif
	.endif

;-------------------- first the fully qualified key is needed
;-------------------- for this we read all parents of our treeview item

	mov esi,0									;now get all parents of selected entry
	.while (1)
		invoke TreeView_GetParent( m_hWndTV, hTreeItem)
		.break .if (eax == 0)
		push hTreeItem							;save them on stack
		mov hTreeItem,eax
		inc esi
	.endw

	mov byte ptr szFullKey,0					;init full qualified key

;-------------------- now we read the list (in reverse order) and
;-------------------- concat all names to get the full name in szFullKey

	lea eax,szKey
	mov tvi.pszText,eax
	mov tvi.mask_,TVIF_TEXT or TVIF_PARAM		;init constant values for GetItem

	mov pszValueName,NULL 						;NULL=no value name
	.while (hTreeItem)
		mov eax,hTreeItem
		mov tvi.hItem,eax
		mov tvi.cchTextMax,sizeof szKey
		invoke TreeView_GetItem( m_hWndTV,addr tvi)		;get partial key
		.if (szFullKey == 0)					;"HKEY_CLASSES_ROOT" is a "virtual" key
			invoke lstrcmpi,addr szKey,CStr("HKEY_CLASSES_ROOT")
			.if (eax == 0)
				mov szKey,0
			.endif
		.endif
		.if (esi == 0)
			.break .if ((iCmd == CMD_DELKEY) || (iCmd == CMD_RENKEY))
		.endif
		mov eax,tvi.lParam						;get "value" part 
		.if (eax)								;is this a "value" entry?
			movzx ecx,ax						;ecx = size of "value"
			shr eax,16							;eax = pos of value
			sub eax,KEYXOFFS
			lea edx,szKey
			mov pszValueName,edx					;start of value
			mov byte ptr [edx+eax],0
		.else
			.if (byte ptr szFullKey)
				invoke lstrcat,addr szFullKey,CStr("\")
			.endif
			invoke lstrcat,addr szFullKey,addr szKey
		.endif
		.break .if (esi == 0)					;we are done
		pop hTreeItem							;get next item from stack
		dec esi
	.endw

;-------------------- the full qualified key is now in szFullKey
;-------------------- (if CMD_DELKEY its without the name of the actual item
;-------------------- so szFullKey holds the full qualified name of the parent)
;-------------------- the name of the item itself is in szKey
;-------------------- pszValueName points to name of a value (or is NULL)

	mov pszValue, NULL
	.if (iCmd == CMD_EDIT)
		invoke IsWindowVisible, m_hWndHE
		.if (eax)
			invoke SendMessage, m_hWndHE, HEM_GETSIZE, 0, 0
			mov m_dwSize, eax
			invoke malloc, m_dwSize
			.if (eax)
				mov pszValue, eax
				mov m_pMem, eax
				mov m_dwPos, 0
				mov estrm.dwCookie, __this
				mov estrm.dwError,0
				mov estrm.pfnCallback,offset streamoutcb
				invoke SendMessage, m_hWndHE, EM_STREAMOUT, SF_TEXT, addr estrm
;;------------------------ do NOT free m_pMem here, ptr is in pszValue!!!
				mov m_pMem, NULL
				mov eax, m_dwSize
				mov dwValueLength, eax
			.else
				ret
			.endif
		.else
		    invoke GetWindowTextLength, m_hWndEdit
			inc eax
			mov dwValueLength, eax
			invoke malloc, eax
			mov pszValue, eax
			invoke GetWindowText, m_hWndEdit, pszValue, dwValueLength
		.endif
	.elseif (iCmd == CMD_RENVALUE)
		invoke lstrcmp, pszValueName, pszNewName
		.if (eax == 0)
			 ret
		.endif
		invoke lstrcmp, pszNewName, g_pszStandard
		.if (eax == 0)
			mov eax,pszNewName
			mov byte ptr [eax],0
		.endif
	.elseif (iCmd == CMD_RENKEY)
		invoke lstrcmp, addr szKey, pszNewName
		.if (eax == 0)
			 ret
		.endif
	.endif

	.if (pszValueName)
		invoke lstrcmp, pszValueName, g_pszStandard
		.if (eax == 0)
			mov eax,pszValueName
			mov byte ptr [eax],0
		.endif
	.endif

	invoke RegOpenKeyEx, m_hRoot,addr szFullKey,0,KEY_SET_VALUE or KEY_QUERY_VALUE,addr hKey
	.if (eax != ERROR_SUCCESS)
		mov pszError,CStr("key not found or no write access")
		mov dwError, eax
	.else

		.if (iCmd == CMD_NEWKEY)				;insert a key?
			mov dwDis,REG_OPENED_EXISTING_KEY
			mov esi,1
			.while (dwDis == REG_OPENED_EXISTING_KEY)
				invoke wsprintf,addr szKey,CStr("New Key #%u"), esi
				invoke RegCreateKeyEx,hKey,addr szKey,NULL,0,REG_OPTION_NON_VOLATILE,\
						KEY_WRITE, NULL, addr hSubKey, addr dwDis
				.if (eax != ERROR_SUCCESS)
					mov pszError,CStr("cannot create key")
					mov dwError, eax
					.break
				.endif
				invoke RegCloseKey, hSubKey
				inc esi
			.endw
		.elseif (iCmd == CMD_NEWVALUE)			;insert a value?

			mov esi,1
			.while (1)
				invoke wsprintf,addr szKey,CStr("New Value #%u"), esi
				@mov dwSize, 0
				invoke RegQueryValueEx, hKey, addr szKey, NULL, addr iType, 0, addr dwSize
				.break .if (eax != ERROR_SUCCESS)
				inc esi
			.endw
			.if (eax == ERROR_FILE_NOT_FOUND)
				mov eax, pszNewName
				mov iType,eax
				.if (eax == REG_DWORD)
					mov ecx,sizeof DWORD
				.elseif (eax == REG_BINARY)
					mov ecx,0
				.else
					mov ecx,sizeof BYTE
				.endif
				mov dwSize,ecx
				mov dwValue,0
				invoke RegSetValueEx, hKey, addr szKey, NULL, iType, addr dwValue, dwSize
			.endif

		.elseif (iCmd == CMD_DELKEY)		;delete a key?

			.if (szFullKey[0] == 0)
				invoke CheckForMainKeys,addr szKey
				.if (eax == 0)
					mov pszError,CStr("Keys CLSID, Interface, Typelib, AppID and",0Ah,"HKEY_CLASSES_ROOT itself cannot be deleted")
				.endif
			.endif
			.if (pszError == NULL)
				invoke DeleteKeyWithSubKeys,hKey,addr szKey
				.if (eax != ERROR_SUCCESS)
					mov pszError,CStr("cannot delete key")
					mov dwError, eax
				.endif
			.endif

		.elseif (iCmd == CMD_RENKEY)    ;rename a key?

			.if (szFullKey[0] == 0)
				invoke CheckForMainKeys,addr szKey
				.if (eax == 0)
					mov pszError,CStr("keys CLSID, Interface, Typelib (and HKEY_CLASSES_ROOT itself) cannot be renamed")
				.endif
			.endif
			.if (pszError == NULL)
				invoke RegRenameKey, hKey, addr szKey, hKey, pszNewName
				.if (eax == ERROR_ALREADY_EXISTS)
					mov pszError,CStr("key already exists")
					mov dwError, eax
				.elseif (eax == ERROR_FILE_NOT_FOUND)
					mov pszError,CStr("key no longer exists")
					mov dwError, eax
				.endif
			.endif

		.elseif (iCmd == CMD_DELVALUE)	;delete a value?

			invoke RegDeleteValue,hKey,pszValueName
			.if (eax != ERROR_SUCCESS)
				mov pszError,CStr("value in registry cannot be deleted")
				mov dwError, eax
			.endif

		.elseif (iCmd == CMD_RENVALUE)  ;rename a value?

			invoke malloc, 1024
			mov pszValue, eax
			mov dwSize, 1024
			mov edx, eax
			invoke RegQueryValueEx, hKey, pszValueName,NULL, addr iType, edx, addr dwSize
			.if (eax != ERROR_SUCCESS)
				mov pszError,CStr("value no longer exists")
				mov dwError, eax
			.else
				invoke RegQueryValueEx, hKey, pszNewName, NULL, addr iType2, NULL, addr dwSize2
;----------------------------------- Standard Value always exists, check this
				mov ecx,pszNewName
				.if ((byte ptr [ecx] == 0) && (dwSize2 == 1))
					inc eax			;set eax to error in this case
				.endif
				.if (eax == ERROR_SUCCESS) 
					mov pszError,CStr("value already exists")
				.else
					mov edx, pszValue
				    invoke RegSetValueEx, hKey, pszNewName, NULL, iType, edx, dwSize
					.if (eax == ERROR_SUCCESS)
						invoke RegDeleteValue, hKey, pszValueName
						.if (eax == ERROR_SUCCESS)
							mov eax,pszNewName
							mov pszValueName,eax
						.endif
					.endif
				.endif
			.endif

		.elseif (iCmd == CMD_EDIT)		;change a value content?

			mov eax, dwValueLength
			mov dwSize,eax
			invoke RegQueryValueEx, hKey, pszValueName, NULL, addr iType, NULL, 0
			.if (eax == ERROR_SUCCESS)
				.if (iType == REG_DWORD)
					mov dwSize,4
					invoke String2DWord, pszValue, pszValue
					.if (eax == 0)
						mov pszError,CStr("no valid data entered")
					.endif
				.elseif (iType == REG_BINARY)
					mov eax, dwValueLength
					mov dwSize, eax
					.if (eax == -1)
						mov pszError,CStr("no valid data (enter 0 to ",MAXBYTES$," byte values)")
					.endif
				.elseif (iType == REG_MULTI_SZ)
					mov ecx, pszValue
					mov edx, ecx
					.while (1)
						mov al,[ecx]
						mov ah, al
						.if (al == 13)
							mov al,00
						.endif
						.if (al != 10)
							mov [edx],al
							inc edx
						.endif
						.break .if (ah == 0)
						inc ecx
					.endw
					sub edx, pszValue
					mov dwSize, edx
				.endif
			.elseif (eax == ERROR_FILE_NOT_FOUND)	
				mov iType,REG_SZ				;there was no previous value (no error)
			.else
				mov pszError,CStr("value not found")	
			.endif
			.if (pszError == NULL)
				invoke RegSetValueEx, hKey, pszValueName, 0, iType, pszValue, dwSize
				.if (eax != ERROR_SUCCESS)
					mov pszError,CStr("value in registry cannot be set")
					mov dwError, eax
				.endif
			.endif
		.endif

		invoke RegCloseKey,hKey
;-------------------------------------------- registry is up to date, now update treeview
		.if (pszError == NULL)
			mov tvi.mask_,TVIF_TEXT or TVIF_PARAM

			.if (iCmd == CMD_NEWKEY)

				mov eax,hTreeItem
				mov tvis.hParent,eax
				.if (eax)
					mov tvis.hInsertAfter,TVI_FIRST
				.else
					mov tvis.hInsertAfter,TVI_LAST
				.endif
				mov tvis.item.mask_,TVIF_TEXT or TVIF_PARAM or TVIF_IMAGE or TVIF_SELECTEDIMAGE
				mov tvis.item.lParam,0
				lea eax,szKey
				mov tvis.item.pszText,eax
				mov tvis.item.iImage,IDX_FOLDER
				mov tvis.item.iSelectedImage,IDX_FOLDER
				invoke TreeView_InsertItem( m_hWndTV, addr tvis)
				invoke TreeView_SelectItem( m_hWndTV, eax)
				invoke SendMessage, m_hWnd, WM_COMMAND, IDM_RENAME, 0

			.elseif (iCmd == CMD_NEWVALUE)

				mov eax,hTreeItem
				mov tvis.hParent,eax
				invoke SetValue, addr szKey, addr dwValue, dwSize, iType, addr szFullKey
				mov tvis.item.lParam,eax
				mov tvis.hInsertAfter,TVI_FIRST
				mov tvis.item.mask_,TVIF_TEXT or TVIF_PARAM or TVIF_IMAGE or TVIF_SELECTEDIMAGE
				lea eax,szFullKey
				mov tvis.item.pszText,eax

				.if ((iType == REG_SZ) || (iType == REG_EXPAND_SZ) || (iType == REG_MULTI_SZ))
					mov tvis.item.iImage,IDX_STRING
					mov tvis.item.iSelectedImage,IDX_STRING
				.else
					mov tvis.item.iImage,IDX_BINARY
					mov tvis.item.iSelectedImage,IDX_BINARY
				.endif
				invoke TreeView_InsertItem( m_hWndTV, addr tvis)
				invoke TreeView_SelectItem( m_hWndTV, eax)
				invoke SendMessage, m_hWnd, WM_COMMAND, IDM_RENAME, 0

			.elseif ((iCmd == CMD_EDIT) || (iCmd == CMD_RENVALUE))

				smalloc edx,1024				;allocate temp buffer on stack
				mov tvi.pszText,edx
				invoke SetValue, pszValueName, pszValue, dwSize, iType, edx
				mov tvi.lParam,eax
				invoke TreeView_SetItem( m_hWndTV,addr tvi)
				sfree							;free temp buffer

			.elseif (iCmd == CMD_RENKEY)

				mov eax,pszNewName
				mov tvi.pszText,eax
				mov tvi.lParam,0
				invoke TreeView_SetItem( m_hWndTV,addr tvi)

			.else		;(iCmd == CMD_DELKEY) || (iCmd == CMD_DELVALUE)

				invoke TreeView_SelectItem( m_hWndTV,NULL)
				invoke TreeView_DeleteItem( m_hWndTV,hTreeItem)
			.endif
		.endif
	.endif

	.if (pszValue)
		invoke free, pszValue
	.endif

	.if (pszError)
		invoke MessageBeep, MB_OK
		.if ( dwError != ERROR_SUCCESS )
			sub esp,128
			mov edx, esp
			invoke wsprintf, edx, CStr("%s [%X]"), pszError, dwError
			mov edx, esp
			invoke MessageBox, m_hWnd, edx, 0, MB_OK
			add esp,128
		.else
			invoke MessageBox, m_hWnd, pszError, 0, MB_OK
		.endif
		xor eax,eax
	.else
		mov eax,m_dwCookie
;------------------------------- if a Cookie has been set, inform parent of change
		.if (eax != -1)
			mov eid.dwCookie,eax
			lea eax,szFullKey
			mov eid.pszKey,eax
			invoke SendMessage, m_hWndParent, WM_COMMAND, IDM_REFRESHLINE, addr eid
		.endif
		mov eax,1
	.endif
	ret
	align 4

UpdateRegistry endp

ifdef @StackBase
	option stackbase:esp
endif


;*** subclass label edit control of treeview (bug in winnt/win95)


editsubclassproc proc uses __this hWnd:HWND,message:dword,wParam:WPARAM,lParam:LPARAM

    DebugOut "editsubclassproc, message=%X, wParam=%X, lParam=%X", message, wParam, lParam

	mov __this,g_pObject

	mov eax,message
	.if (eax == WM_GETDLGCODE)
		mov eax,DLGC_WANTALLKEYS
	.else
		invoke CallWindowProc, m_EditWndProc, hWnd, message, wParam, lParam
	.endif
	ret
	align 4

editsubclassproc endp


RecalcSize proc uses ebx dwYPos:DWORD

local rect:RECT
local rect2:RECT
local dwWidth:DWORD
local dwGripSize:DWORD

	invoke GetWindowRect, m_hWndSplit, addr rect2
	mov eax, rect2.bottom
	sub eax, rect2.top
	mov dwGripSize, eax
	invoke GetWindowRect, m_hWndEdit, addr rect2
	invoke ScreenToClient, m_hWnd, addr rect2.right
	mov eax, rect2.bottom
	sub eax, dwYPos
	sub eax, dwGripSize
	mov m_dwHeightEdit, eax
	mov g_dwHeightEdit, eax

	invoke BeginDeferWindowPos, 4
	mov ebx, eax

	invoke GetClientRect, m_hWnd, addr rect

	mov eax, rect.right
	sub eax, m_dwXPosTV		;subtract left rim
	sub eax, m_dwXPosTV		;subtract right rim
	mov dwWidth, eax

;------------------------------------ set treeview control
	mov eax, rect.bottom
	sub eax, m_dwYPosTV
	sub eax, m_dwHeightBtns
	sub eax, m_dwHeightEdit
	sub eax, dwGripSize
	mov rect.bottom, eax
	invoke DeferWindowPos, ebx, m_hWndTV, NULL, 0, 0, dwWidth, rect.bottom, SWP_NOMOVE or SWP_NOZORDER or SWP_NOACTIVATE

	mov eax, m_dwYPosTV
	add eax, rect.bottom
	invoke DeferWindowPos, ebx, m_hWndSplit, NULL, m_dwXPosTV, eax, dwWidth, dwGripSize, SWP_NOZORDER or SWP_NOACTIVATE

	mov eax, m_dwYPosTV
	add eax, rect.bottom
	add eax, dwGripSize
	mov rect.top, eax
	invoke DeferWindowPos, ebx, m_hWndEdit, NULL, m_dwXPosTV, rect.top, dwWidth, m_dwHeightEdit, SWP_NOZORDER or SWP_NOACTIVATE
	invoke DeferWindowPos, ebx, m_hWndHE, NULL, m_dwXPosTV, rect.top, dwWidth, m_dwHeightEdit, SWP_NOZORDER or SWP_NOACTIVATE
	invoke EndDeferWindowPos, ebx
	ret
	align 4

RecalcSize endp

;--- user resized window (this is a resizeable dialog)

OnSize proc uses ebx esi dwType:dword, dwWidth:dword, dwHeight:dword

local hWndButton:HWND
local dwHeightTV:DWORD
local dwHeightEdit:DWORD
local dwWidthBtn:DWORD
local dwXPos:DWORD
local dwAddX:DWORD
local dwGripSize:DWORD
local rect:RECT

	invoke GetWindowRect, m_hWndSplit, addr rect
	mov eax, rect.bottom
	sub eax, rect.top
	mov dwGripSize, eax

	invoke GetWindowRect, m_hWndTV, addr rect

	mov eax, dwWidth
	sub eax, m_dwXPosTV		;subtract left rim
	sub eax, m_dwXPosTV		;subtract right rim
	mov dwWidth, eax

	mov eax, m_dwHeightEdit
	mov dwHeightEdit, eax

	invoke BeginDeferWindowPos, 4 + NUMBUTTONS
	mov ebx, eax
;------------------------------------ set treeview control
	mov eax, dwHeight
	sub eax, m_dwYPosTV
	sub eax, m_dwHeightBtns
	sub eax, dwGripSize
	sub eax, m_dwHeightEdit
	mov dwHeightTV, eax
	test eax, eax
	.if (SIGN?)
		@mov dwHeightTV, 0
		neg eax
		sub dwHeightEdit, eax
	.endif
	invoke DeferWindowPos, ebx, m_hWndTV, NULL, 0, 0, dwWidth, dwHeightTV, SWP_NOMOVE or SWP_NOZORDER or SWP_NOACTIVATE

	mov ecx, m_dwYPosTV
	add ecx, dwHeightTV
	invoke DeferWindowPos, ebx, m_hWndSplit, NULL, m_dwXPosTV, ecx, dwWidth, dwGripSize, SWP_NOZORDER or SWP_NOACTIVATE

;------------------------------------ set edit control
	mov ecx, m_dwYPosTV
	add ecx, dwHeightTV
	add ecx, dwGripSize
	mov rect.top, ecx

	invoke DeferWindowPos, ebx, m_hWndEdit, NULL, m_dwXPosTV, rect.top, dwWidth, dwHeightEdit, SWP_NOZORDER or SWP_NOACTIVATE
	invoke DeferWindowPos, ebx, m_hWndHE, NULL, m_dwXPosTV, rect.top, dwWidth, dwHeightEdit, SWP_NOZORDER or SWP_NOACTIVATE

;------------------------------------ set the buttons (rearrange XPos)

	mov esi, offset BtnTab

	invoke GetDlgItem, m_hWnd, [esi]
	mov hWndButton, eax
	invoke GetWindowRect, hWndButton, addr rect
	mov eax, rect.right
	sub eax, rect.left
	mov dwWidthBtn, eax

	mov eax, dwWidthBtn
	mov ecx, NUMBUTTONS
	mul ecx
	mov ecx, eax
	mov eax, dwWidth
	sub eax, ecx
	xor edx, edx
	mov ecx, NUMBUTTONS - 1
	div ecx
	mov dwAddX, eax

	invoke ScreenToClient, m_hWnd, addr rect
	mov ecx,dwHeight
	sub ecx, m_dwHeightBtns
	add ecx, m_dwYPosTV
	mov rect.top, ecx

	mov eax, m_dwXPosTV
	mov dwXPos, eax

	mov ecx, NUMBUTTONS
	.while (ecx)
		push ecx
		lodsd
		invoke GetDlgItem, m_hWnd, eax
		invoke DeferWindowPos, ebx, eax, NULL, dwXPos, rect.top, 0, 0, SWP_NOSIZE or SWP_NOZORDER or SWP_NOACTIVATE
		mov eax, dwXPos
		add eax, dwWidthBtn
		add eax, dwAddX
		mov dwXPos, eax
		pop ecx
		dec ecx
	.endw

	invoke EndDeferWindowPos, ebx
	ret
	align 4

OnSize endp


streamincb proc uses __this dwCookie:DWORD, pbBuff:LPBYTE , cb:LONG , pcb:ptr LONG

	mov __this, dwCookie
	mov ecx, m_dwSize
	mov edx, m_dwPos
	sub ecx, edx
	.if (ecx > cb)
		mov ecx, cb
	.endif
	push ecx
	add edx, m_pMem
	invoke CopyMemory, pbBuff, edx, ecx
	mov edx,pcb
	pop ecx
	mov [edx],ecx
	add m_dwPos, ecx
	xor eax, eax
	ret
	align 4

streamincb endp

streamoutcb proc uses __this dwCookie:DWORD, pbBuff:LPBYTE , cb:LONG , pcb:ptr LONG

	mov __this, dwCookie
	mov edx, m_pMem
	add edx, m_dwPos
	invoke CopyMemory, edx, pbBuff, cb
	mov edx,pcb
	mov ecx, cb
	add m_dwPos, ecx
	mov [edx],ecx
	xor eax, eax
	ret
	align 4

streamoutcb endp


;*** edit registry dialog has received WM_NOTIFY


OnNotify proc uses ebx esi pNMHDR:ptr NMHDR

local	hSubMenu:HMENU
local	bDelRenCmd:BOOL
local	hKey:HANDLE
local	pt:POINT
local	dwSize:DWORD
local	dwType:DWORD
local	tvht:TVHITTESTINFO
local	estrm:EDITSTREAM
local	tvi:TVITEM
local	szStr[512]:byte
local	szKey[512]:byte
local	szValue[512]:byte

	mov esi,pNMHDR

	.if ([esi].NMHDR.idFrom == IDC_SPLITBTN)
		.if ([esi].NMHDR.code == SBN_SETSIZE)
			invoke RecalcSize, [esi].SBNOTIFY.iPos
		.endif
		jmp done
	.endif

	assume esi:ptr NMHDR
	mov eax,[esi].code
	.if (eax == NM_RCLICK)

		DebugOut "OnNotify, NM_RCLICK, Code=%d", [esi].NMHDR.code

		.if ([esi].NMHDR.idFrom == IDC_TREE1)
			.if (g_hMenuEdit)
				invoke GetSubMenu, g_hMenuEdit,0
				mov hSubMenu,eax
				invoke GetCursorPos,addr pt
				invoke GetCursorPos,addr tvht.pt
							; get the item below hit point
				invoke ScreenToClient, m_hWndTV,addr tvht.pt
				invoke TreeView_HitTest( m_hWndTV,addr tvht)
				.if (tvht.hItem == NULL)
					mov bDelRenCmd, MF_BYCOMMAND or MF_DISABLED or MF_GRAYED
					mov ecx,MF_BYPOSITION or MF_ENABLED
				.else
					mov bDelRenCmd,MF_BYCOMMAND or MF_ENABLED
					mov tvi.mask_, TVIF_PARAM
					mov eax, tvht.hItem
					mov tvi.hItem, eax
					invoke TreeView_GetItem( m_hWndTV, addr tvi)
					.if (tvi.lParam)						; disable "New" for "Values"
						mov ecx,MF_BYPOSITION or MF_DISABLED or MF_GRAYED
					.else
						mov ecx,MF_BYPOSITION or MF_ENABLED
					.endif
				.endif
				invoke EnableMenuItem, hSubMenu, 0, ecx
				invoke EnableMenuItem, hSubMenu, IDM_RENAME, bDelRenCmd
				invoke EnableMenuItem, hSubMenu, IDM_DELETE, bDelRenCmd

				invoke TrackPopupMenu,hSubMenu,\
					TPM_LEFTALIGN or TPM_LEFTBUTTON or TPM_RETURNCMD,\
					pt.x,pt.y,0, m_hWnd,NULL
				.if (eax)
					mov ecx,tvht.hItem
					mov m_hTmpItem,ecx
					invoke OnCommand, eax, 0
					mov m_hTmpItem, NULL
				.endif
			.endif
		.endif

	.elseif (eax == TVN_SELCHANGING)

;		DebugOut "OnNotify, TVN_SELCHANGING"

		assume esi:ptr NMTREEVIEW
		invoke IsWindowVisible, m_hWndHE
		.if (eax)
			invoke SendMessage, m_hWndHE, EM_GETMODIFY, 0, 0
			.if (eax)
				invoke UpdateRegistry, [esi].itemOld.hItem,CMD_EDIT, 0
				.if (!eax)
					invoke SetWindowLong, m_hWnd,DWL_MSGRESULT,1	;prevent tv selection
				.endif
			.endif
			invoke SendMessage, m_hWndHE, EM_SETMODIFY, 0, 0
		.elseif (m_bChanged == TRUE)
			invoke UpdateRegistry, [esi].itemOld.hItem,CMD_EDIT, 0
			.if (!eax)
				invoke SetWindowLong, m_hWnd,DWL_MSGRESULT,1	;prevent tv selection
			.else
				mov m_bChanged, FALSE
			.endif
		.endif
		 mov eax,1

	.elseif (eax == TVN_SELCHANGED)

;		DebugOut "OnNotify, TVN_SELCHANGED"

		.if ([esi].itemNew.state & TVIS_SELECTED)
			mov eax,[esi].itemNew.hItem
		.else
			xor eax,eax
		.endif
		mov m_hSelItem,eax
		mov m_hTmpItem,0


		invoke ShowWindow, m_hWndEdit, SW_SHOW
		invoke EnableWindow, m_hWndEdit, FALSE
		invoke ShowWindow, m_hWndHE, SW_HIDE

		mov [esi].itemNew.mask_,TVIF_TEXT or TVIF_PARAM ;read treeview item
		lea eax,szStr
		mov [esi].itemNew.pszText,eax
		mov [esi].itemNew.cchTextMax,sizeof szStr
		invoke TreeView_GetItem( [esi].hdr.hwndFrom,addr [esi].itemNew)
		movzx eax,word ptr [esi].itemNew.lParam+2		;eax = pos of value
;---------------------------------------- if its a key, do nothing
		.if (!eax)
			mov eax,1
			invoke SetWindowText, m_hWndEdit, addr g_szNull
			jmp done
		.endif
		lea edx,szStr
		mov byte ptr [edx + eax - KEYXOFFS],0

		invoke GetFullKey, [esi].itemNew.hItem, addr szKey, sizeof szKey


		invoke RegOpenKeyEx, m_hRoot, addr szKey, NULL, KEY_READ, addr hKey
		.if (eax == ERROR_SUCCESS)
			invoke lstrcmp, addr szStr, g_pszStandard
			.if (!eax)
				mov szStr, 0
			.endif
			invoke RegQueryValueEx, hKey, addr szStr, NULL, addr dwType, NULL, addr m_dwSize
			.if (eax == ERROR_SUCCESS)
				.if (dwType == REG_BINARY)
					invoke ShowWindow, m_hWndHE, SW_SHOW
					invoke EnableWindow, m_hWndHE, TRUE
					invoke ShowWindow, m_hWndEdit, SW_HIDE
					invoke malloc, m_dwSize
					.if (eax)
						mov m_pMem, eax
						@mov m_dwPos, 0
						invoke RegQueryValueEx, hKey, addr szStr, NULL, NULL, m_pMem, addr m_dwSize
						.if (eax == ERROR_SUCCESS)
							mov estrm.dwCookie, __this
							mov estrm.dwError,0
							mov estrm.pfnCallback,offset streamincb
							invoke SendMessage, m_hWndHE, EM_STREAMIN, SF_TEXT, addr estrm
						.endif
						invoke free, m_pMem
						mov m_pMem, NULL
					.endif
				.else
					invoke ShowWindow, m_hWndEdit, SW_SHOW
					invoke EnableWindow,m_hWndEdit, TRUE
					invoke ShowWindow, m_hWndHE, SW_HIDE
					movzx ecx, word ptr [esi].itemNew.lParam+0	;ecx = size of "value"
					movzx eax, word ptr [esi].itemNew.lParam+2
					push edi
					push esi
					lea esi, szStr
					add esi, eax
					lea edi,szValue
					.while (ecx)
						lodsb
						stosb
						.if (al == 13)
							mov al, 10
							stosb
						.endif
						dec ecx
					.endw
					mov al,00
					stosb
					pop esi
					pop edi
					invoke SetWindowText, m_hWndEdit, addr szValue	;update edit control
				.endif
			.endif
			invoke RegCloseKey, hKey
		.endif
		mov m_bChanged,FALSE

		mov ecx,FALSE
		.if ([esi].itemNew.state & TVIS_SELECTED)
			.if (![esi].itemNew.lParam)
				 mov ecx,TRUE
			.endif
		.endif
		mov eax,1

	.elseif (eax == TVN_ITEMEXPANDED)

		mov [esi].itemNew.mask_,TVIF_IMAGE or TVIF_SELECTEDIMAGE
		.if ([esi].itemNew.state & TVIS_EXPANDED)
			mov [esi].itemNew.iImage,IDX_FOLDEROPEN
			mov [esi].itemNew.iSelectedImage,IDX_FOLDEROPEN
		.else
			mov [esi].itemNew.iImage,IDX_FOLDER
			mov [esi].itemNew.iSelectedImage,IDX_FOLDER
		.endif
		invoke TreeView_SetItem( [esi].hdr.hwndFrom,addr [esi].itemNew)
		mov eax,1

	.elseif (eax == TVN_KEYDOWN)

		DebugOut "OnNotify, TVN_KEYDOWN"

		assume esi:ptr NMTVKEYDOWN
		.if ([esi].wVKey == VK_DELETE)
			invoke PostMessage,m_hWnd,WM_COMMAND,IDM_DELETE,0
		.endif

	.elseif (eax == TVN_BEGINLABELEDIT)

		DebugOut "TVN_BEGINLABELEDIT"

		invoke TreeView_GetEditControl( m_hWndTV)
		mov ebx, eax
		mov g_pObject,__this
		invoke SetWindowLong, ebx, GWL_WNDPROC, editsubclassproc
		mov m_EditWndProc,eax

		assume esi:ptr NMTVDISPINFO
		.if ([esi].item.lParam == 0)    ;rename key
			mov eax,1
		.else
			invoke lstrcpy,addr szStr,[esi].item.pszText
			mov eax,[esi].item.lParam
			shr eax,16
			sub eax,KEYXOFFS
			lea ecx,szStr
			mov byte ptr [ecx+eax],0
			invoke SetWindowText, ebx, addr szStr
		.endif

	.elseif (eax == TVN_ENDLABELEDIT)

		DebugOut "TVN_ENDLABELEDIT"

		assume esi:ptr NMTVDISPINFO
		.if ([esi].item.pszText != NULL)
			.if ([esi].item.lParam == 0)
				invoke UpdateRegistry, [esi].item.hItem, CMD_RENKEY,[esi].item.pszText
				mov eax,1
			.else
				invoke UpdateRegistry, [esi].item.hItem, CMD_RENVALUE,[esi].item.pszText
				mov eax,1
			.endif
		.endif
if 0
		push eax
		TreeView_GetEditControl m_hWndTV
		invoke SetWindowLong, eax, GWL_WNDPROC, m_EditWndProc
		mov m_EditWndProc,NULL
		pop eax
endif
	.else
		xor eax,eax
	.endif
done:
	ret
	assume esi:nothing
	align 4

OnNotify endp


;*** refresh view


OnRefresh proc uses ebx esi

local	hTreeRoot:HANDLE
local	hKey:HANDLE
local	iMax:dword
local	bExpand:dword
local	hCursorOld:HCURSOR
local	szKey[MAX_PATH]:byte
local	szData[MAX_PATH]:byte

	DebugOut "EditDlg, OnCommand, IDM_REFRESH enter"

	invoke LoadCursor,NULL,IDC_WAIT
	invoke SetCursor, eax
	mov hCursorOld,eax

	invoke TreeView_SelectItem( m_hWndTV, 0 )

	invoke SetWindowRedraw( m_hWndTV, FALSE )

	invoke TreeView_DeleteAllItems( m_hWndTV)
	invoke EnableWindow, m_hWndEdit, FALSE

	xor esi,esi
	.while (esi < m_dwNumKeyPair)
		mov eax,sizeof KEYPAIR
		mul esi
		add eax, m_pKeyPair
		mov ebx,eax
		assume ebx:ptr KEYPAIR
		.if ([ebx].pszRoot != NULL)
			.if ([ebx].pszKey)
				mov iMax,0
			.else
				mov iMax,-1
			.endif
			invoke lstrcpy,addr szKey,[ebx].pszRoot

;------------------------------------ open the root: i.e. HKEY_CLASSES_ROOT/CLSID

			invoke RegOpenKeyEx, m_hRoot, addr szKey, 0, KEY_READ, addr hKey
			DebugOut "EditDlg, OnCommand, IDM_REFRESH: RegOpenKeyEx(%s)=%X", addr szKey, eax
			.if (eax == ERROR_SUCCESS)
				.if ([ebx].pszKey)
					mov bExpand, TRUE
				.else
					mov eax,[ebx].bExpand
					mov bExpand,eax
				.endif

;------------------------------------ look if key is already in editor
				invoke FindKey, NULL, [ebx].pszRoot
				.if (!eax)
;------------------------------------ this call will read the items
;------------------------------------ and insert them in treeview
;------------------------------------ here we may stop reading after first key
					invoke ReadAllSubItems, m_hWndTV, hKey, 0,
							addr szKey, bExpand, iMax
				.endif
				mov hTreeRoot,eax
				invoke RegCloseKey,hKey

;------------------------------------ now use second key
;------------------------------------ thats i.e. HKEY_CLASSES_ROOT/CLSID/{<GUID>}
;------------------------------------ here always read all subkeys

				.if ([ebx].pszKey)
					invoke lstrcat,addr szKey,CStr("\")
					invoke lstrcat,addr szKey,[ebx].pszKey
					invoke RegOpenKeyEx, m_hRoot, addr szKey, 0, KEY_READ, addr hKey
					.if (eax == ERROR_SUCCESS)
						invoke FindKey, hTreeRoot, [ebx].pszKey
						.if (!eax)
							invoke ReadAllSubItems, m_hWndTV, hKey, hTreeRoot,\
								[ebx].pszKey, [ebx].bExpand, -1
						.endif
						invoke RegCloseKey,hKey
					.endif
				.endif
				invoke TreeView_Expand( m_hWndTV, hTreeRoot, TVE_EXPAND)
			.elseif ( eax != ERROR_FILE_NOT_FOUND )
				mov ecx, eax
				invoke FormatMessage, FORMAT_MESSAGE_FROM_SYSTEM, 0,
					ecx, 0, addr szData, sizeof szData, 0
				invoke MessageBox, m_hWnd, addr szData, addr szKey, MB_OK
			.endif
		.endif
		inc esi
	.endw

	invoke SetWindowRedraw( m_hWndTV, TRUE)

	invoke SetCursor, hCursorOld
	mov m_bChanged,FALSE

	DebugOut "EditDlg, OnCommand, IDM_REFRESH exit"
	ret
	assume ebx:nothing
	align 4

OnRefresh endp


ExpandChildren proc uses esi hItem:HTREEITEM

	mov esi, hItem
	.while (esi)
		invoke TreeView_GetChild( m_hWndTV, esi)
		.if (eax)
			invoke ExpandChildren, eax
		.endif
		invoke TreeView_Expand( m_hWndTV, esi, m_dwExpandMode)
		invoke TreeView_GetNextSibling( m_hWndTV, esi)
		mov esi, eax
	.endw
	ret
	align 4

ExpandChildren endp


;*** WM_COMMAND, IDC_EXPAND


OnExpand proc

local	hCursorOld:HCURSOR

	invoke LoadCursor,NULL,IDC_WAIT
	invoke SetCursor, eax
	mov hCursorOld,eax
	invoke SetWindowRedraw( m_hWndTV, FALSE)
	invoke TreeView_GetRoot( m_hWndTV)
	invoke ExpandChildren, eax
	.if (m_dwExpandMode == TVE_EXPAND)
		mov m_dwExpandMode, TVE_COLLAPSE
	.else
		mov m_dwExpandMode, TVE_EXPAND
	.endif
	invoke SetWindowRedraw( m_hWndTV, TRUE)
	invoke SetCursor, hCursorOld
	ret
	align 4

OnExpand endp


;*** edit registry dialog has received WM_COMMAND


OnCommand proc uses ebx esi wParam:WPARAM, lParam:LPARAM

local	tvi:TVITEM

	movzx eax,word ptr wParam

	.if (eax == IDCANCEL)

		invoke PostMessage, m_hWnd, WM_CLOSE, 0, 0

	.elseif (eax == IDOK)

;------------------ is it really a "click" notification (LabelEdit!)
		.if (word ptr wParam+2 != BN_CLICKED)
			ret
		.endif
		invoke IsWindowVisible, m_hWndHE
		.if (eax)
			invoke SendMessage, m_hWndHE, EM_GETMODIFY, 0, 0
		.else
			movzx eax, m_bChanged
		.endif
		.if (eax)
			invoke UpdateRegistry, 0, CMD_EDIT, 0
			.if (!eax)
				ret
			.endif
		.endif

		invoke PostMessage,m_hWnd, WM_CLOSE, 0, 0

	.elseif (eax == IDC_EXPAND)

		invoke OnExpand

	.elseif (eax == IDM_DELETE)

		DebugOut "EditDlg, OnCommand, IDM_DELETE"

		mov eax,m_hTmpItem
		.if (eax == 0)
			mov eax,m_hSelItem
		.endif
		.if (eax == 0)
			invoke MessageBeep,MB_OK
			ret
		.endif
		mov tvi.hItem,eax
		mov tvi.mask_,TVIF_PARAM
		invoke TreeView_GetItem( m_hWndTV,addr tvi)
		.if (m_bConfirmDelete == TRUE)
			mov m_bMBDisplayed, TRUE
			.if (tvi.lParam)
				invoke MessageBox,m_hWnd,CStr("Are you sure?"),CStr("Delete registry value"),\
					MB_YESNO or MB_DEFBUTTON2 or MB_ICONQUESTION
			.else
				invoke MessageBox,m_hWnd,CStr("Are you sure?"),CStr("Delete registry key"),\
					MB_YESNO or MB_DEFBUTTON2 or MB_ICONQUESTION
			.endif
		.else
			mov eax,IDYES
		.endif
		.if (eax == IDYES)
			invoke TreeView_GetNextSibling( m_hWndTV, tvi.hItem )
			.if ( !eax )
				invoke TreeView_GetPrevSibling( m_hWndTV, tvi.hItem )
			.endif
			push eax
			.if (tvi.lParam)
				invoke UpdateRegistry, tvi.hItem, CMD_DELVALUE, 0
			.else
				invoke UpdateRegistry, tvi.hItem, CMD_DELKEY, 0
			.endif
			pop ecx
			.if ( eax && ecx )
				invoke TreeView_SelectItem( m_hWndTV, ecx )
			.endif
		.endif

	.elseif (eax == IDM_RENAME)

		mov eax,m_hTmpItem
		.if (eax == 0)
			mov eax,m_hSelItem
		.endif
		.if (eax == 0)
			invoke MessageBeep,MB_OK
			ret
		.endif
		invoke TreeView_EditLabel( m_hWndTV, eax)

	.elseif (eax == IDM_KEY)

		invoke UpdateRegistry, m_hTmpItem, CMD_NEWKEY, 0

	.elseif ((eax == IDM_STRING) || (eax == IDM_BINARY) || (eax == IDM_DWORD))

		.if (eax == IDM_DWORD)
			mov ecx,REG_DWORD
		.elseif (eax == IDM_BINARY)
			mov ecx,REG_BINARY
		.else
			mov ecx,REG_SZ
		.endif

		invoke UpdateRegistry, m_hTmpItem, CMD_NEWVALUE, ecx

	.elseif (eax == IDC_EDIT1)

		movzx ecx, word ptr wParam+2
		.if (ecx == EN_CHANGE)
			mov m_bChanged,TRUE
		.endif

	.elseif (eax == IDC_CUSTOM1)

		movzx ecx, word ptr wParam+2
		.if (ecx == EN_SETFOCUS)
			mov m_bHexEdHasFocus, TRUE
		.elseif (ecx == EN_KILLFOCUS)
			mov m_bHexEdHasFocus, FALSE
		.endif

	.elseif (eax == IDM_REFRESH)

		invoke OnRefresh

	.elseif (eax == IDM_SELECTALL)

		.if (m_bHexEdHasFocus)
			sub esp, sizeof CHARRANGE
			mov [esp].CHARRANGE.cpMin, 0
			mov [esp].CHARRANGE.cpMax, -1
			invoke SendMessage, m_hWndHE, EM_EXSETSEL, 0, esp
			add esp, size CHARRANGE
		.endif

	.else
		xor eax,eax
	.endif
	ret
	align 4

OnCommand endp


;*** handle WM_INITDIALOG


OnInitDialog proc

local	dwTVHeight:DWORD
local	dwHeight:DWORD
local	rect:RECT
local	rect2:RECT

	invoke SendMessage, m_hWnd, WM_SETICON, ICON_SMALL, g_hIconApp
	invoke SendMessage, m_hWnd, WM_SETICON, ICON_BIG, g_hIconApp
;------------------------- get some control hWnds

	invoke GetDlgItem, m_hWnd,IDC_TREE1
	mov m_hWndTV,eax
	invoke GetDlgItem, m_hWnd,IDC_EDIT1
	mov m_hWndEdit,eax
	invoke GetDlgItem, m_hWnd,IDC_CUSTOM1
	mov m_hWndHE,eax
	invoke GetDlgItem, m_hWnd,IDC_SPLITBTN
	mov m_hWndSplit,eax
	invoke Create@CSplittButton, eax, m_hWndTV, m_hWndEdit

	invoke TreeView_SetImageList( m_hWndTV, g_himlSmall, TVSIL_NORMAL)

;------------------------- get position of treeview in main window

	invoke GetWindowRect, m_hWndTV, addr rect
	invoke ScreenToClient, m_hWnd, addr rect
	mov eax,rect.left
	mov m_dwXPosTV, eax
	mov ecx,rect.top
	mov m_dwYPosTV, ecx

	invoke GetClientRect, m_hWnd, addr rect

;------------------------- get height of buttons

	invoke GetDlgItem, m_hWnd, IDOK
	mov ecx, eax
	invoke GetWindowRect, ecx, addr rect2
	mov eax, rect2.bottom 
	sub eax, rect2.top
	push eax
	invoke ScreenToClient, m_hWnd, addr rect2
	pop ecx
	mov eax, rect.bottom
	sub eax, rect2.top
	sub eax, ecx
	mov edx, eax
	add eax, ecx
	add eax, edx
	mov m_dwHeightBtns, eax

;------------------------- get height of edit in main window from bottom

	invoke GetWindowRect, m_hWndEdit, addr rect
	mov eax, g_dwHeightEdit
	.if (!eax)
		mov eax, rect.bottom
		sub eax, rect.top
	.endif
	mov m_dwHeightEdit, eax

;------------------------- set hexedit control to edit control place (hidden)

	invoke ScreenToClient, m_hWnd, addr rect.right
	mov ecx, rect.right
	sub ecx, rect.left
	mov edx, rect.bottom
	sub edx, rect.top
	invoke SetWindowPos, m_hWndHE, NULL, rect.left, rect.top, ecx, edx, SWP_NOZORDER

;------------------------- restore from last window's pos & size

	.if (g_rect.right)
		mov eax, g_rect.right
		sub eax, g_rect.left
		mov ecx, g_rect.bottom
		sub ecx, g_rect.top
		invoke SetWindowPos, m_hWnd, NULL, g_rect.left, g_rect.top, eax, ecx, SWP_NOZORDER or SWP_NOACTIVATE
	.endif

;------------------------- initiate a "refresh"

	invoke PostMessage, m_hWnd, WM_COMMAND, IDM_REFRESH, 0

;------------------------- now show dialog

	.if (m_bModeless)
		invoke ShowWindow, m_hWnd, SW_SHOWNORMAL
	.endif
	ret
	align 4

OnInitDialog endp


;*** Dialog Proc for "edit registry" dialog


EditRegistryDialog proc uses __this thisarg, message:dword,wParam:WPARAM,lParam:LPARAM

local	rect:RECT

	mov __this,this@

	mov eax,message

	.if (eax == WM_INITDIALOG)

		invoke OnInitDialog
		mov eax,1

	.elseif (eax == WM_CLOSE)

;---------------------------------------- get normal window pos & size
		sub esp, sizeof WINDOWPLACEMENT
		mov edx, esp
		mov [edx].WINDOWPLACEMENT.length_, sizeof WINDOWPLACEMENT
		invoke GetWindowPlacement, m_hWnd, edx
		mov edx, esp
		invoke CopyRect, addr g_rect, addr [edx].WINDOWPLACEMENT.rcNormalPosition
		add esp, sizeof WINDOWPLACEMENT

		.if (m_bModeless)
;			.if (m_bMBDisplayed)
				invoke SetActiveWindow, m_hWndParent
				DebugOut "EditRegistryDialog: SetActiveWindow( %X )=%X", m_hWndParent, eax
;			.endif
			invoke DestroyWindow, m_hWnd
		.else
			movzx eax,m_bRC
			invoke EndDialog, m_hWnd, eax
		.endif
		mov eax,1

	.elseif (eax == WM_DESTROY)

		invoke Destroy@CEditDlg, __this

	.elseif (eax == WM_COMMAND)

		invoke OnCommand, wParam, lParam

	.elseif (eax == WM_NOTIFY)

		invoke OnNotify, lParam

	.elseif (eax == WM_SIZE)

		.if (wParam != SIZE_MINIMIZED)
			movzx eax, word ptr lParam+0
			movzx ecx, word ptr lParam+2
			invoke OnSize, wParam, eax, ecx
		.endif

	.elseif (eax == WM_ACTIVATE)

		movzx eax,word ptr wParam
		.if (eax == WA_INACTIVE)
			mov g_hWndDlg, NULL
		.else
			mov eax,m_hWnd
			mov g_hWndDlg, eax
		.endif
if 0
	.elseif (eax == WM_WINDOWPOSCHANGING)
		mov ecx, lParam
		.if (!([ecx].WINDOWPOS.flags & SWP_NOSIZE))
			mov edx, m_dwHeightBtns
			add edx, m_dwYPosTV
			add edx, 16
			invoke SetRect, addr rect, 0, 0, 1, edx
			invoke AdjustWindowRect, addr rect, WS_OVERLAPPEDWINDOW, FALSE
			mov ecx, lParam
			mov edx, rect.bottom
			sub edx, rect.top
			.if (edx > [ecx].WINDOWPOS.cy)
				mov [ecx].WINDOWPOS.cy, edx
			.endif
		.endif
		mov eax, 1
endif
	.elseif (eax == WM_SIZING)
		mov edx, m_dwHeightBtns
		add edx, m_dwYPosTV
		add edx, 32
		invoke SetRect, addr rect, 0, 0, 1, edx
		invoke AdjustWindowRect, addr rect, WS_OVERLAPPEDWINDOW, FALSE
		mov edx, rect.bottom
		sub edx, rect.top
		mov ecx, lParam
		mov eax, [ecx].RECT.bottom
		sub eax, [ecx].RECT.top
		sub eax, edx
		.if (CARRY?)
			neg eax
			add [ecx].RECT.bottom, eax
		.endif
		mov eax, 1
if ?HTMLHELP
	.elseif (eax == WM_HELP)

		invoke DoHtmlHelp, HH_DISPLAY_TOPIC, CStr("editdialog.htm")
endif
	.else
		xor eax,eax ;indicates "no processing"
	.endif
	ret
	align 4

EditRegistryDialog endp

;*** native dialog proc: get this_ pointer, then call dialog proc

editdialogproc proc hWnd:HWND,message:dword,wParam:WPARAM,lParam:LPARAM

;    DebugOut "editdialogproc, message=%X, wParam=%X, lParam=%X", message, wParam, lParam

	.if (message == WM_INITDIALOG)
		invoke SetWindowLong,hWnd,DWL_USER,lParam
		mov eax,lParam
		mov ecx,hWnd
		mov [eax].CEditDlg.hWnd,ecx
	.else
		invoke GetWindowLong,hWnd,DWL_USER
	.endif
	.if (eax != 0)
		invoke EditRegistryDialog,eax,message,wParam,lParam
		.if (message == WM_DESTROY)
			invoke SetWindowLong, hWnd, DWL_USER, 0
		.endif
	.endif
	ret
	align 4

editdialogproc endp

SetCookie@CEditDlg proc public uses __this thisarg, dwCookie:DWORD

	mov __this,this@
	mov eax,dwCookie
	mov m_dwCookie,eax
	ret
	align 4

SetCookie@CEditDlg endp


;*** now show editor dialog box
;*** the editor must have been told what to display with
;*** previous SetKeys() calls


Show@CEditDlg proc public uses __this thisarg

	mov __this,this@

	.if (m_bModeless)
		invoke CreateDialogParam,g_hInstance,IDD_EDITREGDLG,m_hWndParent,editdialogproc,__this
	.else
		invoke DialogBoxParam,g_hInstance,IDD_EDITREGDLG,m_hWndParent,editdialogproc,__this
	.endif
	ret
	align 4

Show@CEditDlg endp


SetRoot@CEditDlg proc public uses __this thisarg, hRoot:HANDLE

	mov __this,this@

	mov eax, hRoot
	mov m_hRoot, eax
	ret
	align 4

SetRoot@CEditDlg endp


;*** set the keys for the editor
;*** this must be done before calling Show method


SetKeys@CEditDlg proc public uses ebx esi __this thisarg, numPairs:DWORD, pKP:ptr KEYPAIR

	mov __this,this@

	mov ecx,m_dwNumKeyPair

	add ecx,numPairs
	mov eax,sizeof KEYPAIR
	mul ecx
	invoke malloc, eax
	.if (eax == NULL)
		ret
	.endif
	mov esi,m_pKeyPair
	mov m_pKeyPair,eax
	mov ebx,eax
	mov ecx,m_dwNumKeyPair
	.if (ecx)
		mov eax, sizeof KEYPAIR
		mul ecx
		mov ecx,eax
		mov eax,esi
		push edi
		mov edi,ebx
		rep movsb
		mov ebx,edi
		pop edi
		invoke free,eax
	.endif
;--------------------------------------- now add the new key pairs
	mov esi,pKP
	mov eax,numPairs
	add m_dwNumKeyPair,eax

	.if (m_bModeless)
		.while (eax)
			push eax
			mov [ebx].KEYPAIR.pszRoot,NULL
			.if ([esi].KEYPAIR.pszRoot)
				invoke lstrlen,[esi].KEYPAIR.pszRoot
				inc eax
				invoke malloc, eax
				.if (eax)
					mov [ebx].KEYPAIR.pszRoot,eax
					invoke lstrcpy, eax, [esi].KEYPAIR.pszRoot
				.endif
			.endif
			mov [ebx].KEYPAIR.pszKey,NULL
			.if ([esi].KEYPAIR.pszKey)
				invoke lstrlen,[esi].KEYPAIR.pszKey
				inc eax
				invoke malloc, eax
				.if (eax)
					mov [ebx].KEYPAIR.pszKey,eax
					invoke lstrcpy, eax, [esi].KEYPAIR.pszKey
				.endif
			.endif
			mov eax,[esi].KEYPAIR.bExpand
			mov [ebx].KEYPAIR.bExpand,eax
			pop eax
			add esi,sizeof KEYPAIR
			add ebx,sizeof KEYPAIR
			dec eax
		.endw
	.else
		mov ecx,sizeof KEYPAIR
		mul ecx
		mov ecx,eax
		push edi
		mov edi,ebx
		rep movsb
		mov ebx,edi
		pop edi
	.endif
	ret
	align 4

SetKeys@CEditDlg endp

;*** CEditDlg constructor

Create@CEditDlg proc public uses __this hWndParent:HWND, bModeless:BOOL, bConfirmDelete:BOOL

	DebugOut "Create@CEditDlg"

	invoke malloc,sizeof CEditDlg
	.if (eax == NULL)
		ret
	.endif

	mov __this,eax

	.if (!g_hMenuEdit)
		invoke Init@CEditDlg
	.endif

	invoke Init@CHexEdit, g_hInstance

	mov eax, g_dwHeightEdit
	mov m_dwHeightEdit, eax

	mov eax,hWndParent
	mov m_hWndParent,eax
	mov m_hRoot, HKEY_CLASSES_ROOT

	mov eax,bConfirmDelete
	mov m_bConfirmDelete,al
	mov eax,bModeless
	mov m_bModeless,al

	mov m_dwExpandMode, TVE_EXPAND
	mov m_dwCookie,-1

	return __this
	align 4

Create@CEditDlg endp

;*** CEditDlg destructor

Destroy@CEditDlg proc public uses __this thisarg

	DebugOut "Destroy@CEditDlg"

	mov __this,this@

	.if (m_bModeless)
		push esi
		mov esi,m_pKeyPair
		mov ecx,m_dwNumKeyPair
		push esi
		.while (ecx)
			push ecx
			.if ([esi].KEYPAIR.pszRoot)
				invoke free, [esi].KEYPAIR.pszRoot
			.endif
			.if ([esi].KEYPAIR.pszKey)
				invoke free, [esi].KEYPAIR.pszKey
			.endif
			add esi,sizeof KEYPAIR
			pop ecx
			dec ecx
		.endw
		pop esi
		.if (esi)
			invoke free, esi
		.endif
		pop esi
	.endif
	invoke free,__this
	ret
	align 4

Destroy@CEditDlg endp

;*** static initialization/deinitialization functions

__this	textequ <error>		;do not use member variables here

;*** create image list for registry editor

CreateImageLists proc uses esi

local	_cx:dword
local	_cy:dword

;------------ set the small image list

	invoke GetSystemMetrics, SM_CXSMICON
	mov _cx, eax
	invoke GetSystemMetrics, SM_CYSMICON
	mov _cy, eax

	.if (g_himlSmall)
		invoke ImageList_Destroy, g_himlSmall
	.endif

	invoke ImageList_Create, _cx, _cy, ILC_COLORDDB or ILC_MASK, NUMICONS, 0
	mov g_himlSmall, eax

	.if (eax)
		mov esi,offset IconTab
		mov ecx,NUMICONS
		.while (ecx)
			push ecx
			lodsd
			invoke LoadImage, g_hInstance, eax, IMAGE_ICON, _cx, _cy, LR_DEFAULTCOLOR
			invoke ImageList_AddIcon(g_himlSmall, eax)
			pop ecx
			dec ecx
		.endw
	.endif

;------------ set the large image list

	.if (g_himlLarge)
		invoke ImageList_Destroy,g_himlLarge
	.endif

	invoke GetSystemMetrics, SM_CXICON
	mov _cx, eax
	invoke GetSystemMetrics, SM_CYICON
	mov _cy, eax

	invoke ImageList_Create, _cx, _cy, ILC_COLORDDB or ILC_MASK, NUMICONS, 0
	mov g_himlLarge, eax

	.if (eax)
		mov esi,offset IconTab
		mov ecx,NUMICONS
		.while (ecx)
			push ecx
			lodsd
			invoke LoadImage, g_hInstance, eax, IMAGE_ICON, _cx, _cy, LR_DEFAULTCOLOR
			invoke ImageList_AddIcon(g_himlLarge, eax)
			pop ecx
			dec ecx
		.endw
	.endif

	ret
	align 4

CreateImageLists endp

;*** delete image lists for registry editor

DestroyImageLists proc

	.if (g_himlSmall)
		invoke ImageList_Destroy, g_himlSmall
	.endif
	.if (g_himlLarge)
		invoke ImageList_Destroy, g_himlLarge
	.endif
	ret
	align 4

DestroyImageLists endp

Deinit@CEditDlg proc

	invoke DestroyImageLists
	.if (g_hMenuEdit)
		invoke DestroyMenu,g_hMenuEdit
		mov g_hMenuEdit,0
	.endif
	ret
	align 4

Deinit@CEditDlg endp


Init@CEditDlg proc

	invoke CreateImageLists
	invoke LoadMenu, g_hInstance, IDR_MENU3
	mov g_hMenuEdit,eax
	invoke atexit, offset Deinit@CEditDlg
	ret
	align 4

Init@CEditDlg endp

	end
