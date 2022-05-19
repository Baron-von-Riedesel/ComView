
;*** definition of CDocument methods 
;*** this class holds all list infos of main dialog

	.386
	.model flat,stdcall
	option casemap :none   ; case sensitive
	option proc:private

INSIDE_CDOCUMENT equ 1

	include COMView.inc
	include classes.inc
	include rsrc.inc
	include debugout.inc

?USEHASH	equ 1			; 0 doesnt work

BEGIN_CLASS CDocument
pszRoot		LPSTR ?			; key root in HKCR
hWnd		HWND  ?			; parent window (for message boxes)
pItems		PVOID ?			; pointer to start of items
dwCurItem	DWORD ?			; current item
numItems	DWORD ?			; number of items
numStrng	DWORD ?			; number of strings in item
ItemSize	DWORD ?			; size of 1 item
numFree		DWORD ?			; number of free items in document
pStrPool	PVOID ?			; pointer to start of string pool
PoolSize	DWORD ?			; free memory in pool
pStrNext	PVOID ?			; pointer to free memory in pool
END_CLASS

__this	textequ <edi>
_this	textequ <[__this].CDocument>
thisarg equ <this@:ptr CDocument>

;*** private methods. thisarg can be omitted cause "this" ptr (=edi) is always set

MemoryError				proto :HWND
AllocStringPoolBlock	proto :DWORD
SetCurrentItem			proto :DWORD
SetItemData				proto :DWORD, :DWORD
SetItemString			proto :DWORD, :LPSTR, :DWORD
GetClassModulePath		proto :HANDLE, :LPSTR, iMax:dword, pszKey:LPSTR
RefreshLineCLSID		proto :HANDLE, :DWORD, :LPSTR, :DWORD, :LPSTR, :LPSTR
RefreshCLSID			proto
ListTypeLibVersions		proto :HANDLE, :DWORD, :LPSTR
RefreshTypeLib			proto
RefreshInterface		proto
RefreshAppID			proto
RefreshCompCat			proto
RefreshHKCR				proto
RefreshObject			proto
ResizeTable				proto
AllocTable				proto iItems:DWORD

	MEMBER pszRoot, hWnd, pItems, dwCurItem, numItems, numStrng
	MEMBER ItemSize, numFree, pStrPool, PoolSize, pStrNext

VMEMSIZE	equ 40000h			;size of further mem blocks for string pool

externdef NUMREGKEYSEARCHES:abs

;*** sorting must be done by hand
;*** we use qsort() from CRTDLL.DLL, which should be installed on every Win32 system
;*** define prototypes of qsort (+compareproc) so we can use invoke

QUICKSORTCP		typedef proto c a1:ptr, a2:ptr
LPQUICKSORTCP	typedef ptr QUICKSORTCP
QUICKSORT		typedef proto c pArray:ptr, numelem:DWORD, sizelem:DWORD, :LPQUICKSORTCP
LPQUICKSORT		typedef ptr QUICKSORT

REFRESHPROC		typedef proto
PREFRESHPROC	typedef ptr REFRESHPROC

_CMode struct
iMode	dd ?
pRefresh PREFRESHPROC ?
_CMode ends

	.data

g_iSortCol	dd 0			;for sorting
g_iSortDir	dd 0			;for sorting
g_dwBase	dd 0			;for sorting
g_hCrtLib	HANDLE 0		;hLib of CRTDLL.DLL
_qsort		LPQUICKSORT 0	;procaddress of qsort

	.const

ModeTab	label dword
	dd MODE_CLSID,		RefreshCLSID
	dd MODE_TYPELIB,	RefreshTypeLib
	dd MODE_INTERFACE,	RefreshInterface
	dd MODE_APPID,		RefreshAppID
	dd MODE_COMPCAT,	RefreshCompCat
	dd MODE_HKCR,		RefreshHKCR
	dd MODE_OBJECT,		RefreshObject
	dd MODE_ROT,		RefreshROT
NUMMODES equ ($ - ModeTab) / 8

	.code

;--- throw exception in case of insufficient memory 

MemoryError proc hWnd:HWND

ifndef __JWASM__
	nop	;Masm needs this
endif
	invoke MessageBox, hWnd, CStr("Not enough memory"), 0, MB_OK
	invoke RaiseException, STATUS_NO_MEMORY, EXCEPTION_NONCONTINUABLE, NULL, NULL
	xor eax,eax
	ret
	align 4

MemoryError endp


SplitUserdefinedColumn proc pszSource:LPSTR, pszKey:LPSTR, pszValueName:LPSTR

		mov edx, pszSource
		mov ecx, pszKey
		mov eax, pszValueName
		mov byte ptr [eax],0
		.if (byte ptr [edx])
			xor eax,eax
			.while (1)
				mov ah,al
				mov al,[edx]
				.break .if (!al)
				inc edx
				.if ((ax == "[") || (ax == "\["))
					.if (ah)
						mov byte ptr [ecx-1],0
					.endif
					push ecx
					invoke lstrcpy, pszValueName, edx
					invoke lstrlen, pszValueName
					mov ecx,pszValueName
					mov byte ptr [ecx+eax-1],0
					pop ecx
					.break
				.endif
				mov [ecx],al
				inc ecx
			.endw
			mov byte ptr [ecx],0
		.endif
		ret
		align 4

SplitUserdefinedColumn endp

;--- alloc new block of string pool


AllocStringPoolBlock proc dwSize:DWORD

		invoke VirtualAlloc, NULL, dwSize, MEM_COMMIT, PAGE_READWRITE
		.if (!eax)
			invoke MemoryError, m_hWnd
			ret
		.endif
		lea ecx,[eax+4]
		mov m_pStrNext,ecx
		mov edx, dwSize
		sub edx,4
		mov m_PoolSize,edx

		lea edx, m_pStrPool
		mov ecx,[edx]
		.while (ecx)
			mov edx,ecx
			mov ecx,[edx]
		.endw
		mov [eax],ecx
		mov [edx],eax
		ret
		align 4

AllocStringPoolBlock endp


;*** set current item for SetItemData/SetItemString.
;*** for sequential insert only!


SetCurrentItem proc dwItem:DWORD

		mov eax,dwItem
		.if (eax >= m_numItems)
			.if (!m_numFree)
				invoke ResizeTable
			.endif
			inc m_numItems
			dec m_numFree
		.endif
		mov eax,dwItem
		mov m_dwCurItem,eax
		ret
		align 4

SetCurrentItem endp

if ?REMOVEITEM
;--- remove item from view

RemoveItem@CDocument proc public uses __this esi thisarg, dwItem:DWORD

local pDest:LPSTR

		mov __this,this@

		mov eax, dwItem
		mov ecx, m_numItems
		.if (eax >= ecx)
			ret
		.endif
		mov esi, m_ItemSize
		mul esi
		add eax, m_pItems
		mov pDest, eax
		add esi, eax
		mov eax, ecx
		dec eax
		sub eax, dwItem
		mul m_ItemSize
		mov ecx, eax
		mov eax, m_ItemSize
		push edi
		mov edi, pDest
		.if (ecx)
			shr ecx, 2
			rep movsd
		.endif
		mov ecx, eax
		shr ecx, 2
		xor eax, eax
		rep stosd
		pop edi
		dec m_numItems
		inc m_numFree
		mov eax, m_dwCurItem
		.if (eax >= m_numItems)
			dec m_dwCurItem
		.endif
		ret
		align 4

RemoveItem@CDocument endp
endif

;*** get item data of document


GetItemData@CDocument proc public uses __this thisarg, dwItem:DWORD, dwSubItem:DWORD

		mov __this,this@

		mov eax,dwItem
		mul m_ItemSize
		add eax, m_pItems
		mov ecx,dwSubItem
		mov eax,[eax+ecx*4]
		ret
		align 4

GetItemData@CDocument endp


;*** set data of subitem of current item
;*** SetCurrentItem should be called before

SetItemData proc dwSubItem:DWORD, dwValue:DWORD

		mov ecx, m_pItems
		mov eax, m_dwCurItem
		mul m_ItemSize
		add ecx,eax
		mov edx,dwSubItem
		shl edx,2
		add ecx,edx
		mov eax,dwValue
		mov [ecx],eax
		ret
		align 4

SetItemData endp


;*** set string item of current item
;*** SetCurrentItem should be called before


SetItemString proc uses esi ebx dwSubItem:DWORD, pSrc:LPSTR, dwSize:DWORD

		mov eax,dwSize
		.if (eax == -1)
			invoke lstrlen, pSrc
		.endif
		add eax,4
		and al,0FCh
		mov ebx,eax

;------------------------------------- enough room in string pool?
		.if (eax > m_PoolSize)
;------------------------------------- no, get another memory block (256kB)
			invoke AllocStringPoolBlock, VMEMSIZE
		.endif
		
		sub m_PoolSize,ebx

		invoke SetItemData, dwSubItem, m_pStrNext
		push edi
		mov ecx,ebx
		mov edi, m_pStrNext
		mov esi,pSrc
		shr ecx,2
		rep movsd
		mov eax,edi
		pop edi
		mov m_pStrNext,eax
		ret
		align 4

SetItemString endp


;--- set flags in itemdata


SetItemFlag@CDocument proc public uses __this thisarg, dwItem:DWORD, bFlag:DWORD, bMask:DWORD

		mov __this,this@

		mov eax, dwItem
		.if (eax >= m_numItems)
			.if (eax == -1)
				push esi
				xor esi, esi
				.while (esi < m_numItems)
					invoke SetItemFlag@CDocument, __this, esi, bFlag, bMask
					inc esi
				.endw
				pop esi
			.endif
			ret
		.endif
		mul m_ItemSize
		add eax, m_ItemSize
		mov edx,bMask
		xor dl,-1
		mov ecx, m_pItems
		and byte ptr [ecx+eax-4],dl
		mov edx,bFlag
		or byte ptr [ecx+eax-4],dl
		ret
		align 4

SetItemFlag@CDocument endp


;*** return flags of an item


GetItemFlag@CDocument proc public uses __this thisarg, dwItem:DWORD, bMask:DWORD
		
		mov __this,this@

		mov eax,dwItem
		mul m_ItemSize
		add eax, m_ItemSize
		mov ecx, m_pItems
		mov al,[ecx+eax-4]
		and al,byte ptr bMask
		ret
		align 4

GetItemFlag@CDocument endp


;*** return number of items


GetItemCount@CDocument proc public uses __this thisarg

		mov __this,this@

		mov eax, m_numItems
		ret
		align 4

GetItemCount@CDocument endp


;*** compare proc for sorting listview (cdecl type)


compareproc proc c pItem1:ptr, pItem2:ptr

local	dwTmp1:dword
local	dwTmp2:dword

		mov eax,g_iSortCol
		mov ecx,pItem1
		mov edx,pItem2
		mov ecx,[ecx+eax*4]
		mov edx,[edx+eax*4]
		.if (!ecx)
			mov ecx,CStr("")
		.endif
		.if (!edx)
			mov edx,CStr("")
		.endif
		mov eax,g_dwBase
		.if (eax)	
			push eax
			push edx
			invoke String2Number, ecx, addr dwTmp1, eax
			pop edx
			pop eax
			invoke String2Number, edx, addr dwTmp2, eax
			mov eax,dwTmp1
			mov ecx,dwTmp2
			.if (g_iSortDir == 1)
				xchg eax,ecx
			.endif
			sub eax,ecx
		.else
			.if (g_iSortDir == 0)
				invoke lstrcmp, ecx, edx
			.else
				invoke lstrcmp, edx, ecx
			.endif
		.endif
		ret
		align 4

compareproc endp


;*** resort document


Sort@CDocument proc public uses __this thisarg, dwIndex:DWORD, iDirection:DWORD, dwFlags:DWORD

		mov __this,this@
;------------------------ theres a good chance MSVCRT is already loaded
;------------------------ if so, use it, else load CRTDLL
		.if (!g_hCrtLib)
			invoke GetModuleHandle,CStr("MSVCRT")
			.if (!eax)
				invoke LoadLibrary,CStr("CRTDLL")
			.endif
			mov g_hCrtLib,eax
			.if (eax)
				invoke GetProcAddress, g_hCrtLib, CStr("qsort")
				mov _qsort,eax
			.endif
		.endif
		.if (!_qsort)
			ret
		.endif

;-------------------------------- compareproc has no user parameter, so
;-------------------------------- save parms in global var
		mov eax,dwIndex
		mov g_iSortCol,eax
		mov eax,iDirection
		mov g_iSortDir,eax
		mov eax,dwFlags
		mov g_dwBase,eax		;ax <> 0 -> numeric compare

		invoke _qsort, m_pItems, m_numItems, m_ItemSize, compareproc

		ret
		align 4

Sort@CDocument endp


;*** not implemented yet. Should refresh an item if editor has changed it


RefreshItem@CDocument proc public uses esi ebx __this thisarg, dwItem:DWORD, pszKey:LPSTR
		
local szKey[256]:byte
if 0
		mov __this,this@
		mov eax,dwItem
		.if (eax < m_numItems)
			invoke GetItemData@CDocument, __this, eax, 0
			.if (m_pszRoot)
				invoke wsprintf, addr szKey, CStr("%s\%s"), m_pszRoot, eax
			.else
				invoke wsprintf, addr szKey, CStr("%s"), eax
			.endif
			pushad
			mov ecx,eax
			mov esi,pszKey
			lea edi,szKey
			repz cmpsb
			popad
			.if (ZERO?)
				mov eax,eax		;do refresh
			.endif
		.endif
endif
		ret
		align 4
RefreshItem@CDocument endp


;*** get "module path" of a CLSID entry
;*** helper proc for RefreshCLSID


GetClassModulePath proc uses ebx hKey:HANDLE, pszValue:LPSTR,
						iMax:dword, pszKey:LPSTR

local	hSubKey:HANDLE
local	iType:dword
local	pszLastKey:LPSTR
local	dwMax:DWORD

		xor ebx, ebx
		mov pszLastKey, ebx
		.while (ebx < NUMREGKEYSEARCHES)
			invoke RegOpenKeyEx,hKey,[ebx*sizeof dword+offset pRegKeys],0,KEY_READ,addr hSubKey
			.if (eax == ERROR_SUCCESS)
				mov eax, [ebx*sizeof dword+offset pRegKeys]
				mov pszLastKey, eax
				mov eax,iMax
				mov dwMax,eax
				mov ecx, pszValue
				mov byte ptr [ecx],0
				invoke RegQueryValueEx,hSubKey,addr g_szNull,NULL,addr iType, ecx, addr dwMax
				push eax
				invoke RegCloseKey,hSubKey
				pop eax
				.break .if ((eax == S_OK) && (dwMax > 1))
			.endif
			inc ebx
		.endw
		.if (pszLastKey)
			invoke lstrcpy, pszKey, pszLastKey
		.endif
		ret
		align 4

GetClassModulePath endp


;*** set 1 item of list HKEY_CLASSES_ROOT\CLSID ***


RefreshLineCLSID proc hKey:HANDLE, dwItem:DWORD, pszClsId:LPSTR, dwSize:DWORD, pszUserKey:LPSTR, pszUserValueName:LPSTR

local	hSubKey:HANDLE
local	hSubKey2:HANDLE
local	dwType:dword
local	filetime:FILETIME
local	szKey[256]:byte
local	szValue[256]:byte
local	szValueName[64]:byte

		invoke SetCurrentItem, dwItem
		invoke SetItemString, 0, pszClsId, dwSize
		invoke RegOpenKeyEx, hKey, pszClsId, NULL, KEY_READ, addr hSubKey
		.if (eax == ERROR_SUCCESS)
			mov dwSize,sizeof szValue
			invoke RegQueryValueEx,hSubKey,addr g_szNull,NULL,addr dwType,addr szValue,addr dwSize
			.if (eax == ERROR_SUCCESS)
				invoke SetItemString, 1, addr szValue, dwSize
			.endif
;--------------------------------------------- search for entry InProcServer32/LocalServer32 ...
			mov szKey, 0
			mov szValue, 0
			invoke GetClassModulePath, hSubKey, addr szValue, sizeof szValue, addr szKey
			.if (szKey)
				invoke SetItemString, 2, addr szKey, -1
			.endif
			.if (szValue)
				invoke SetItemString, 3, addr szValue, -1
			.endif
;--------------------------------------------- search for entry ProgID
			invoke RegOpenKeyEx,hSubKey,CStr("ProgID"),0,KEY_READ,addr hSubKey2
			.if (eax == ERROR_SUCCESS)
				mov dwSize,sizeof szValue
				invoke RegQueryValueEx,hSubKey2,addr g_szNull,NULL,addr dwType,addr szValue,addr dwSize
				.if (eax == ERROR_SUCCESS)
					invoke SetItemString, 4, addr szValue, dwSize
				.endif
				invoke RegCloseKey,hSubKey2
			.endif
;--------------------------------------------- search for entry TypeLib
			invoke RegOpenKeyEx,hSubKey, addr g_szTypeLib,0,KEY_READ,addr hSubKey2
			.if (eax == ERROR_SUCCESS)
				mov dwSize,sizeof szValue
				invoke RegQueryValueEx,hSubKey2,addr g_szNull,NULL,addr dwType,addr szValue,addr dwSize
				.if (eax == ERROR_SUCCESS)
					invoke SetItemString, 5, addr szValue, dwSize
				.endif
				invoke RegCloseKey,hSubKey2
			.endif
;--------------------------------------------- search for userdefined entry
			.if (g_szUserColCLSID)
				mov ecx, pszUserKey
				.if (byte ptr [ecx])
					invoke RegOpenKeyEx,hSubKey, ecx,0,KEY_READ,addr hSubKey2
				.else
					invoke RegOpenKeyEx,hKey,pszClsId,0,KEY_READ,addr hSubKey2
				.endif
				.if (eax == ERROR_SUCCESS)
					mov dwSize,sizeof szValue - 1
					invoke RegQueryValueEx, hSubKey2, pszUserValueName, NULL,
							addr dwType, addr szValue+1, addr dwSize
					.if (eax != ERROR_SUCCESS)
						mov szValue, 0
						mov ecx, pszUserValueName
						.if (byte ptr [ecx] == 0)
							invoke lstrcpy,addr szValue,CStr("<no default>")
						.endif
					.else
						mov szValue,'"'
						invoke lstrlen,addr szValue
						lea ecx,szValue
						mov word ptr [eax+ecx],'"'
					.endif
					invoke SetItemString, 6, addr szValue, -1
					invoke RegCloseKey,hSubKey2
				.endif
			.endif
			invoke RegCloseKey,hSubKey
		.endif
		ret
		align 4

RefreshLineCLSID endp


;*** list HKEY_CLASSES_ROOT\CLSID entries ***


RefreshCLSID proc uses ebx esi

local	hKey:HANDLE
local	dwSize:dword
local	dwType:dword
local	dwItems:dword
local	filetime:FILETIME
local	szClsId[256]:byte
local	szUserKey[64]:byte
local	szUserValueName[64]:byte

		invoke SplitUserdefinedColumn, offset g_szUserColCLSID, addr szUserKey, addr szUserValueName

		invoke RegOpenKeyEx, HKEY_CLASSES_ROOT, m_pszRoot, 0, KEY_READ, addr hKey
		.if (eax == ERROR_SUCCESS)

			invoke RegQueryInfoKey, hKey, NULL, 0, NULL,
				addr dwItems, NULL, NULL, NULL, NULL, NULL, NULL, NULL
			.if (eax != ERROR_SUCCESS)
				invoke wsprintf, addr szClsId, CStr("RegQueryInfoKey failed [%X]"), eax
				invoke MessageBox, m_hWnd, addr szClsId, 0, MB_OK
				invoke RegCloseKey, hKey
				ret
			.endif

			invoke AllocTable, dwItems

			xor ebx,ebx
			.while (ebx < dwItems)
				mov dwSize,sizeof szClsId
				invoke RegEnumKeyEx,hKey,ebx,addr szClsId,addr dwSize,0,NULL,0,addr filetime
				.break .if (eax != ERROR_SUCCESS)
				invoke RefreshLineCLSID, hKey, ebx, addr szClsId, dwSize, addr szUserKey, addr szUserValueName
				inc ebx
			.endw
			invoke RegCloseKey,hKey
		.endif
		ret
		align 4

RefreshCLSID endp


;*** list "<version>\win32" key entry in a HKCR\TypeLib\{<GUID>} entry ***
;*** hKey: handle to TypeLib\GUID entry ***


ListTypeLibVersions proc uses ebx hKey:HANDLE, dwItem:DWORD, pszGUID:LPSTR

local	hSubKey:HANDLE
local	hSubKey2:HANDLE
local	szKey[256]:byte
local	szKey2[256]:byte
local	szLCID[256]:byte
local	szValue[256]:byte
local	szValue2[256]:byte
local	dwSize:dword
local	filetime:FILETIME
local	dwIndex:DWORD
local	guid:GUID
local	lcid:LCID
local	dwMajor:DWORD
local	dwMinor:DWORD
local	bstr:BSTR
local	bFirst:BOOL

		xor ebx,ebx
		.while (1)
;-------------------------------------- enum all versions
			mov dwSize,sizeof szKey
			invoke RegEnumKeyEx,hKey,ebx,addr szKey,addr dwSize,0,NULL,0,addr filetime
			.break .if (eax != ERROR_SUCCESS)

			invoke SetCurrentItem, dwItem
			inc dwItem
;-------------------------------------- write GUID (pszGUID)
			invoke SetItemString, 0, pszGUID, -1
;-------------------------------------- write Version key (szKey)
			invoke SetItemString, 3, addr szKey, dwSize

			mov dwSize,sizeof szValue2
			mov byte ptr szValue2,0
			invoke RegQueryValue,hKey,addr szKey,addr szValue2,addr dwSize
;-------------------------------------- write std value of Version key (szValue2)
			invoke SetItemString, 1, addr szValue2, dwSize

			invoke RegOpenKeyEx,hKey,addr szKey,0,KEY_READ,addr hSubKey
			.if (eax == ERROR_SUCCESS)
				mov dwIndex,0
				mov bFirst,TRUE
;-------------------------------------- enum all LCIDs
				.while (1)
					mov dwSize,sizeof szKey2
					invoke RegEnumKeyEx,hSubKey, dwIndex,addr szKey2,addr dwSize,0,NULL,0,addr filetime
					.break .if (eax != ERROR_SUCCESS)
					mov al,byte ptr szKey2
;-------------------------------------- is it a LCID?
					.if (al >= '0' && (al <= '9'))
						.if (bFirst == FALSE)
							invoke SetCurrentItem, dwItem
							inc dwItem
;-------------------------------------- write GUID (pszGlsid)
							invoke SetItemString, 0, pszGUID, -1
;-------------------------------------- write Version key (szKey)
							invoke SetItemString, 3, addr szKey, -1
;-------------------------------------- write std value of Version key (szValue2)
							invoke SetItemString, 1, addr szValue2, -1
						.endif

;-------------------------------------- write LCID key (szKey2)
						invoke SetItemString, 4, addr szKey2, dwSize

;-------------------------------------- now get path
						.if (g_bUseQueryPath)
							invoke GUIDFromLPSTR, pszGUID, addr guid
							invoke String2Number, addr szKey2, addr lcid, 16
							invoke String22DWords, addr szKey, addr dwMajor, addr dwMinor
							invoke SysAllocStringByteLen, NULL, MAX_PATH
							mov bstr,eax
;-------------------------------------- using QueryPathOfRegTypeLib will
;-------------------------------------- return win16 entry if no win32 entry exists
							invoke QueryPathOfRegTypeLib, addr guid,\
								dwMajor, dwMinor, lcid, addr bstr
							.if (eax == S_OK)
								invoke WideCharToMultiByte,CP_ACP,0,\
									bstr,-1,addr szValue, sizeof szValue,0,0
								mov dwSize,eax
							.endif
							.if (bstr)
								invoke SysFreeString, bstr
								mov eax,ERROR_SUCCESS
							.endif
						.else
;-------------------------------------- search win32 entry by hand (default)
							invoke lstrcat,addr szKey2,CStr("\win32")
							invoke RegOpenKeyEx,hSubKey,addr szKey2,0,KEY_READ,addr hSubKey2
							.if (eax == ERROR_SUCCESS)
								mov dwSize,sizeof szValue
								invoke RegQueryValue,hSubKey2,0,addr szValue,addr dwSize
								push eax
								invoke RegCloseKey,hSubKey2
								pop eax
							.endif
						.endif
						.if (eax == ERROR_SUCCESS)
;-------------------------------------- write path of exe/dll (szValue)
							invoke SetItemString, 2, addr szValue, dwSize
						.endif
						mov bFirst,FALSE
					.endif
					inc dwIndex
				.endw

				invoke RegCloseKey,hSubKey
			.endif
			inc ebx
		.endw
		mov eax,dwItem
		ret
		align 4

ListTypeLibVersions endp


;*** list HKEY_CLASSES_ROOT\TypeLib entries ***


RefreshTypeLib proc uses ebx

local	hKey:HANDLE
local	dwItems:DWORD
local	dwItem:DWORD
local	hGuId:HANDLE
local	szText[256]:byte
local	dwSize:dword
local	filetime:FILETIME

		invoke RegOpenKeyEx,HKEY_CLASSES_ROOT,m_pszRoot,0,KEY_READ,addr hKey
		.if (eax == ERROR_SUCCESS)

;----------------------------------- since some libs can take more than 1 line
;----------------------------------- this query is only a first hint
			invoke RegQueryInfoKey,hKey, NULL, 0, NULL,
				addr dwItems, NULL, NULL, NULL, NULL, NULL, NULL, NULL
			.if (eax != ERROR_SUCCESS)
				invoke wsprintf, addr szText, CStr("RegQueryInfoKey failed [%X]"), eax
				invoke MessageBox, m_hWnd, addr szText, 0, MB_OK
				invoke RegCloseKey, hKey
				ret
			.endif

			mov eax,dwItems
			mov ecx,eax					;start with a 25% reserve here
			shr ecx,2					;so possibly we dont need to resize
			add eax,ecx

			invoke AllocTable, eax

			xor ebx, ebx
			mov dwItem,0
			.while (ebx < dwItems)
				mov dwSize,sizeof szText
				invoke RegEnumKeyEx,hKey,ebx,addr szText,addr dwSize,0,NULL,0,addr filetime
				.break	.if (eax != ERROR_SUCCESS)
				invoke RegOpenKeyEx,hKey,addr szText,0,KEY_READ,addr hGuId
				.if (eax == ERROR_SUCCESS)
					invoke ListTypeLibVersions, hGuId, dwItem, addr szText
					mov dwItem, eax
					invoke RegCloseKey,hGuId
				.endif
				inc ebx
			.endw
			invoke RegCloseKey,hKey
		.endif
		ret
		align 4

RefreshTypeLib endp


;*** list HKEY_CLASSES_ROOT\Interface entries ***


RefreshInterface proc uses ebx

local	hKey:HANDLE
local	dwItems:DWORD
local	hSubKey:HANDLE
local	hSubKey2:HANDLE
local	dwSize:dword
local	filetime:FILETIME
local	iType:dword
local	szText[256]:byte
local	szKey[256]:byte
local	szValue[256]:byte
local	szUserKey[64]:byte
local	szUserValueName[64]:byte

		invoke SplitUserdefinedColumn, offset g_szUserColInterface, addr szUserKey, addr szUserValueName

		invoke RegOpenKeyEx,HKEY_CLASSES_ROOT,m_pszRoot,0,KEY_READ,addr hKey
		.if (eax == ERROR_SUCCESS)

			invoke RegQueryInfoKey,hKey, NULL, 0, NULL,\
				addr dwItems, NULL, NULL, NULL, NULL, NULL, NULL, NULL
			.if (eax != ERROR_SUCCESS)
				invoke wsprintf, addr szText, CStr("RegQueryInfoKey failed [%X]"), eax
				invoke MessageBox, m_hWnd, addr szText, 0, MB_OK
				invoke RegCloseKey, hKey
				ret
			.endif

			invoke AllocTable, dwItems

			xor ebx,ebx
			.while (ebx < dwItems)
				mov dwSize,sizeof szText
				invoke RegEnumKeyEx,hKey,ebx,addr szText,addr dwSize,0,NULL,0,addr filetime
				.break .if (eax != ERROR_SUCCESS)
				invoke SetCurrentItem, ebx
				invoke SetItemString, 0, addr szText, dwSize

				invoke RegOpenKeyEx,hKey, addr szText, 0, KEY_READ, addr hSubKey
				.if (eax == ERROR_SUCCESS)
					mov dwSize,sizeof szValue
					invoke RegQueryValueEx, hSubKey, addr g_szNull, 0, addr iType, addr szValue, addr dwSize
					.if (eax == ERROR_SUCCESS)
						invoke SetItemString, 1, addr szValue, dwSize
					.endif
					invoke RegOpenKeyEx, hSubKey, CStr("ProxyStubClsid32"), 0, KEY_QUERY_VALUE, addr hSubKey2
					.if (eax == ERROR_SUCCESS)
						mov dwSize,sizeof szValue
						invoke RegQueryValueEx, hSubKey2, addr g_szNull, 0, addr iType, addr szValue, addr dwSize
						.if (eax == ERROR_SUCCESS)
							invoke SetItemString, 2, addr szValue, dwSize
						.endif
						invoke RegCloseKey,hSubKey2
					.endif
					invoke RegOpenKeyEx, hSubKey, addr g_szTypeLib, 0, KEY_QUERY_VALUE, addr hSubKey2
					.if (eax == ERROR_SUCCESS)
						mov dwSize,sizeof szValue
						invoke RegQueryValueEx, hSubKey2, addr g_szNull, 0, addr iType, addr szValue, addr dwSize
						.if (eax == ERROR_SUCCESS)
							invoke SetItemString, 3, addr szValue, dwSize
						.endif
						invoke RegCloseKey,hSubKey2
					.endif

;--------------------------------------------- search for userdefined entry
			.if (g_szUserColInterface)
				.if (szUserKey)
					invoke RegOpenKeyEx,hSubKey,addr szUserKey,0,KEY_QUERY_VALUE,addr hSubKey2
				.else
					mov eax,hSubKey
					mov hSubKey2, eax
					mov eax,ERROR_SUCCESS
				.endif
				.if (eax == ERROR_SUCCESS)
					mov dwSize,sizeof szValue - 1
					invoke RegQueryValueEx, hSubKey2, addr szUserValueName, NULL,\
							NULL, addr szValue+1, addr dwSize
					.if (eax != ERROR_SUCCESS)
						mov szValue, 0
						.if (!szUserValueName)
							invoke lstrcpy,addr szValue,CStr("<no default>")
						.endif
					.else
						mov szValue,'"'
						invoke lstrlen,addr szValue
						lea ecx,szValue
						mov word ptr [eax+ecx],'"'
					.endif
					invoke SetItemString, 4, addr szValue, -1
					.if (szUserKey)
						invoke RegCloseKey,hSubKey2
					.endif
				.endif
			.endif


					invoke RegCloseKey,hSubKey
				.endif


				inc ebx
			.endw
			invoke RegCloseKey,hKey
		.endif
		ret
		align 4

RefreshInterface endp


;*** list HKEY_CLASSES_ROOT\AppID entries ***


RefreshAppID proc uses ebx

local	hKey:HANDLE
local	dwItems:DWORD
local	hSubKey:HANDLE
local	szText[256]:byte
local	szKey[256]:byte
local	szValue[256]:byte
local	dwSize:dword
local	filetime:FILETIME
local	iType:dword

		invoke RegOpenKeyEx,HKEY_CLASSES_ROOT, m_pszRoot, 0, KEY_READ, addr hKey
		.if (eax == ERROR_SUCCESS)

			invoke RegQueryInfoKey, hKey, NULL, 0, NULL,\
				addr dwItems, NULL, NULL, NULL, NULL, NULL, NULL, NULL
			.if (eax != ERROR_SUCCESS)
				invoke wsprintf, addr szText, CStr("RegQueryInfoKey failed [%X]"), eax
				invoke MessageBox, m_hWnd, addr szText, 0, MB_OK
				invoke RegCloseKey, hKey
				ret
			.endif

			invoke AllocTable, dwItems

			xor ebx,ebx
			.while (ebx < dwItems)
				mov dwSize,sizeof szKey
				invoke RegEnumKeyEx,hKey,ebx,addr szKey,addr dwSize,0,NULL,0,addr filetime
				.break .if (eax != ERROR_SUCCESS)
				invoke SetCurrentItem, ebx
				invoke SetItemString, 0, addr szKey, dwSize
				invoke RegOpenKeyEx, hKey, addr szKey,0, KEY_READ, addr hSubKey
				.if (eax == ERROR_SUCCESS)
					mov dwSize,sizeof szValue
					invoke RegQueryValueEx,hSubKey,addr g_szNull,NULL, addr iType, addr szValue, addr dwSize
					.if (eax == ERROR_SUCCESS)
						invoke SetItemString, 1, addr szValue, dwSize
					.endif
					mov dwSize,sizeof szValue
					invoke RegQueryValueEx,hSubKey,CStr("AppID"),NULL, addr iType, addr szValue, addr dwSize
					.if (eax == ERROR_SUCCESS)
						invoke SetItemString, 2, addr szValue, dwSize
					.endif
					mov dwSize,4
					invoke RegQueryValueEx,hSubKey,CStr("AuthenticationLevel"),NULL, addr iType, addr szValue, addr dwSize
					.if (eax == ERROR_SUCCESS)
						mov eax, dword ptr szValue
						invoke wsprintf, addr szValue, CStr("%u"), eax
						invoke SetItemString, 3, addr szValue, -1
					.endif
					mov dwSize,sizeof szValue
					invoke RegQueryValueEx,hSubKey,CStr("RunAs"),NULL, addr iType, addr szValue, addr dwSize
					.if (eax == ERROR_SUCCESS)
						invoke SetItemString, 4, addr szValue, -1
					.endif

					mov dwSize,sizeof szValue-3
					invoke RegQueryValueEx,hSubKey,CStr("DllSurrogate"),NULL, addr iType, addr szValue+1, addr dwSize
					.if (eax == ERROR_SUCCESS)
						mov szValue,'"'
						mov eax,dwSize
						lea ecx,szValue
						mov word ptr [eax+ecx],'"'
						invoke SetItemString, 5, addr szValue, -1
					.endif

					mov dwSize,sizeof szValue-3
					invoke RegQueryValueEx,hSubKey,CStr("LocalService"),NULL, addr iType, addr szValue+1, addr dwSize
					.if (eax == ERROR_SUCCESS)
						mov szValue,'"'
						mov eax,dwSize
						lea ecx,szValue
						mov word ptr [eax+ecx],'"'
						invoke SetItemString, 6, addr szValue, -1
					.endif

					invoke RegCloseKey, hSubKey
				.endif
				inc ebx
			.endw
			invoke RegCloseKey,hKey
		.endif
		ret
		align 4

RefreshAppID endp


;*** list HKEY_CLASSES_ROOT\Component Categories entries ***


RefreshCompCat proc uses ebx

local	hKey:HANDLE
local	dwItems:DWORD
local	hSubKey:HANDLE
local	szText[256]:byte
local	szKey[256]:byte
local	szValue[256]:byte
local	szValueName[256]:byte
local	dwSize:dword
local	dwSize2:dword
local	dwType:dword
local	filetime:FILETIME
local	iType:dword

		invoke RegOpenKeyEx, HKEY_CLASSES_ROOT, m_pszRoot, 0, KEY_READ, addr hKey
		.if (eax == ERROR_SUCCESS)

			invoke RegQueryInfoKey,hKey, NULL, 0, NULL,\
				addr dwItems, NULL, NULL, NULL, NULL, NULL, NULL, NULL
			.if (eax != ERROR_SUCCESS)
				invoke wsprintf, addr szText, CStr("RegQueryInfoKey failed [%X]"), eax
				invoke MessageBox, m_hWnd, addr szText, 0, MB_OK
				invoke RegCloseKey, hKey
				ret
			.endif

			invoke AllocTable, dwItems

			xor ebx,ebx
			.while (ebx < dwItems)
				mov dwSize,sizeof szKey
				invoke RegEnumKeyEx, hKey, ebx, addr szKey, addr dwSize, 0, NULL, 0, addr filetime
				.break .if (eax != ERROR_SUCCESS)
				invoke SetCurrentItem, ebx
				invoke SetItemString, 0, addr szKey, dwSize
				invoke RegOpenKeyEx,hKey,addr szKey,NULL,KEY_READ,addr hSubKey
				.if (eax == ERROR_SUCCESS)
					mov dwSize,sizeof szValue
					mov dwSize2,sizeof szValueName
					invoke RegEnumValue,hSubKey,0,addr szValueName,addr dwSize2,\
								NULL,addr dwType,addr szValue,addr dwSize
					.if (eax == ERROR_SUCCESS)
						invoke SetItemString, 1, addr szValue, dwSize
					.endif
					invoke RegCloseKey,hSubKey
				.endif
				inc ebx
			.endw
			invoke RegCloseKey,hKey
		.endif
		ret
		align 4

RefreshCompCat endp


;*** list HKEY_CLASSES_ROOT entries ***


RefreshHKCR proc uses ebx esi

local	hKey:HANDLE
local	dwItems:DWORD
local	hSubKey:HANDLE
local	hSubKey2:HANDLE
local	dwSize:dword
local	dwSize2:dword
local	dwType:dword
local	filetime:FILETIME
local	iType:dword
local	szKey[256]:byte
local	szValue[256]:byte
local	szText[256]:byte
local	szUserKey[64]:byte
local	szUserValueName[64]:byte

		invoke SplitUserdefinedColumn, offset g_szUserColHKCR, addr szUserKey, addr szUserValueName

		invoke RegOpenKeyEx, HKEY_CLASSES_ROOT, m_pszRoot, 0, KEY_READ, addr hKey
		.if (eax == ERROR_SUCCESS)

			invoke RegQueryInfoKey, hKey, NULL, 0, NULL,\
				addr dwItems, NULL, NULL, NULL, NULL, NULL, NULL, NULL
			.if (eax != ERROR_SUCCESS)
				invoke wsprintf, addr szText, CStr("RegQueryInfoKey failed [%X]"), eax
				invoke MessageBox, m_hWnd, addr szText, 0, MB_OK
				invoke RegCloseKey, hKey
				ret
			.endif

			invoke AllocTable, dwItems

			xor ebx,ebx
			.while (ebx < dwItems)
				mov dwSize,sizeof szKey
				invoke RegEnumKeyEx,hKey,ebx,addr szKey,addr dwSize,0,NULL,0,addr filetime
				.break .if (eax != ERROR_SUCCESS)
				invoke SetCurrentItem, ebx
				invoke SetItemString, 0, addr szKey, dwSize
				invoke RegOpenKeyEx,hKey,addr szKey,NULL,KEY_READ,addr hSubKey
				.if (eax == ERROR_SUCCESS)
					mov dwSize,sizeof szValue
					invoke RegQueryValueEx,hSubKey,addr g_szNull,0,addr iType,addr szValue,addr dwSize
					.if (eax == ERROR_SUCCESS)
						invoke SetItemString, 1, addr szValue, dwSize
					.endif
					invoke RegOpenKeyEx,hSubKey, addr g_szRootCLSID, NULL, KEY_READ, addr hSubKey2
					.if (eax == ERROR_SUCCESS)
						mov dwSize,sizeof szValue
						invoke RegQueryValueEx,hSubKey2,addr g_szNull,0,addr iType,addr szValue,addr dwSize
						.if (eax == ERROR_SUCCESS)
							invoke SetItemString, 2, addr szValue, dwSize
						.endif
						invoke RegCloseKey,hSubKey2
					.endif
					invoke RegOpenKeyEx,hSubKey,CStr("Shell\Open\Command"),NULL,KEY_READ,addr hSubKey2
					.if (eax == ERROR_SUCCESS)
						mov dwSize,sizeof szValue
						invoke RegQueryValueEx,hSubKey2,addr g_szNull,0,addr iType,addr szValue,addr dwSize
						.if (eax == ERROR_SUCCESS)
							invoke SetItemString, 3, addr szValue, dwSize
						.endif
						invoke RegCloseKey,hSubKey2
					.endif

;--------------------------------------------- search for userdefined entry
					.if (g_szUserColHKCR)
						.if (szUserKey)
							invoke RegOpenKeyEx,hSubKey,addr szUserKey,0,KEY_READ,addr hSubKey2
						.else
							mov eax,hSubKey
							mov hSubKey2, eax
							mov eax,ERROR_SUCCESS
						.endif
						.if (eax == ERROR_SUCCESS)
							mov dwSize,sizeof szValue - 1
							invoke RegQueryValueEx, hSubKey2, addr szUserValueName,NULL,\
									addr dwType, addr szValue+1, addr dwSize
							.if (eax != ERROR_SUCCESS)
								.if (szUserValueName)
									mov szValue, 0
								.else
									invoke lstrcpy, addr szValue, CStr("<no default>")
								.endif
							.elseif (dwType == REG_DWORD)
								mov eax, dword ptr szValue+1
								invoke wsprintf, addr szValue, CStr("%08X"), eax
							.elseif (dwType == REG_BINARY)
								mov szText,0
								mov ecx,dwSize
								.if (ecx > 8)
									mov ecx,8
								.endif
								lea esi,szValue+1
								push edi
								lea edi, szText
								.while (ecx)
									push ecx
									lodsb
									movzx eax,al
									invoke wsprintf, edi, CStr("%02X "), eax
									add edi, eax
									pop ecx
									dec ecx
								.endw
								pop edi
								invoke lstrcpy, addr szValue, addr szText
							.else
								mov szValue,'"'
								invoke lstrlen,addr szValue
								lea ecx,szValue
								mov word ptr [eax+ecx],'"'
							.endif
							invoke SetItemString, 4, addr szValue, -1
							.if (szUserKey)
								invoke RegCloseKey,hSubKey2
							.endif
						.endif
					.endif

					invoke RegCloseKey,hSubKey
				.endif
				inc ebx
			.endw
			invoke RegCloseKey,hKey
		.endif
		ret
		align 4

RefreshHKCR endp


;*** list created objects
;*** these objects are saved in a linked list
;*** save ptr to object in itemdata


RefreshObject proc uses ebx

local	guid:GUID
local	wszGUID[40]:word
local	dwNumItems:DWORD
local	dwItem:DWORD
local	pMoniker:DWORD
local	pBindCtx:LPBINDCTX
local	pwszName:LPOLESTR
local	szText[MAX_PATH]:byte

		invoke GetItemCount@CList, g_pObjects
		mov dwNumItems,eax

		invoke CreateBindCtx, NULL, addr pBindCtx

		invoke AllocTable, dwNumItems

		mov dwItem, 0
		.while (1)

;---------------------------------------------- get next created object

			invoke GetItem@CList, g_pObjects, dwItem
			.break .if (!eax)
			mov ebx, eax
			invoke GetGUID@CObjectItem, ebx, addr guid
			invoke StringFromGUID2, addr guid, addr wszGUID, LENGTHOF wszGUID
			invoke WideCharToMultiByte,CP_ACP,0, addr wszGUID, -1, addr szText, 80,0,0

			invoke SetCurrentItem, dwItem
			invoke SetItemString, 0, addr szText, 40
			invoke GetTextFromCLSID, addr guid, addr szText, sizeof szText
			invoke SetItemString, 1, addr szText, -1

            invoke vf(ebx, IObjectItem, GetMoniker)
			.if (eax)
				mov pMoniker, eax
				invoke vf(pMoniker, IMoniker, GetDisplayName), pBindCtx, NULL, addr pwszName
				.if (eax == S_OK)
					invoke WideCharToMultiByte,CP_ACP,0, pwszName, -1, addr szText, sizeof szText,0,0 
					invoke SetItemString, 2, addr szText, -1
					invoke CoTaskMemFree, pwszName
				.endif
			.endif

			invoke vf(ebx, IObjectItem, GetConnectionList), FALSE
			.if (eax)
				invoke GetItemCount@CList, eax
				invoke wsprintf, addr szText, CStr("%u"), eax
				invoke SetItemString, 3, addr szText, eax
			.endif

			invoke SetItemData, DATACOL_IN_OBJECT, ebx

			inc dwItem
		.endw
		invoke vf(pBindCtx, IUnknown, Release)
		ret
		align 4

RefreshObject endp


RefreshROT proc uses ebx

local	pROT:LPRUNNINGOBJECTTABLE
local	pEnumMoniker:LPENUMMONIKER
local	pMoniker:LPMONIKER
local	pBC:LPBC
local	pOleStr:LPSTR
local	hKey:HANDLE
local	dwHash:DWORD
local	pUnknown:LPUNKNOWN
local	dwSize:DWORD
local	dwType:DWORD
local	clsid:GUID
local	wszCLSID[40]:word
local	szText[128]:byte
local	szKey[128]:byte

		invoke AllocTable, 64

		mov pBC, NULL
		invoke CreateBindCtx, NULL, addr pBC

		invoke GetRunningObjectTable, NULL, addr pROT
		.if (eax == S_OK)
			invoke vf(pROT, IRunningObjectTable, EnumRunning), addr pEnumMoniker
			.if (eax == S_OK)
				xor ebx, ebx
				.while (1)
					invoke vf(pEnumMoniker, IEnumMoniker, Next), 1, addr pMoniker, NULL
					.break .if (eax != S_OK)
					invoke SetCurrentItem, ebx
					invoke vf(pMoniker, IMoniker, IsSystemMoniker), addr dwType
					.if (eax == S_OK)
						mov eax, dwType
						.if (eax == MKSYS_NONE)
							mov eax, CStr("None")
						.elseif (eax == MKSYS_GENERICCOMPOSITE)
							mov eax, CStr("Generic Composite Moniker")
						.elseif (eax == MKSYS_FILEMONIKER)
							mov eax, CStr("File Moniker")
						.elseif (eax == MKSYS_ANTIMONIKER)
							mov eax, CStr("Anti-Moniker")
						.elseif (eax == MKSYS_ITEMMONIKER)
							mov eax, CStr("Item Moniker")
						.elseif (eax == MKSYS_POINTERMONIKER)
							mov eax, CStr("Pointer Moniker")
						.elseif (eax == MKSYS_CLASSMONIKER)
							mov eax, CStr("Class Moniker")
						.elseif (eax == MKSYS_OBJREFMONIKER)
							mov eax, CStr("Object Reference Moniker")
						.elseif (eax == MKSYS_SESSIONMONIKER)
							mov eax, CStr("Session Moniker")
						.else
							mov eax, CStr("???")
						.endif
						invoke SetItemString, 2, eax, -1
					.endif
if ?USEHASH
					invoke vf(pMoniker, IMoniker, Hash), addr dwHash
					invoke SetItemData, 5, dwHash
					DebugOut "RefreshROT, pMoniker=%X, dwHash=%X", pMoniker, dwHash
else
					invoke vf(pROT, IRunningObjectTable, GetObject_), pMoniker, addr pUnknown
					.if (eax == S_OK)
						invoke SetItemData, 5, pUnknown
						invoke vf(pUnknown, IUnknown, Release)
						DebugOut "RefreshROT, pMoniker=%X, pUnknown=%X", pMoniker, pUnknown
					.endif
endif
					invoke vf(pMoniker, IMoniker, IsRunning), pBC, NULL, NULL
					.if (eax == S_OK)
						invoke SetItemString, 3, CStr("yes"), -1
					.endif
if 0
					invoke vf(pMoniker, IMoniker, GetClassID), addr clsid
					.if (eax == S_OK)
						invoke StringFromGUID2, addr clsid, addr wszCLSID, 40
						invoke WideCharToMultiByte,CP_ACP,0,addr wszCLSID,-1,addr szText, sizeof szText,0,0 
						invoke SetItemString, 0, addr szText, -1
						invoke GetTextFromCLSID, addr clsid, addr szText, sizeof szText
						invoke SetItemString, 4, addr szText, -1
					.endif
endif
					invoke vf(pMoniker, IMoniker, GetDisplayName), pBC, NULL, addr pOleStr
					.if (eax == S_OK)
						invoke WideCharToMultiByte,CP_ACP,0,pOleStr,-1,addr szText, sizeof szText,0,0 
						invoke SetItemString, 1, addr szText, -1
if 1
						.if (word ptr szText == "{!")
							mov edx, pOleStr
							inc edx
							inc edx
							invoke CLSIDFromString, edx, addr clsid
							.if (eax == S_OK)
								invoke SetItemString, 0, addr szText+1, -1
								invoke GetTextFromCLSID, addr clsid, addr szText, sizeof szText
								invoke SetItemString, 4, addr szText, -1
							.endif
						.endif
endif
					.endif
					.if (pOleStr)
						invoke CoTaskMemFree, pOleStr
					.endif
					invoke vf(pMoniker, IUnknown, Release)
					inc ebx
				.endw
				invoke vf(pEnumMoniker, IUnknown, Release)
			.endif
			invoke vf(pROT, IUnknown, Release)
		.endif
		.if (pBC)
			invoke vf(pBC, IUnknown, Release)
		.endif
		ret
		align 4

RefreshROT endp

FindROTItem@CDocument proc public uses __this this_:ptr CDocument, dwIndex:DWORD

local	pROT:LPRUNNINGOBJECTTABLE
local	pUnknown:LPUNKNOWN
local	pEnumMoniker:LPENUMMONIKER
local	pMoniker:LPMONIKER
local	dwHash:DWORD
local	dwSearchHash:DWORD
local	bFound:BOOLEAN

		mov __this, this_
		mov bFound, FALSE
;----------------------------------- get hash
		invoke GetItemData@CDocument, __this, dwIndex, 5
		mov dwSearchHash, eax
		invoke GetRunningObjectTable, NULL, addr pROT
		.if (eax == S_OK)
			invoke vf(pROT, IRunningObjectTable, EnumRunning), addr pEnumMoniker
			.if (eax == S_OK)
				.while (1)
					invoke vf(pEnumMoniker, IEnumMoniker, Next), 1, addr pMoniker, NULL
					.break .if (eax != S_OK)
if ?USEHASH
					invoke vf(pMoniker, IMoniker, Hash), addr dwHash
					mov ecx, dwHash
					.if ((eax == S_OK) && (ecx == dwSearchHash))
						mov bFound, TRUE
						.break
					.endif
else
					invoke vf(pROT, IRunningObjectTable, GetObject_), pMoniker, addr pUnknown
					.if (eax == S_OK)
						invoke vf(pUnknown, IUnknown, Release)
						mov ecx, dwSearchHash
						.if (ecx == pUnknown)
							mov bFound, TRUE
							.break
						.endif
					.endif
endif
					invoke vf(pMoniker, IUnknown, Release)
				.endw
				invoke vf(pEnumMoniker, IUnknown, Release)
			.endif
			invoke vf(pROT, IUnknown, Release)
		.endif
		xor eax, eax
		.if (bFound)
			mov eax, pMoniker
		.endif
		ret
		align 4

FindROTItem@CDocument endp

;--- list may grow dynamically, resize item table


ResizeTable proc uses esi

local	dwReserve:dword

		mov eax, m_numItems
		shr eax,2						;increase document 25%
		mov dwReserve,eax
		add eax, m_numItems
		mul m_ItemSize
		invoke malloc, eax
		.if (!eax)
			invoke MemoryError, m_hWnd
			ret
		.endif
		mov esi, m_pItems		;copy document to new area
		mov m_pItems, eax
		mov eax, m_numItems
		mul m_ItemSize
		shr eax,2
		mov ecx, eax			;memory size to copy
		push edi
		push esi
		mov edi, m_pItems
		rep movsd
		pop esi
		pop edi
		invoke free, esi
		mov eax, dwReserve
		mov m_numFree, eax
		ret
		align 4

ResizeTable endp


;*** create item table + string pool


AllocTable proc dwItems:DWORD

		mov eax, m_ItemSize
		mul dwItems
		invoke malloc, eax
		mov m_pItems,eax
		.if (!eax)
			invoke MemoryError, m_hWnd
			ret
		.endif

		mov eax,dwItems
		mov m_numFree,eax

;----------------------------------- now alloc first block of string pool
		mul m_numStrng
		mov ecx,32					;assumed string size average
		mul ecx
		and ax,0F000h
		add eax,1000h
		invoke AllocStringPoolBlock, eax
		ret
		align 4

AllocTable endp


;*** create document (list of items)
;*** on memory errors just terminate app


Create@CDocument proc public uses __this  hWnd:HWND, iMode:DWORD,
				iNumStrings:DWORD, pszRoot:LPSTR

		invoke malloc, sizeof CDocument
		.if (!eax)
			invoke MemoryError, hWnd
			ret
		.endif
		mov __this,eax
	
		mov m_numItems,0
		mov eax,pszRoot
		mov m_pszRoot, eax
		mov eax,hWnd
		mov m_hWnd,eax
		mov eax,iNumStrings
		mov m_numStrng,eax

;---------------------------------------- calc item size
		mov eax,iNumStrings
		inc eax							; 1 extra dword for item data
		.if ((iMode == MODE_ROT) || (iMode == MODE_OBJECT))
			inc eax						; 1 extra data for ROT (hash)
		.endif
		shl eax,2
		mov m_ItemSize,eax				; size of each item in document

		mov ecx,NUMMODES
		mov edx,offset ModeTab
		mov eax,iMode
		.while (ecx)
			.break .if (eax == [edx]._CMode.iMode)
			add edx,2*4
			dec ecx
		.endw

		invoke [edx]._CMode.pRefresh

		mov eax,__this
		ret
		align 4

Create@CDocument endp


;-------------------------- delete document


Destroy@CDocument proc public uses __this thisarg

		mov __this,this@

		xor eax,eax
		xchg eax, m_pStrPool
		.while (eax)
			push [eax]
			invoke VirtualFree, eax, NULL, MEM_RELEASE
			pop eax
		.endw

		xor eax,eax
		xchg eax, m_pItems
		.if (eax)
			invoke free, eax
		.endif

		invoke free, __this

		ret
		align 4

Destroy@CDocument endp

	end
