
	.386
	.model flat,stdcall
	option casemap :none   ; case sensitive

	include COMView.inc
	include classes.inc
	include rsrc.inc

__this	textequ <ebx>
_this	textequ <[__this].CInterfaceItem>
thisarg	textequ <this@:ptr CInterfaceItem>

	MEMBER iid, TypelibGUID, dwVerMajor, dwVerMinor, pszName

	.code

Create@CInterfaceItem proc uses __this refiid:ptr IID

local	hKey:HANDLE
local	hSubKey:HANDLE
local	dwSize:dword
local	iType:dword
local	wszIID[40]:word
local	szKey[128]:byte
local	szTypeLib[64]:byte
local	szVersion[64]:byte

		invoke malloc, sizeof CInterfaceItem
		.if (eax)
			mov __this,eax
			mov m_dwVerMajor,-1
			mov m_dwVerMinor,-1
			invoke CopyMemory, addr m_iid, refiid, sizeof IID
			invoke StringFromGUID2, refiid, addr wszIID, LENGTHOF wszIID
		
			invoke wsprintf, addr szKey, CStr("%s\%S"), addr g_szInterface, addr wszIID
			invoke RegOpenKeyEx, HKEY_CLASSES_ROOT, addr szKey, 0, KEY_READ, addr hKey
			.if (eax == ERROR_SUCCESS)
				invoke RegQueryValueEx, hKey, addr g_szNull, NULL, NULL, NULL, addr dwSize
				.if (eax == ERROR_SUCCESS)
					invoke malloc, dwSize
					mov m_pszName, eax
					invoke RegQueryValueEx, hKey, addr g_szNull, NULL, NULL, m_pszName, addr dwSize
				.endif
				invoke RegOpenKeyEx, hKey, addr g_szTypeLib, 0, KEY_READ, addr hSubKey
				.if (eax == ERROR_SUCCESS)
					mov dwSize,sizeof szTypeLib
					invoke RegQueryValueEx, hSubKey, addr g_szNull, NULL, addr iType, addr szTypeLib, addr dwSize
					.if (szTypeLib)
						invoke MultiByteToWideChar, CP_ACP, MB_PRECOMPOSED,\
							addr szTypeLib, -1, addr wszIID, LENGTHOF wszIID 
						invoke IIDFromString, addr wszIID, addr m_TypelibGUID
						mov m_dwVerMajor,0
						mov m_dwVerMinor,0
						mov dwSize,sizeof szVersion
						invoke RegQueryValueEx, hSubKey,CStr("Version"), NULL, addr iType, addr szVersion, addr dwSize
						.if (szVersion)
							invoke String22DWords, addr szVersion, addr m_dwVerMajor,addr m_dwVerMinor
						.endif
					.endif
					invoke RegCloseKey,hSubKey
				.endif
				invoke RegCloseKey,hKey
			.endif
			mov eax, __this
		.endif
		ret
		align 4

Create@CInterfaceItem endp


SetName@CInterfaceItem proc uses __this thisarg, pszName:LPSTR

		mov __this,this@

		invoke free, m_pszName
		mov m_pszName, NULL
		.if (pszName)
			invoke lstrlen, pszName
			inc eax
			invoke malloc, eax
			mov m_pszName, eax
			invoke lstrcpy, m_pszName, pszName
		.endif
		ret
		align 4

SetName@CInterfaceItem endp

SetTypeLibAttr@CInterfaceItem proc uses __this thisarg, pTLibAttr:ptr TLIBATTR
		mov __this,this@

		mov ecx, pTLibAttr
		movzx eax, [ecx].TLIBATTR.wMajorVerNum
		mov m_dwVerMajor, eax
		movzx eax, [ecx].TLIBATTR.wMinorVerNum
		mov m_dwVerMinor, eax
		invoke CopyMemory, addr m_TypelibGUID, addr [ecx].TLIBATTR.guid, sizeof GUID
		ret
		align 4
SetTypeLibAttr@CInterfaceItem endp

Destroy@CInterfaceItem proc uses __this thisarg

		mov __this,this@

		invoke free, m_pszName
		invoke free,__this
		ret
		align 4

Destroy@CInterfaceItem endp

	end
