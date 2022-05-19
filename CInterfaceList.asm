
	.386
	.model flat,stdcall
	option casemap :none   ; case sensitive
	option proc:private

	include COMView.inc
	include classes.inc
	include rsrc.inc

;*** creates a list of GUIDs (used for IIDs only hence the name CInterfaceList)

	.code

__this	textequ <ebx>
_this	textequ <[__this].CInterfaceList>
thisarg	textequ <this@:ptr CInterfaceList>

	MEMBER pRegIIDs, cntIID, cntRegIID, ppszNames, pszStrings

Create@CInterfaceList proc public uses __this hWnd:HWND

		invoke malloc, sizeof CInterfaceList
		.if (eax)
			mov __this,eax
			invoke ReadAllRegSubKeys, hWnd, addr g_szRootInterface, addr m_cntIID
			.if (eax == 0)
				invoke free,__this
				xor eax,eax
			.else
				mov m_pRegIIDs,eax
				mov eax,__this
			.endif
		.endif
		ret
		align 4

Create@CInterfaceList endp


Destroy@CInterfaceList proc public uses __this thisarg

		mov __this,this@

		invoke free, m_pRegIIDs
		.if (m_ppszNames)
			invoke free, m_ppszNames
		.endif
		.if (m_pszStrings)
			invoke free, m_pszStrings
		.endif
		invoke free, __this
		ret
		align 4

Destroy@CInterfaceList endp

Find@CInterfaceList proc public uses esi edi __this thisarg, riid:REFIID

		mov __this,this@
		mov ecx,m_cntIID
		mov esi,riid
		mov edi, m_pRegIIDs
		.while (ecx)
			push ecx
			push esi
			push edi
			mov ecx,4
			repz cmpsd
			pop edi
			pop esi
			pop ecx
			.break .if (ZERO?)
			add edi, sizeof IID
			dec ecx
		.endw
		.if (!ecx)
			xor eax, eax
		.else
			mov eax, edi
		.endif
		ret
		align 4

Find@CInterfaceList endp

AddIIDs@CInterfaceList proc public uses __this edi esi thisarg, dwCount:DWORD, pIIDs:ptr IID, pszNames:LPSTR

		mov __this,this@
		mov eax, m_cntIID
		add eax, dwCount
		shl eax, 4
		invoke malloc, eax
		.if (eax)
;------------------------------------------------ add new IIDs
			mov edi, eax
			mov ecx, m_cntIID
			shl ecx, 4
			push ecx
			invoke CopyMemory, edi, m_pRegIIDs, ecx
			invoke free, m_pRegIIDs
			mov m_pRegIIDs, edi
			pop ecx
			add edi, ecx
			mov eax, dwCount
			mov m_cntRegIID, eax
			add m_cntIID, eax
			shl eax, 4
			mov ecx, eax
			invoke CopyMemory, edi, pIIDs, ecx

;------------------------------------------------ alloc string pointers
			mov eax, dwCount
			shl eax, 2
			invoke malloc, eax
			mov m_ppszNames, eax
			mov edi, eax
;------------------------------------------------ copy strings
			mov esi, pszNames
			mov ecx, dwCount
			.while (ecx)
				mov eax, esi
				sub eax, pszNames
				stosd
				.while (1)
					lodsb
					.break .if (!al)
				.endw
				dec ecx
			.endw
			sub esi, pszNames
			invoke malloc, esi
			mov m_pszStrings, eax
			invoke CopyMemory, m_pszStrings, pszNames, esi
		.endif
		ret
AddIIDs@CInterfaceList endp

GetName@CInterfaceList proc public uses __this thisarg, dwIndex:DWORD

		mov __this,this@
		mov ecx, dwIndex
		mov eax, m_cntIID
		sub eax, m_cntRegIID
		mov edx, m_ppszNames
		.if (edx && (ecx >= eax) && (ecx < m_cntIID))
			sub ecx, eax
			mov eax, [edx+ecx*4]
			add eax, m_pszStrings
		.else
			xor eax, eax
		.endif
		ret
		align 4

GetName@CInterfaceList endp

	end
