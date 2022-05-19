
	.386
	.model flat,stdcall
	option casemap :none   ; case sensitive
	option proc:private

INSIDE_CDROPTARGET equ 1

	include COMView.inc
	include shellapi.inc
	include classes.inc
	include debugout.inc
	include rsrc.inc

;*** CDropTarget methods

CDropTarget struct
_IDropTarget	IDropTarget <>
dwRefCount		DWORD ?
hWndMain		HWND ?
CDropTarget ends


	.const

CDropTargetVtbl label dword
	IUnknownVtbl {QueryInterface, AddRef, Release}
    dd DragEnter, DragOver, DragLeave, Drop

iftab label dword
	dd IID_IUnknown				, 0
	dd IID_IDropTarget			, CDropTarget._IDropTarget
NUMIFENTRIES textequ %($ - offset iftab) / (4 * 2)

	.data

g_bRButton BOOLEAN FALSE

	.code

__this	textequ <ebx>
_this	textequ <[__this].CDropTarget>

;;IsEqualIID equ <IsEqualGUID>

	MEMBER _IDropTarget, dwRefCount, hWndMain

Create@CDropTarget proc public uses __this hWndMain:HWND

	invoke malloc, sizeof CDropTarget
	mov __this, eax
	mov m__IDropTarget.lpVtbl, offset CDropTargetVtbl
	@mov m_hWndMain, hWndMain
	mov m_dwRefCount,1
	return __this
	align 4

Create@CDropTarget endp

Destroy@CDropTarget proc this_:ptr CDropTarget

	invoke free, this_
	ret
	align 4

Destroy@CDropTarget endp


AddRef proto :ptr CDropTarget

QueryInterface proc this_:ptr CDropTarget, riid:REFIID , ppReturn:ptr LPUNKNOWN

	invoke IsInterfaceSupported, riid, offset iftab, NUMIFENTRIES,  this_, ppReturn
	ret
	align 4

QueryInterface endp

AddRef proc this_:ptr CDropTarget
	mov ecx,this_
	inc [ecx].CDropTarget.dwRefCount
	return [ecx].CDropTarget.dwRefCount
	align 4
AddRef endp

Release proc uses __this this_:ptr CDropTarget
	mov __this,this_
	dec m_dwRefCount
	.if (!m_dwRefCount)
		invoke Destroy@CDropTarget, __this
		return 0
	.endif
	return m_dwRefCount
	align 4
Release endp

DragEnter proc uses __this this_:ptr CDropTarget, pDataObj:LPDATAOBJECT, grfKeyState:DWORD, pt:POINTL , pdwEffect:ptr DWORD

local fe:FORMATETC

	mov __this, this_
	mov fe.cfFormat,CF_HDROP
	mov fe.ptd,NULL
	mov fe.dwAspect,DVASPECT_CONTENT
	mov fe.lindex,-1
	mov fe.tymed,TYMED_HGLOBAL
	invoke vf(pDataObj,IDataObject, QueryGetData), addr fe
	mov ecx, pdwEffect
	.if ((eax != S_OK) || (g_bAcceptDrop == FALSE))
		mov dword ptr [ecx], DROPEFFECT_NONE
	.else
		mov dword ptr [ecx], DROPEFFECT_COPY
	.endif
	DebugOut "CDropTarget::DragEnter, QueryGetData()=%X", eax
	.if (grfKeyState & MK_RBUTTON)
		mov g_bRButton, TRUE
	.else
		mov g_bRButton, FALSE
	.endif
	return S_OK
	align 4

DragEnter endp

DragOver proc uses __this this_:ptr CDropTarget, grfKeyState:DWORD, pt:POINTL, pdwEffect:ptr DWORD

	DebugOut "CDropTarget::DragOver"
	mov ecx, pdwEffect
	.if (g_bAcceptDrop == FALSE)
		mov dword ptr [ecx], DROPEFFECT_NONE
	.else
		mov dword ptr [ecx], DROPEFFECT_COPY
	.endif
	.if (grfKeyState & MK_RBUTTON)
		mov g_bRButton, TRUE
	.else
		mov g_bRButton, FALSE
	.endif
	return S_OK
	align 4

DragOver endp

DragLeave proc uses __this this_:ptr CDropTarget
	DebugOut "CDropTarget::DragLeave"
	return S_OK
	align 4
DragLeave endp

	.const

MenuDescTab label dword
	dd IDM_LOADFILE,	CStr("&Bind to File")
	dd IDM_CREATELINK,	CStr("&Create Link to File")
	dd -1, 0
	dd IDM_LOADTYPELIB, CStr("&Load Type Library")
	dd IDM_REGISTER,	CStr("&Register Server")
	dd IDM_UNREGISTER,	CStr("&Unregister Server")
	dd -1, 0
	dd IDM_OPENSTORAGE, CStr("&Open Storage")
	dd IDM_OPENSTREAM,	CStr("Open &Stream")
	dd -1, 0
	dd IDCANCEL, CStr("&Cancel")
NUMMENUENTRIES equ ($ - offset MenuDescTab) / (sizeof DWORD * 2)	

	.code

ShowContextMenu proc uses esi grfKeyState:DWORD, dwEffect:DWORD

local pt:POINT

	invoke CreatePopupMenu
	mov esi, eax
	mov edx, offset MenuDescTab
	mov ecx, NUMMENUENTRIES
	.while (ecx)
		push ecx
		push edx
		.if (dword ptr [edx] != -1)
			invoke AppendMenu, esi, MF_STRING, [edx+0], [edx+4]
		.else
			invoke AppendMenu, esi, MF_SEPARATOR, -1, 0
		.endif
		pop edx
		pop ecx
		add edx, 8
		dec ecx
	.endw

	.if (g_bBindIsDefault)
		mov ecx, IDM_LOADFILE
	.else
		mov ecx, IDM_LOADTYPELIB
	.endif
	invoke SetMenuDefaultItem, esi, ecx, FALSE
	invoke GetCursorPos, addr pt
	invoke GetFocus
	invoke TrackPopupMenu, esi,\
		TPM_LEFTALIGN or TPM_LEFTBUTTON or TPM_RETURNCMD,\
		pt.x, pt.y, 0, m_hWndMain, NULL
	push eax
	invoke DestroyMenu, esi
	pop eax
	ret
	align 4

ShowContextMenu endp



Drop proc uses esi __this this_:ptr CDropTarget, pDataObj:LPDATAOBJECT, grfKeyState:DWORD, pt:POINTL, pdwEffect:ptr DWORD

local dwNumFiles:DWORD
local dwCmd:DWORD
local fe:FORMATETC
local stgmedium:STGMEDIUM
local szPath[MAX_PATH]:SBYTE

	DebugOut "CDropTarget::Drop"
	mov __this, this_
	mov fe.cfFormat,CF_HDROP
	mov fe.ptd,NULL
	mov fe.dwAspect,DVASPECT_CONTENT
	mov fe.lindex,-1
	mov fe.tymed,TYMED_HGLOBAL
	invoke vf(pDataObj, IDataObject, GetData), addr fe, addr stgmedium
	.if (eax != S_OK)
		invoke SetErrorText@CMainDlg, g_pMainDlg, CStr("IDataObject::GetData failed [%X]"), eax, TRUE
		return E_FAIL
	.endif
	invoke DragQueryFile, stgmedium.hGlobal, -1, NULL, 0
	mov dwNumFiles, eax

	.if (g_bBindIsDefault)
		mov dwCmd, 0
	.else
		mov dwCmd, IDM_LOADTYPELIB
	.endif
	mov edx, grfKeyState
	.if (g_bRButton)
		mov ecx, pdwEffect
		invoke ShowContextMenu, edx, [ecx]
		.if (!eax || (eax == IDCANCEL))
			jmp exit
		.endif
		mov dwCmd, eax
	.endif

	xor esi, esi
	.while (esi < dwNumFiles)
		invoke DragQueryFile, stgmedium.hGlobal, esi, addr szPath, MAX_PATH
		.if (eax)
			mov ecx, g_pMainDlg
			.if (dwCmd == 0)
				invoke SmartLoad@CMainDlg, ecx, addr szPath
			.elseif (dwCmd == IDM_LOADTYPELIB)
				invoke Create2@CTypeLibDlg, addr szPath, NULL, FALSE
				invoke Show@CTypeLibDlg, eax, m_hWndMain, FALSE
			.elseif (dwCmd == IDM_REGISTER)
				invoke Register@CMainDlg, ecx, addr szPath
			.elseif (dwCmd == IDM_UNREGISTER)
				invoke FileOperation@CMainDlg, ecx, FILEOP_UNREGISTER, addr szPath
			.elseif (dwCmd == IDM_LOADFILE)
				invoke LoadFile@CMainDlg, ecx, addr szPath, TRUE
			.elseif (dwCmd == IDM_CREATELINK)
				invoke CreateLink@CMainDlg, ecx, addr szPath
			.elseif (dwCmd == IDM_OPENSTORAGE)
				invoke OpenStorage@CMainDlg, ecx, addr szPath, NULL
			.elseif (dwCmd == IDM_OPENSTREAM)
				invoke OpenStream@CMainDlg, ecx, addr szPath
			.endif
		.endif
		inc esi
	.endw
exit:
	invoke ReleaseStgMedium, addr stgmedium
	return S_OK
	align 4

Drop endp

	end
