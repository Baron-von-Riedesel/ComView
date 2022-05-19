
	.386
	.model flat,stdcall
	option casemap:none
	option proc:private

;--- a CDataObject is required by DoDragDrop (COMView acting as drop source)

INSIDE_CDATAOBJECT equ 1
	include COMView.inc
	include classes.inc
	include debugout.inc

;--- declare the object structure

CDataObject struct
_IDataObject	IDataObject  <>
dwRefCount	dd	?
pData		dd	?
lenData 	dd	?
pFmtEtc		dd	?
NumFmtEtc	dd	?
myFormat	dd	?
CDataObject ends
;-----------------------------------------------------------------------------

QueryInterface proto :ptr CDataObject, :REFIID, :ptr LPUNKNOWN
AddRef	proto :ptr CDataObject
Release proto :ptr CDataObject

__this	textequ <ebx>
_this	textequ <[__this].CDataObject>

	MEMBER _IDataObject, dwRefCount, pData, lenData, pFmtEtc, NumFmtEtc, myFormat

	.data

externdef stdcall IID_IDataObject:GUID

	.const

; define the vtable

vtblIDataObject label dword
	IUnknownVtbl {QueryInterface, AddRef, Release}
	dd offset GetData
	dd offset GetDataHere
	dd offset QueryGetData
	dd offset GetCanonicalFormatEtc
	dd offset SetData
	dd offset EnumFormatEtc
	dd offset DAdvise
	dd offset DUnadvise
	dd offset EnumDAdvise

iftab label dword
	dd IID_IUnknown				, 0
	dd IID_IDataObject			, CDataObject._IDataObject
NUMIFENTRIES textequ %($ - offset iftab) / (4 * 2)

	.code

;--- Constructor

Create@CDataObject proc public uses __this lpInitData:LPVOID, nInitData:DWORD, myFormat:dword

	DebugOut "Create@CDataObject"

	invoke malloc, sizeof CDataObject
	.if (eax == 0)
		ret
	.endif
	mov __this, eax
	mov m__IDataObject.lpVtbl, offset vtblIDataObject
	mov m_dwRefCount, 1
	mov eax,myFormat
	mov m_myFormat,eax

	DebugOut "Create@CDataObject, format=%X", eax

;------------------ allocate the data buffer
	mov eax, nInitData
	mov m_lenData, eax
	invoke malloc, eax
	mov m_pData, eax
	mov eax, myFormat
	.if (eax == g_dwMyCBFormat)
		invoke VariantInit, m_pData
		invoke VariantCopy, m_pData, lpInitData
	.else
		invoke CopyMemory, m_pData, lpInitData, m_lenData
	.endif

;------------------ Allocate the FORMATETC to describe our data
	invoke malloc, SIZEOF FORMATETC
	mov m_pFmtEtc, eax

	mov ecx, myFormat
	mov (FORMATETC ptr [eax]).cfFormat, cx
	mov (FORMATETC ptr [eax]).ptd,NULL
	mov (FORMATETC ptr [eax]).dwAspect,DVASPECT_CONTENT
	mov (FORMATETC ptr [eax]).lindex,-1
	mov (FORMATETC ptr [eax]).tymed, TYMED_HGLOBAL

	mov m_NumFmtEtc,1

	return __this
	align 4

Create@CDataObject endp


;--- IUnknown methods

QueryInterface proc this_:ptr CDataObject, riid:REFIID, ppReturn:ptr LPUNKNOWN

;	DebugOut "CDataObject::QueryInterface"
	invoke IsInterfaceSupported, riid, offset iftab, NUMIFENTRIES,  this_, ppReturn
	ret
	align 4

QueryInterface endp

AddRef proc this_:ptr CDataObject

	mov eax, this_
	inc [eax].CDataObject.dwRefCount
	mov eax, [eax].CDataObject.dwRefCount
	ret
	align 4
AddRef endp

Release proc uses __this this_:ptr CDataObject

	mov __this, this_
	dec m_dwRefCount
;---------------------------- if reference count is zero destroy the object
	mov eax, m_dwRefCount
	.if (eax == 0)
;---------------------------- free allocated data
		mov eax, m_myFormat
		.if (eax == g_dwMyCBFormat)
			invoke VariantClear, m_pData
		.endif
		invoke free, m_pData
		invoke free, m_pFmtEtc
;---------------------------- free the object
		invoke free, __this
		DebugOut "CDataObject destroyed"
		xor eax, eax
	.endif
	ret
	align 4
Release endp

;-----------------------------------------------------------------------------
;*** Retrieves data described by a specific FormatEtc into a StgMedium allocated by this function

GetData proc uses __this this_:ptr CDataObject, pFE:ptr FORMATETC, pSTM:ptr STGMEDIUM

LOCAL	handle:DWORD
local	pMem:LPVOID

		mov __this, this_
;----------------------------- Check the aspects we support.
		mov ecx, pFE
		.if (!([ecx].FORMATETC.dwAspect & DVASPECT_CONTENT))
			mov eax,DATA_E_FORMATETC
			jmp exit
		.endif
;----------------------------- only 1 format supported
		movzx eax, [ecx].FORMATETC.cfFormat
		.if (eax != m_myFormat)
			mov eax,DATA_E_FORMATETC
			jmp exit
		.endif
;----------------------------- copy our data for the dragtarget
		invoke GlobalAlloc, GMEM_SHARE or GMEM_MOVEABLE, m_lenData
		.if (!eax)
			mov eax,STG_E_MEDIUMFULL
			jmp exit
		.endif
		mov handle, eax
		invoke GlobalLock, eax
		mov pMem, eax
		mov eax, g_dwMyCBFormat
		.if (eax == m_myFormat)
			invoke VariantInit, pMem
			invoke VariantCopy, pMem, m_pData
		.else
			invoke CopyMemory, pMem, m_pData, m_lenData
		.endif
		invoke GlobalUnlock, handle
;----------------------------- put the data in the storage medium
		mov ecx, pSTM
		mov eax, handle
		mov [ecx].STGMEDIUM.hGlobal, eax
		mov [ecx].STGMEDIUM.tymed, TYMED_HGLOBAL
		mov eax, g_dwMyCBFormat
		.if (eax == m_myFormat)
			mov eax, m_pData
			mov eax, [eax].VARIANT.punkVal
			mov [ecx].STGMEDIUM.pUnkForRelease, eax
		.else
			mov [ecx].STGMEDIUM.pUnkForRelease, NULL
		.endif
		mov eax, S_OK
exit:
ifdef _DEBUG
		.if (eax != S_OK)
			mov ecx, pFE
			movzx ecx, [ecx].FORMATETC.cfFormat
			DebugOut "CDataObject::GetData(%X) failed [%X]", ecx, eax
		.else
			DebugOut "CDataObject::GetData succeeded"
		.endif
endif
		ret
		align 4
GetData endp

;-----------------------------------------------------------------------------
;*** Renders the specific FormatEtc into caller-allocated medium provided in pSTM

GetDataHere proc this_:ptr CDataObject, pFE:DWORD, pSTM:DWORD

		DebugOut "CDataObject::GetDataHere(%X, %X)", pFE, pSTM
		mov eax,E_NOTIMPL
		ret
		align 4
GetDataHere endp

;-----------------------------------------------------------------------------
;*** Tests if a call to GetData with this FormatEtc will provide any rendering

QueryGetData proc uses __this this_:ptr CDataObject, pFE:ptr FORMATETC

		mov __this, this_
;--------------------- Check the aspects we support.
		mov ecx, pFE
		.if (!([ecx].FORMATETC.dwAspect & DVASPECT_CONTENT))
			mov eax, DATA_E_FORMATETC
			jmp exit
		.endif
;--------------------- ...and our special clipboard format
		movzx eax, [ecx].FORMATETC.cfFormat
		.if (eax == m_myFormat)
			mov eax, S_OK
			jmp exit
		.endif
		mov eax,S_FALSE
exit:
ifdef _DEBUG
		mov ecx, pFE
		movzx ecx, [ecx].FORMATETC.cfFormat
		.if (eax != S_OK)
			DebugOut "CDataObject::QueryGetData(%X) failed [%X]", ecx, eax
		.else
			DebugOut "CDataObject::QueryGetData(%X) succeeded", ecx
		.endif
endif
		ret
		align 4
QueryGetData endp

;-----------------------------------------------------------------------------
;*** Provides the caller with an equivalent FormatEtc to the one provided

GetCanonicalFormatEtc proc uses __this esi edi this_:ptr CDataObject, pFEIn:ptr FORMATETC, pFEOut:ptr FORMATETC

		DebugOut "IDataObject::GetCanonicalFormatEtc(%X, %X)", pFEIn, pFEOut

		mov esi, pFEIn
		mov edi, pFEOut
		mov ecx, SIZEOF FORMATETC
		mov edx, edi
		rep movsb
		mov [edx].FORMATETC.ptd, NULL
		mov eax, DATA_S_SAMEFORMATETC
		ret
		align 4
GetCanonicalFormatEtc endp

;-----------------------------------------------------------------------------
;*** Places data described by a FormatEtc and living in a StgMedium into the object

SetData proc this_:ptr CDataObject, pFE:ptr FORMATETC, pSTM:ptr STGMEDIUM, fRelease:BOOL

		DebugOut "IDataObject::SetData(%X, %X, %X)", pFE, pSTM, fRelease
		mov eax,E_NOTIMPL
		ret
		align 4
SetData endp

;-----------------------------------------------------------------------------
;*** Returns an IEnumFORMATETC object through which the caller can iterate
;*** to learn about all the data formats this object can provide

EnumFormatEtc proc uses __this this_:ptr CDataObject, dwDir:DWORD, ppEnum:ptr ptr IEnumFORMATETC

		mov __this,this_
		DebugOut "IDataObject::EnumFormatEtc(%X, %X)", dwDir, ppEnum
		.IF ((dwDir == DATADIR_GET) || (dwDir == DATADIR_SET))
			invoke Create@CEnumFORMATETC, __this, 1, m_pFmtEtc
		.ELSE
			xor eax,eax
		.ENDIF
		mov ecx, ppEnum
		mov DWORD PTR [ecx], eax
		.if (eax)
			mov eax, S_OK
		.else
			mov eax, E_OUTOFMEMORY
		.endif
		ret
		align 4

EnumFormatEtc endp

DAdvise proc this_:ptr CDataObject, pFE:ptr FORMATETC, dwFlags:DWORD, pAdviseSink:LPADVISESINK, pdwConn:ptr DWORD

		DebugOut "IDataObject::DAdvise(%X, %X, %X, %X)", pFE, dwFlags, pAdviseSink, pdwConn
		mov ecx, pdwConn
		mov dword ptr [ecx], NULL
		return OLE_E_ADVISENOTSUPPORTED
		align 4

DAdvise endp

DUnadvise proc this_:ptr CDataObject, dwConn:DWORD

		DebugOut "IDataObject::DUnadvise(%X)", dwConn
		return OLE_E_ADVISENOTSUPPORTED
		align 4

DUnadvise endp

EnumDAdvise proc this_:ptr CDataObject, ppEnum:ptr ptr IEnumSTATDATA

		DebugOut "IDataObject::EnumDAdvise(%X)", ppEnum
		mov ecx, ppEnum
		mov dword ptr [ecx], NULL
		return OLE_E_ADVISENOTSUPPORTED 
		align 4

EnumDAdvise endp

	end
