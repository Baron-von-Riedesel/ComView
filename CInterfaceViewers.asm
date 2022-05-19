
;*** definition of internal interface viewers

	.386
	.model flat,stdcall
	option casemap :none   ; case sensitive
	option proc:private

	include COMView.inc
	include statusbar.inc
	include servprov.inc
	include shlobj.inc
	include classes.inc
	include rsrc.inc
	include debugout.inc

ViewOleWindow				proto protoViewProc
ViewDataObject				proto protoViewProc
ViewSpecifyPropertyPages	proto protoViewProc
ViewServiceProvider			proto protoViewProc
ViewClassFactory2			proto protoViewProc
ViewPersistFile				proto protoViewProc
ViewDockingWindow			proto protoViewProc
ViewPropertyStorage			proto protoViewProc
ViewPropertySetStorage		proto protoViewProc

SafeRelease proto pUnknown:LPUNKNOWN

	.const


InterfaceViewerTab label dword
		dd offset IID_IOleWindow, offset ViewOleWindow
		dd offset IID_IDataObject, offset ViewDataObject
		dd offset IID_ISpecifyPropertyPages, offset ViewSpecifyPropertyPages
		dd offset IID_IServiceProvider, offset ViewServiceProvider
		dd offset IID_IClassFactory2, offset ViewClassFactory2
		dd offset IID_IPersistFile, offset ViewPersistFile
		dd offset IID_IDockingWindow, offset ViewDockingWindow
		dd offset IID_IStorage, offset ViewStorage
		dd offset IID_IPropertyStorage, offset ViewPropertyStorage
		dd offset IID_IPropertySetStorage, offset ViewPropertySetStorage
		dd offset IID_IOleLink, offset ViewOleLink
		dd 0

	.code

DisplayError proc hWnd:HWND, pszFormatString:LPSTR, hr:DWORD

local	szText[260]:BYTE

		invoke wsprintf, addr szText, pszFormatString, hr
		invoke GetDlgItem, hWnd, IDC_STATUSBAR
		mov ecx, eax
		StatusBar_SetText ecx, 0, addr szText
		invoke MessageBeep, MB_OK
		ret
DisplayError endp

		.const

ColumnsDetails label CColHdr
		CColHdr <CStr("Name")		, 40>
		CColHdr <CStr("Value")		, 60>
NUMCOLS_DETAILS textequ %($ - ColumnsDetails) / sizeof CColHdr

		.code

viewdetailproc proc uses esi hWnd:HWND, message:DWORD, wParam:WPARAM, lParam:LPARAM

local	hWndLV:HWND
local	pList:ptr CList
local	lvi:LVITEM

		mov eax, message
		.if (eax == WM_INITDIALOG)
			invoke GetDlgItem, hWnd, IDC_LIST1
			mov hWndLV, eax
			invoke SetLVColumns, hWndLV, NUMCOLS_DETAILS, addr ColumnsDetails
			xor esi, esi
			invoke GetItem@CList, lParam, esi
			invoke SetWindowText, hWnd, eax
			inc esi
			@mov lvi.iItem, 0
			mov lvi.mask_, LVIF_TEXT
			.while (1)
				invoke GetItem@CList, lParam, esi
				.break .if (!eax)
				@mov lvi.iSubItem, 0
				mov lvi.pszText, eax
				invoke ListView_InsertItem( hWndLV, addr lvi)
				inc esi
				inc lvi.iSubItem
				invoke GetItem@CList, lParam, esi
				.break .if (!eax)
				mov lvi.pszText, eax
				invoke ListView_SetItem( hWndLV, addr lvi)
				inc esi
				inc lvi.iItem
			.endw
			invoke GetDlgItem, hWnd, IDOK
			invoke ShowWindow, eax, SW_HIDE
			mov eax, 1
		.elseif (eax == WM_CLOSE)
			invoke EndDialog, hWnd, 0
		.elseif (eax == WM_COMMAND)
			movzx eax, word ptr wParam+0
			.if (eax == IDCANCEL)
				invoke EndDialog, hWnd, 0
			.elseif (eax == IDOK)
				invoke EndDialog, hWnd, eax
			.endif
		.else
			xor eax, eax
		.endif
		ret
		align 4

viewdetailproc endp


ViewPersistFile proc hWnd:HWND, pUnknown:LPUNKNOWN, pClsid:ptr CLSID

local	pwszFile:ptr WORD
local	pPersistFile:LPPERSISTFILE
local	szFile[MAX_PATH]:byte

		invoke vf(pUnknown, IUnknown, QueryInterface), addr IID_IPersistFile, addr pPersistFile
		.if (eax == S_OK)
			invoke vf(pPersistFile, IPersistFile, GetCurFile), addr pwszFile
			.if (eax == S_OK)
				invoke WideCharToMultiByte,CP_ACP,0,pwszFile,-1,addr szFile,sizeof szFile,0,0 
				invoke CoTaskMemFree, pwszFile
				invoke MessageBox, hWnd, addr szFile, CStr("Current File"), MB_OK
			.else
				invoke DisplayError, hWnd, CStr("IPersistFile::GetCurFile returned %X"), eax
			.endif
			invoke vf(pPersistFile, IUnknown, Release)
		.endif
		ret
ViewPersistFile endp


;-------------------------------------------------------------------------


ViewOleWindow proc hWnd:HWND, pUnknown:LPUNKNOWN, pClsid:ptr CLSID

local	pOleWindow:LPOLEWINDOW
local	hOleWnd:HWND

		invoke vf(pUnknown, IUnknown, QueryInterface), addr IID_IOleWindow, addr pOleWindow
		.if (eax == S_OK)
			invoke vf(pOleWindow, IOleWindow, GetWindow_), addr hOleWnd
			.if (eax == S_OK)
				invoke DisplayError, hWnd, CStr("IOleWindow::GetWindow hWnd=%X"), hOleWnd
				invoke IsWindowVisible, hOleWnd
				.if (eax)
					invoke ShowWindow, hOleWnd, SW_HIDE
					invoke MessageBox, hWnd, CStr("Press Ok to restore window"), addr g_szHint, MB_OK
					invoke ShowWindow, hOleWnd, SW_SHOWNORMAL
				.endif
			.else
				invoke DisplayError, hWnd, CStr("IOleWindow::GetWindow returned %X"), eax
			.endif
			invoke vf(pOleWindow, IUnknown, Release)
		.endif
		ret

ViewOleWindow endp


;-------------------------------------------------------------------------


ViewDataObject proc hWnd:HWND, pUnknown:LPUNKNOWN, pClsid:ptr CLSID

local	pDataObject:LPDATAOBJECT
local	pEnumFORMATETC:LPENUMFORMATETC
local	szText[80]:byte

		invoke vf(pUnknown, IUnknown, QueryInterface), addr IID_IDataObject, addr pDataObject
		.if (eax == S_OK)
			invoke vf(pDataObject, IDataObject, EnumFormatEtc), DATADIR_GET, addr pEnumFORMATETC
			.if (eax == S_OK)
				invoke Create@CEnumFORMATETCDlg, pEnumFORMATETC, pDataObject
				.if (eax)
					invoke CreateDialogParam, g_hInstance, IDD_ENUMFORMATETCDLG, hWnd, classdialogproc, eax
				.endif
				invoke vf(pEnumFORMATETC, IUnknown, Release)
			.else
				invoke DisplayError, hWnd, CStr("EnumFormatEtc failed [%X]"), eax
			.endif
			invoke vf(pDataObject, IUnknown, Release)
		.endif
		ret
ViewDataObject endp


;-------------------------------------------------------------------------


ViewSpecifyPropertyPages proc uses esi edi ebx hWnd:HWND, pUnknown:LPUNKNOWN, pClsid:ptr CLSID

local pSpecifyPropertyPages:LPSPECIFYPROPERTYPAGES
local cauuid:CAUUID
local wszGUID[40]:word

	invoke vf(pUnknown, IUnknown, QueryInterface), addr IID_ISpecifyPropertyPages, addr pSpecifyPropertyPages
	.if (eax == S_OK)
		invoke vf(pSpecifyPropertyPages, ISpecifyPropertyPages, GetPages), addr cauuid
		.if (eax == S_OK)
			.if (cauuid.cElems)
				invoke StringFromGUID2, pClsid, addr wszGUID, 40
				invoke OleCreatePropertyFrame, hWnd, 20, 20, addr wszGUID,\
					1, addr pUnknown, cauuid.cElems, cauuid.pElems, g_LCID, NULL, NULL
				invoke CoTaskMemFree, cauuid.pElems
			.endif	
		.endif
		invoke vf(pSpecifyPropertyPages, IUnknown, Release)
	.endif
	ret
	align 4

ViewSpecifyPropertyPages endp


;-------------------------------------------------------------------------


serviceproviderdlgproc proc hWnd:HWND, message:DWORD, wParam:WPARAM, lParam:LPARAM

local pServiceProvider:LPSERVICEPROVIDER
local pUnknown:LPUNKNOWN
local hWndOwner:HWND
local szSID[40]:byte
local szIID[40]:byte
local wszSID[40]:word
local wszIID[40]:word
local sid:GUID
local iid:IID
local szText[128]:byte

	mov eax, message
	.if (eax == WM_INITDIALOG)
		mov ecx, lParam
		invoke SetWindowLong, hWnd, DWL_USER, ecx
		mov eax, 1

	.elseif (eax == WM_CLOSE)
		invoke EndDialog, hWnd, 0
	.elseif (eax == WM_COMMAND)
		movzx eax, word ptr wParam+0
		.if (eax == IDCANCEL)
			invoke EndDialog, hWnd, 0
		.elseif (eax == IDOK)
			invoke GetParent, hWnd
			mov hWndOwner, eax
			invoke GetDlgItemText, hWnd, IDC_EDIT1, addr szSID, sizeof szSID
			invoke GetDlgItemText, hWnd, IDC_EDIT2, addr szIID, sizeof szIID
			.if ((!szSID) || (!szIID))
				invoke MessageBeep, MB_OK
				jmp exit
			.endif
			invoke MultiByteToWideChar,CP_ACP,MB_PRECOMPOSED,\
						addr szSID, -1, addr wszSID, 40 
			invoke MultiByteToWideChar,CP_ACP,MB_PRECOMPOSED,\
						addr szIID, -1, addr wszIID, 40 
			invoke CLSIDFromString, addr wszSID, addr sid
			.if (eax != S_OK)
				invoke DisplayError, hWndOwner, CStr("CLSIDFromString for Service GUID failed [%X]"), eax
				jmp exit
			.endif
			invoke CLSIDFromString, addr wszIID, addr iid
			.if (eax != S_OK)
				invoke DisplayError, hWndOwner, CStr("CLSIDFromString for IID failed [%X]"), eax
				jmp exit
			.endif
			invoke GetWindowLong, hWnd, DWL_USER
			mov pServiceProvider, eax
			invoke vf(pServiceProvider, IServiceProvider, QueryService), addr sid, addr iid, addr pUnknown
			.if (eax == S_OK)
				invoke Create@CObjectItem, pUnknown, addr sid
				.if (eax)
					push eax
					invoke vf(eax, IObjectItem, ShowObjectDlg), hWndOwner
					pop eax
					invoke vf(eax, IObjectItem, Release)
				.endif
				invoke vf(pUnknown, IUnknown, Release)
				invoke EndDialog, hWnd, 0
			.else
				invoke DisplayError, hWndOwner, CStr("QueryService failed [%X]"), eax
			.endif
		.endif
	.else
exit:
		xor eax, eax
	.endif
	ret
	align 4

serviceproviderdlgproc endp


ViewServiceProvider proc hWnd:HWND, pUnknown:LPUNKNOWN, pClsid:ptr CLSID

local pServiceProvider:LPSERVICEPROVIDER
local cauuid:CAUUID
local wszGUID[40]:word

		invoke vf(pUnknown, IUnknown, QueryInterface), addr IID_IServiceProvider, addr pServiceProvider
		.if (eax == S_OK)
			invoke DialogBoxParam, g_hInstance, IDD_SERVICEPROVIDERDLG, hWnd, serviceproviderdlgproc, pServiceProvider
			invoke vf(pServiceProvider, IServiceProvider, Release)
		.endif
		ret
		align 4

ViewServiceProvider endp


;-------------------------------------------------------------------------


ViewClassFactory2 proc hWnd:HWND, pUnknown:LPUNKNOWN, pClsid:ptr CLSID

local pClassFactory2:LPCLASSFACTORY2
local licinfo:LICINFO
local bstr:BSTR
local dwBytes:DWORD
local szText[256]:byte
local szLicKey[128]:byte

		invoke vf(pUnknown, IUnknown, QueryInterface), addr IID_IClassFactory2, addr pClassFactory2
		.if (eax == S_OK)
			mov licinfo.cbLicInfo, sizeof LICINFO
			invoke vf(pClassFactory2, IClassFactory2, GetLicInfo), addr licinfo
			.if (eax == S_OK)
				invoke wsprintf, addr szText, CStr("Runtime Key Available: %u",10,"Full License exists: %u"),
					licinfo.fRuntimeKeyAvail, licinfo.fLicVerified
				mov dwBytes, eax
				.if (licinfo.fRuntimeKeyAvail)
					invoke vf(pClassFactory2, IClassFactory2, RequestLicKey), NULL, addr bstr
					.if (eax == S_OK)
						invoke WideCharToMultiByte,CP_ACP,0,bstr,-1,addr szLicKey,sizeof szLicKey,0,0 
						lea ecx, szText
						add ecx, dwBytes
						invoke wsprintf, ecx, CStr(10,"License Key: %s"), addr szLicKey
						invoke SysFreeString, bstr
					.else
						invoke DisplayError, hWnd, CStr("RequestLicKey failed [%X]"), eax
					.endif
				.endif
				invoke MessageBox, hWnd, addr szText, CStr("GetLicInfo()"), MB_OK
			.else
				invoke DisplayError, hWnd, CStr("GetLicInfo failed [%X]"), eax
			.endif
			invoke vf(pClassFactory2, IUnknown, Release)
		.endif
		ret
		align 4

ViewClassFactory2 endp


;-------------------------------------------------------------------------


ViewDockingWindow proc hWnd:HWND, pUnknown:LPUNKNOWN, pClsid:ptr CLSID

local	pDockingWindow:ptr IDockingWindow
local	hOleWnd:HWND

		invoke vf(pUnknown, IUnknown, QueryInterface), addr IID_IDockingWindow, addr pDockingWindow
		.if (eax == S_OK)
			invoke vf(pDockingWindow, IDockingWindow, GetWindow), addr hOleWnd
			invoke DisplayError, hWnd, CStr("IDockingWindow::GetWindow returned hWnd=%X"), hOleWnd
			invoke vf(pDockingWindow, IDockingWindow, ShowDW), FALSE
			invoke MessageBox, hWnd, CStr("Press Ok to restore docking window"), addr g_szHint, MB_OK
			invoke vf(pDockingWindow, IDockingWindow, ShowDW), TRUE
			invoke vf(pDockingWindow, IUnknown, Release)
		.endif
		ret

ViewDockingWindow endp


;-------------------------------------------------------------------------


ViewStorage proc hWnd:HWND, pUnknown:LPUNKNOWN, pClsid:ptr CLSID

		invoke Create@CViewStorageDlg, pUnknown, NULL, NULL
		.if (eax)
			invoke Show@CViewStorageDlg, eax, NULL
		.endif
		ret
			
ViewStorage endp

;-------------------------------------------------------------------------


ViewPropertyStorage proc public hWnd:HWND, pUnknown:LPUNKNOWN, pClsid:ptr CLSID

		invoke Create@CPropertyStorageDlg, pUnknown, pClsid
		.if (eax)
			invoke Show@CPropertyStorageDlg, eax, hWnd
		.endif
		ret

ViewPropertyStorage endp


;-------------------------------------------------------------------------

	.const

ColumnsEnumPropSet label CColHdr
		CColHdr <CStr("FmtID")	, 50>
		CColHdr <g_szCLSID		, 50>
NUMCOLS_ENUMPROPSET textequ %($ - ColumnsEnumPropSet) / sizeof CColHdr

	.code

OnInitDialog proc hWnd:HWND, lParam:LPARAM

local	pPropSetStg:LPPROPERTYSETSTORAGE
local	pEnumSTATPROPSETSTG:LPENUMSTATPROPSETSTG
local	pPropertyStorage:LPPROPERTYSTORAGE
local	pEnumSTATPROPSTG:LPENUMSTATPROPSTG
local	spss:STATPROPSETSTG
local	hWndLV:HWND
local	wszGUID[40]:word
local	szGUID[40]:byte
local	lvi:LVITEM

		invoke GetDlgItem, hWnd, IDC_LIST1
		mov hWndLV, eax
		invoke SetLVColumns, hWndLV, NUMCOLS_ENUMPROPSET, addr ColumnsEnumPropSet

		invoke SetWindowText, hWnd, CStr("View IPropertySetStorage")

		mov ecx, lParam
		mov pPropSetStg, ecx
		invoke SetWindowLong, hWnd, DWL_USER, ecx

		invoke vf(pPropSetStg, IPropertySetStorage, Enum), addr pEnumSTATPROPSETSTG
		.if (eax != S_OK)
			jmp done
		.endif
		mov lvi.iItem, 0
		.while (1)
			invoke vf(pEnumSTATPROPSETSTG, IEnumSTATPROPSETSTG, Next), 1, addr spss, NULL
			.break .if (eax != S_OK)
			invoke StringFromGUID2, addr spss.fmtid, addr wszGUID, 40
			invoke WideCharToMultiByte,CP_ACP,0, addr wszGUID,-1,addr szGUID,sizeof szGUID,0,0
			mov lvi.mask_, LVIF_TEXT
			lea eax, szGUID
			mov lvi.pszText, eax
			mov lvi.iSubItem, 0
			invoke ListView_InsertItem( hWndLV, addr lvi)

			invoke StringFromGUID2, addr spss.clsid, addr wszGUID, 40
			invoke WideCharToMultiByte,CP_ACP,0, addr wszGUID,-1,addr szGUID,sizeof szGUID,0,0
			lea eax, szGUID
			mov lvi.pszText, eax
			inc lvi.iSubItem
			invoke ListView_SetItem( hWndLV, addr lvi)
			inc lvi.iItem
		.endw
		invoke vf(pEnumSTATPROPSETSTG, IUnknown, Release)
done:
		ret
OnInitDialog endp

enumpropsetproc proc hWnd:HWND, message:DWORD, wParam:WPARAM, lParam:LPARAM
	
local	pPropSetStg:LPPROPERTYSETSTORAGE
local	pPropertyStorage:LPPROPERTYSTORAGE
local	hWndLV:HWND
local	wszGUID[40]:word
local	szGUID[40]:byte
local	fmtid:FMTID
local	lvi:LVITEM

	mov eax, message
	.if (eax == WM_INITDIALOG)

		invoke OnInitDialog, hWnd, lParam
		mov eax, 1

	.elseif (eax == WM_CLOSE)

		invoke EndDialog, hWnd, 0

	.elseif (eax == WM_COMMAND)

		movzx eax, word ptr wParam+0
		.if (eax == IDCANCEL)
			invoke EndDialog, hWnd, 0
		.endif

	.elseif (eax == WM_NOTIFY)

		invoke GetDlgItem, hWnd, IDC_LIST1
		mov hWndLV, eax

		mov ecx, lParam
		.if ([ecx].NMHDR.idFrom == IDC_LIST1)
			.if ([ecx].NMHDR.code == NM_DBLCLK)
				mov eax, [ecx].NMLISTVIEW.iItem
				mov lvi.iItem, eax
				mov lvi.iSubItem, 0
				mov lvi.mask_, LVIF_TEXT
				lea eax, szGUID
				mov lvi.pszText, eax
				mov lvi.cchTextMax, sizeof szGUID
				invoke ListView_GetItem( hWndLV, addr lvi)
				invoke MultiByteToWideChar,CP_ACP,MB_PRECOMPOSED,\
						addr szGUID, -1, addr wszGUID, 40 
				invoke CLSIDFromString, addr wszGUID, addr fmtid
				invoke GetWindowLong, hWnd, DWL_USER
				mov pPropSetStg, eax
				invoke vf(pPropSetStg, IPropertySetStorage, Open),\
					addr fmtid, STGM_READ or STGM_SHARE_EXCLUSIVE, addr pPropertyStorage
				.if (eax == S_OK)
					invoke ViewPropertyStorage, hWnd, pPropertyStorage, NULL
					invoke vf(pPropertyStorage, IUnknown, Release)
				.else
					invoke OutputMessage, hWnd, eax, CStr("IPropertySetStorage::Open Error"), 0
				.endif
			.endif
		.endif
	.else
		xor eax, eax
	.endif
	ret

enumpropsetproc endp

ViewPropertySetStorage proc hWnd:HWND, pUnknown:LPUNKNOWN, pClsid:ptr CLSID

local	pPropSetStg:LPPROPERTYSETSTORAGE

		invoke vf(pUnknown, IUnknown, QueryInterface), addr IID_IPropertySetStorage, addr pPropSetStg
		.if (eax == S_OK)
			invoke DialogBoxParam, g_hInstance, IDD_ENUMFORMATETCDLG, hWnd, enumpropsetproc, pPropSetStg
			invoke vf(pPropSetStg, IUnknown, Release)
		.endif
		ret
ViewPropertySetStorage endp

;-------------------------------------------------------------------------

ViewOleLink proc hWnd:HWND, pUnknown:LPUNKNOWN, pClsid:ptr CLSID

local	pOleLink:LPOLELINK
local	pwszString:LPOLESTR
local	dwOptions:DWORD
local	pList:ptr CList
local	szString[MAX_PATH]:byte

		invoke vf(pUnknown, IUnknown, QueryInterface), addr IID_IOleLink, addr pOleLink
		.if (eax == S_OK)
			invoke Create@CList, LISTF_STRINGS
			mov pList, eax
			invoke AddItem@CList, pList, CStr("IOleLink")
			invoke AddItem@CList, pList, CStr("Source Display Name")
			invoke vf(pOleLink, IOleLink, GetSourceDisplayName), addr pwszString
			.if (eax == S_OK)
				invoke WideCharToMultiByte,CP_ACP,0,pwszString,-1,addr szString,sizeof szString,0,0 
				invoke CoTaskMemFree, pwszString
			.else
				invoke wsprintf, addr szString, CStr("Failed [%X]"), eax
			.endif
			invoke AddItem@CList, pList, addr szString

			invoke AddItem@CList, pList, CStr("Update Options")
			invoke vf(pOleLink, IOleLink, GetUpdateOptions), addr dwOptions
			.if (eax == S_OK)
				invoke wsprintf, addr szString, CStr("%X"), dwOptions
			.else
				invoke wsprintf, addr szString, CStr("Failed [%X]"), eax
			.endif
			invoke AddItem@CList, pList, addr szString
			invoke DialogBoxParam, g_hInstance, IDD_VTBLDLG, hWnd, viewdetailproc, pList
			invoke vf(pOleLink, IUnknown, Release)
			invoke Destroy@CList, pList
		.endif
		ret
ViewOleLink endp

	end
