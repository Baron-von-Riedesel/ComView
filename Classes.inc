
;*** defines the classes of this project
;*** only public members/methods are listed here
;*** in normal case every class gets its own source file for method definition

;   class               description
;------------------------------------------------------------------------
;   CConnection         connection to an event source
;   CContainer          implements OLE container to host OLE controls
;   CCreateInclude      creates include file(s)
;   CDocument           document class of main dialog data
;   CEventSink          event sink (IConnectionPoint)
;   CInterfaceItem      describes 1 entry in HKEY_CLASSES_ROOT\Interface
;   CInterfaceList      List of all entries in "Interfaces"
;   CList               linked list
;   CObjectItem         item of CObjectList
;   COptions            document class of options

;   CAdviseSink         implements IAdviseSink (IOleObject, IViewObject, IDataObject)
;   CDataObject         implements IDataObject
;   CDropSource         implements IDropSource
;   CDropTarget         implements IDropTarget
;   CEnumFORMATETC      implements IEnumFORMATETC
;   CPropertyBag        implements IPropertyBag

;   CDlg                dialog box class (inherited from all DLG classes)
;   CMainDlg            main dialog class (this is a dialog app) (only 1 instance) 
;   CTypeLibDlg         dialog shows type library information (ITypeLib)
;   CTypeInfoDlg        tabbed dialog shows type information (ITypeInfo)
;   CTypeInfoPageDlg    handles 1 tab of CTypeInfoDlg
;   COptionDlg          handles tabbed option dialog
;   CAmbientPropDlg     handles ambient properties dialog
;   COptPageDlg         handles 1 tab of option dialog
;   CProgressDlg        small dialog showing progress bar for lengthy operations (save to disk)
;   CObjectDlg          dialog shows a created object
;   CViewObjectDlg      dialog shows object with UI (IOleObject)
;   CPropertiesDlg      dialog to show/edit properties
;   CParamsDlg          dialog to enter method parameters
;   CCollectionDlg      dialog to display a collection
;   CLogWindow          window to display logs (events)
;   CEnumFORMATETCDlg   dialog to display EnumFORMATETC results from IDataObject
;   CViewStorageDlg     dialog to display IStorage objects
;   CPropertyStorageDlg dialog to display IPropertyStorage objects


BEGIN_CLASS macro name, inherits
local classpointer
classname equ <name>
thisarg equ <:ptr classname>
classname struct
classpointer catstr <p>,<name>,< typedef ptr >,<name>
classpointer
ifnb <inherits>
	inherits <>
endif
	endm

END_CLASS macro
classname ends
classname equ <>
	endm

MEMBER macro names:VARARG
local x
	for y,<names>
x	CatStr <_this.>,<y>
m_&y	textequ x
	endm
	endm

;---------------------- start class definitions -------------------------

;*** CInterfaceItem describes 1 interface entry in registry

BEGIN_CLASS CInterfaceItem
iid			GUID <>
TypelibGUID	GUID <>
dwVerMajor	DWORD ?
dwVerMinor	DWORD ?
pszName		LPSTR ?
END_CLASS

Create@CInterfaceItem	proto refiid:ptr IID
SetName@CInterfaceItem	proto :ptr CInterfaceItem, :LPSTR
SetTypeLibAttr@CInterfaceItem	proto :ptr CInterfaceItem, :ptr TLIBATTR
Destroy@CInterfaceItem	proto :ptr CInterfaceItem

;*** CInterfaceList list of interfaces class

BEGIN_CLASS CInterfaceList
pRegIIDs	PGUID ?
cntIID		dd ?
cntRegIID	dd ?
ppszNames	LPVOID ?
pszStrings	LPSTR ?
END_CLASS

Create@CInterfaceList	proto hWnd:HWND
Destroy@CInterfaceList	proto :ptr CInterfaceList
Find@CInterfaceList		proto :ptr CInterfaceList, :REFIID
AddIIDs@CInterfaceList	proto :ptr CInterfaceList, :DWORD, :ptr IID, :LPSTR
GetName@CInterfaceList	proto :ptr CInterfaceList, dwIndex:DWORD


;*** CObjectItem class: describes COM objects created by COMView ***

BEGIN_INTERFACE IObjectItem
	STDMETHOD	AddRef
	STDMETHOD	Release
	STDMETHOD	GetObjectDlg
	STDMETHOD	SetObjectDlg		,:ptr CObjectDlg
	STDMETHOD	GetViewObjectDlg
	STDMETHOD	SetViewObjectDlg	,:ptr CViewObjectDlg
	STDMETHOD	GetPropDlg
	STDMETHOD	SetPropDlg			,:ptr CPropertiesDlg
	STDMETHOD	GetViewStorageDlg
	STDMETHOD	SetViewStorageDlg	,:ptr CViewStorageDlg
	STDMETHOD	Lock_
	STDMETHOD	Unlock
	STDMETHOD	IsLocked
	STDMETHOD	GetRunLock
	STDMETHOD	SetRunLock			,:BOOL
	STDMETHOD	GetStorage
	STDMETHOD	SetStorage			,:LPSTORAGE
	STDMETHOD	GetMoniker
	STDMETHOD	SetMoniker			,:LPMONIKER
	STDMETHOD	GetConnectionList	,:BOOL
	STDMETHOD	ShowObjectDlg		,:HWND
	STDMETHOD	ShowPropertiesDlg	,:HWND
	STDMETHOD	ShowViewObjectDlg	,:HWND, pItem:ptr CInterfaceItem
	STDMETHOD	ShowViewStorageDlg	,:HWND
	STDMETHOD	GetFlags
	STDMETHOD	SetFlags			,:DWORD
	STDMETHOD	GetDisplayName		,:ptr LPOLESTR
	STDMETHOD	SetWindowText_		,:HWND
	STDMETHOD	AddFilename			,:HWND, :BOOL
	STDMETHOD	Close				,:DWORD
	STDMETHOD	GetCoClassTypeInfo	,:ptr LPTYPEINFO
	STDMETHOD	SetCoClassTypeInfo	,:LPTYPEINFO
	STDMETHOD	GetDefaultInterface ,:BOOL, :REFIID
	STDMETHOD	GetMQI
	STDMETHOD	SetMQI				,:DWORD, :ptr MULTI_QI
END_INTERFACE

ifndef INSIDE_COBJECTITEM
BEGIN_CLASS CObjectItem, IObjectItem
END_CLASS
endif

Create@CObjectItem			proto :LPUNKNOWN, :ptr GUID
GetGUID@CObjectItem			proto :ptr CObjectItem, :ptr GUID
GetUnknown@CObjectItem		proto :ptr CObjectItem
StringFromCLSID@CObjectItem	proto :ptr CObjectItem, :LPSTR
SetStgHlp@CObjectItem		proto :ptr CObjectItem, :LPUNKNOWN
Find@CObjectItem			proto :LPUNKNOWN
FindStorage@CObjectItem		proto :LPSTORAGE

LPOBJECTITEM	typedef ptr CObjectItem

OBJITEMF_INIT		equ 1		;dont initialize object
OBJITEMF_ROT		equ 2		;object is from ROT, dont activate
OBJITEMF_OPENVIEW	equ 4		;open view dialog
OBJITEMF_IGNOREOV	equ 8		;ignore open view flag


;--- CList

ifndef INSIDE_CLIST
BEGIN_CLASS CList
END_CLASS
endif

Create@CList				proto dwFlags:DWORD
Destroy@CList				proto :ptr CList
AddItem@CList				proto :ptr CList, pItem:LPVOID
DeleteItem@CList			proto :ptr CList, dwIndex:DWORD
GetItem@CList				proto :ptr CList, dwIndex:DWORD
GetItemCount@CList			proto :ptr CList
FindItem@CList				proto :ptr CList, pItem:LPVOID
GetItemData@CList			proto :ptr CList, dwIndex:DWORD
SetItemData@CList			proto :ptr CList, dwIndex:DWORD, :LPARAM
FindItemData@CList			proto :ptr CList, :LPARAM
FindAllItemData@CList		proto :ptr CList, :LPARAM

LISTF_STRINGS	equ 1	;items are strings to copy to list

;--- CConnection

ifndef INSIDE_CCONNECTION
BEGIN_CLASS CConnection
END_CLASS
endif

Create@CConnection		proto :ptr CObjectItem, :REFIID
Connect@CConnection		proto :ptr CConnection, :LPUNKNOWN, :HWND, :ptr LPSTR
Disconnect@CConnection	proto :ptr CConnection, :HWND
IsEqualGUID@CConnection	proto :ptr CConnection, :REFIID
Destroy@CConnection		proto :ptr CConnection

;*** CDlg dialog class ***

BEGIN_CLASS CDlg
hWnd		HWND ?
pDlgProc	DLGPROC ?
END_CLASS

;*** CTabDlgPage tab definition for tabbed dialogs

BEGIN_CLASS CTabDlgPage
pTabName	PBYTE ?
dwResID		dword ?
pDlgProc	DLGPROC ?
END_CLASS

;*** CColHdr

BEGIN_CLASS CColHdr
pColText	PPSTR ?
wWidth		dw ?
wFlags		dw ?
END_CLASS

FCOLHDR_RDXMASK		equ 1Fh
FCOLHDR_RDX10		equ 0Ah
FCOLHDR_RDX16		equ 10h
FCOLHDR_ABSOLUTE	equ 20h

;*** CMode mode description class

BEGIN_CLASS CMode
iMode		dd ?		;mode id (MODE_CLSID, MODE_TYPELIB, ...)
pColDesc	pCColHdr ?	;pointer to column descriptor table for this view
iNumCols	dd ?		;# of columns in this view
iCmdID		dd ?		;cmd id (IDM_CLSID, IDM_TYPELIB, ...)
pszRoot		LPSTR ?		;registry root key
pszTabText	LPSTR ?		;tab caption ("CLSID", "TypeLib", ...)
END_CLASS

;*** CDocument holds all list information

ifndef INSIDE_CDOCUMENT
BEGIN_CLASS CDocument
END_CLASS
endif
Create@CDocument		proto :HWND, :DWORD, :DWORD, :LPSTR
Destroy@CDocument		proto :ptr CDocument
GetItemCount@CDocument	proto :ptr CDocument
GetItemData@CDocument	proto :ptr CDocument, :DWORD, :DWORD 
GetItemFlag@CDocument	proto :ptr CDocument, :DWORD, :DWORD
SetItemFlag@CDocument	proto :ptr CDocument, :DWORD, :DWORD, :DWORD
Sort@CDocument			proto :ptr CDocument, :DWORD, :DWORD, :DWORD
RefreshItem@CDocument	proto :ptr CDocument, :DWORD, :LPSTR
RemoveItem@CDocument	proto :ptr CDocument, :DWORD
FindROTItem@CDocument	proto :ptr CDocument, :DWORD

FLAG_FOCUSED	equ 1
FLAG_SELECTED	equ 2
FLAG_IMAGE		equ 4

DATACOL_IN_OBJECT	equ 4

;*** COptions class (static members only)

Load@COptions				proto :LPSTR
Save@COptions				proto :LPSTR
GetCoCreateFlags@COptions	proto
GetInterfaces@COptions		proto :ptr CInterfaceList


;*** CMainDlg main dialog class

ifndef INSIDE_CMAINDLG
BEGIN_CLASS CMainDlg, CDlg
END_CLASS
endif
Create@CMainDlg		proto
Destroy@CMainDlg	proto :ptr CMainDlg
;;GetMode@CMainDlg	proto :ptr CMainDlg
;;GetMenu@CMainDlg	Proto :ptr CMainDlg
SetErrorText@CMainDlg proto :ptr CMainDlg, pszError:LPSTR, hr:DWORD, bBeep:BOOL
LoadFile@CMainDlg	proto :ptr CMainDlg, pszFile:LPSTR, bDisplayError:BOOL
CreateLink@CMainDlg	proto :ptr CMainDlg, pszFile:LPSTR
Register@CMainDlg	proto :ptr CMainDlg, pszFile:LPSTR
OpenStorage@CMainDlg proto :ptr CMainDlg, pszFile:LPSTR, pStorage:LPSTORAGE
OpenStream@CMainDlg	proto :ptr CMainDlg, pszFile:LPSTR
SmartLoad@CMainDlg	proto :ptr CMainDlg, pszFile:LPSTR
RefreshView@CMainDlg proto :ptr CMainDlg, dwMode:DWORD
Reset@CMainDlg		proto :ptr CMainDlg
FileOperation@CMainDlg proto :ptr CMainDlg, dwMode:DWORD, pszFile:LPSTR
UserColChanged@CMainDlg proto :ptr CMainDlg, dwMode:DWORD

;--- static methods

SetBusyState@CMainDlg proto dwStatus:BOOL


;*** iMode constants  enum

MODE_CLSID		equ 0
MODE_TYPELIB	equ 1
MODE_INTERFACE	equ 2
MODE_APPID		equ 3
MODE_COMPCAT	equ 4
MODE_HKCR		equ 5
MODE_OBJECT		equ 6
MODE_ROT		equ 7
MODE_MAX		equ 7

MODE_TYPEINFO	equ MODE_MAX+1	;not a real mode, but used in typelib dialog


;*** COptPageDlg options page dialog class 

ifndef INSIDE_COPTIONDLG
BEGIN_CLASS COptPageDlg, CDlg
END_CLASS
endif

;*** COptionDlg options tabbed dialog class

ifndef INSIDE_COPTIONDLG
BEGIN_CLASS COptionDlg, CDlg
END_CLASS
endif
Create@COptionDlg	proto
Show@COptionDlg		proto :ptr COptionDlg, hWnd:HWND
Destroy@COptionDlg	proto :ptr COptionDlg

;*** CProgressDlg progress bar dialog class

ifndef INSIDE_CPROGRESSDLG
BEGIN_CLASS CProgressDlg, CDlg
END_CLASS
endif
Create@CProgressDlg	proto hWndLV:HWND, pszFileName:LPSTR,iSaveMode:dword, iNumCols:DWORD

;*** iSaveMode enum

SAVE_DISK		equ 0
SAVE_CLIPBOARD	equ 1

;*** CTypeLibDlg type library dialog class 

ifndef INSIDE_CTYPELIBDLG
BEGIN_CLASS CTypeLibDlg, CDlg
END_CLASS
endif
Create@CTypeLibDlg	proto pGUID:ptr GUID, dwVerMajor:dword,dwVerMinor:dword,lcid:dword,pIID:ptr IID
Create2@CTypeLibDlg	proto pszTypeLib:LPSTR ,pIID:ptr IID, bShowError:BOOL
Create3@CTypeLibDlg proto pTypeLib:LPTYPELIB
Create4@CTypeLibDlg proto pTypeInfo:LPTYPEINFO
Show@CTypeLibDlg		proto :ptr CTypeLibDlg, hWnd:HWND, :BOOL
SetCallBack@CTypeLibDlg proto :ptr CTypeLibDlg, fnCallBack:LPVOID, dwCookie:DWORD


;*** CTypeInfoPageDlg type info page dialog class 

ifndef INSIDE_CTYPEINFOPAGEDLG
BEGIN_CLASS CTypeInfoPageDlg, CDlg
END_CLASS
endif

;*** CTypeInfoDlg type info tabbed dialog class 

ifndef INSIDE_CTYPEINFODLG
BEGIN_CLASS CTypeInfoDlg, CDlg
END_CLASS
endif
Create@CTypeInfoDlg		proto pITypeLib: LPTYPELIB, iIndex:dword
Create2@CTypeInfoDlg	proto pTypeInfo:LPTYPEINFO
Create3@CTypeInfoDlg	proto refIID:REFIID, refTLB:REFGUID, verMajor:DWORD, verMinor:DWORD
Show@CTypeInfoDlg		proto :ptr CTypeInfoDlg, hWnd:HWND, dwFlags:DWORD
SetTab@CTypeInfoDlg		proto :ptr CTypeInfoDlg, dwTabIndex:DWORD
FindFunc@CTypeInfoDlg	proto :ptr CTypeInfoDlg, memberID:MEMBERID, invkind:INVOKEKIND
FindVar@CTypeInfoDlg	proto :ptr CTypeInfoDlg, memberID:MEMBERID
SetTypeInfo@CTypeInfoDlg proto :ptr CTypeInfoDlg, pTypeInfo:LPTYPEINFO

;--- flags for Show@CTypeInfoDlg

TYPEINFODLG_FACTIVATE	equ 001h
TYPEINFODLG_FTILE		equ 002h

;*** CCreateInclude class 

ifndef INSIDE_CCREATEINCLUDE
BEGIN_CLASS CCreateInclude
END_CLASS
endif
Create@CCreateInclude	proto pszGUID:LPSTR, lcid:LCID, dwVerMajor:DWORD, dwVerMinor:DWORD
Create2@CCreateInclude	proto pszTypeLib:LPSTR
Create3@CCreateInclude	proto pTypeLib:LPTYPELIB
Run@CCreateInclude		proto :ptr CCreateInclude, hWnd:HWND, dwMode:DWORD
Destroy@CCreateInclude	proto :ptr CCreateInclude
SetOutputFile@CCreateInclude proto :ptr CCreateInclude, :LPSTR
SetStdOut@CCreateInclude proto :ptr CCreateInclude, :DWORD

INCMODE_BASIC		equ 0
INCMODE_DISPHLP		equ 1

;*** CObjectDlg dialog class 

ifndef INSIDE_COBJECTDLG
BEGIN_CLASS CObjectDlg, CDlg
END_CLASS
endif
Create@CObjectDlg		proto pObjectItem:LPOBJECTITEM
Destroy@CObjectDlg		proto :ptr CObjectDlg
Show@CObjectDlg			proto :ptr CObjectDlg, hWnd:HWND
SetPosition@CObjectDlg	proto :ptr CObjectDlg, :ptr POINT

;*** CViewObjectDlg dialog class 

ifndef INSIDE_CVIEWOBJECTDLG
BEGIN_CLASS CViewObjectDlg, CDlg
END_CLASS
endif
Create@CViewObjectDlg			proto :LPOBJECTITEM, :pCInterfaceItem
SetStatusText@CViewObjectDlg	proto :ptr CViewObjectDlg, :DWORD, :LPSTR
SetMenu@CViewObjectDlg			proto :ptr CViewObjectDlg, iCmd:DWORD, dwFlags:DWORD
Show@CViewObjectDlg				proto :ptr CViewObjectDlg, :HWND
TranslateAccelerator@CViewObjectDlg proto :ptr CViewObjectDlg, :ptr MSG

;*** CContainer class 

ifndef INSIDE_CCONTAINER
BEGIN_CLASS CContainer
END_CLASS
endif
Create@CContainer		proto :LPOBJECTITEM, :ptr CViewObjectDlg, prect:ptr RECT
OnAmbientPropertyChange@CContainer proto :ptr CContainer, DispId:DWORD
DoVerb@CContainer		proto :ptr CContainer, dwVerb:DWORD
Close@CContainer		proto :ptr CContainer
Close2@CContainer		proto :ptr CContainer
SetRect@CContainer		proto :ptr CContainer, prect:ptr RECT
Load@CContainer			proto :ptr CContainer
Save@CContainer			proto :ptr CContainer, iType:DWORD, bForceSave:BOOL
Advise@CContainer		proto :ptr CContainer, iType:DWORD
InPlaceDeactivate@CContainer	proto :ptr CContainer
UIDeactivate@CContainer	proto :ptr CContainer
IsWindowless@CContainer	proto :ptr CContainer
IsActive@CContainer		proto :ptr CContainer
OnMessage@CContainer	proto :ptr CContainer, :DWORD, :WPARAM, :LPARAM, :ptr DWORD
TranslateAccelerator@CContainer proto :ptr CContainer, :ptr MSG
OnActivate@CContainer	proto :ptr CContainer, :BOOL
OnMouseMove@CContainer	proto :ptr CContainer, :DWORD, :DWORD, :DWORD
OnSetCursor@CContainer	proto :ptr CContainer, :DWORD, :DWORD, :DWORD
OnMouseClick@CContainer	proto :ptr CContainer 
Update@CContainer		proto :ptr CContainer

SAVE_STREAM		equ 0
SAVE_STORAGE	equ 1
SAVE_FILE		equ 2
SAVE_PROPBAG	equ 3

;*** CPropertiesDlg dialog class 

ifndef INSIDE_CPROPERTIESDLG
BEGIN_CLASS CPropertiesDlg, CDlg
END_CLASS
endif
Create@CPropertiesDlg	proto :LPOBJECTITEM, :ptr CInterfaceItem
Create2@CPropertiesDlg	proto :LPUNKNOWN, :LPTYPEINFO
Show@CPropertiesDlg		proto :ptr CPropertiesDlg, :HWND
SetDispatch@CPropertiesDlg proto :ptr CPropertiesDlg, pNewDisp:LPDISPATCH, :BOOL
OnChanged@CPropertiesDlg proto :ptr CPropertiesDlg, dispId:DWORD
GetObjectItem@CPropertiesDlg proto :ptr CPropertiesDlg
SetWindowPos@CPropertiesDlg	proto :ptr CPropertiesDlg, :ptr POINT

;*** CParamsDlg dialog class 

ifndef INSIDE_CPARAMSDLG
BEGIN_CLASS CParamsDlg, CDlg
END_CLASS
endif
Create@CParamsDlg proto :LPTYPEINFO, :ptr FUNCDESC, :ptr CPropertiesDlg, :LPVOID

;*** CCollectionDlg dialog class 

ifndef INSIDE_CCOLLECTIONDLG
BEGIN_CLASS CCollectionDlg, CDlg
END_CLASS
endif
Create@CCollectionDlg proto :HWND, :LPDISPATCH, :MEMBERID, :ptr VARIANT, :ptr PARAMRETURN 

;*** CLogWindow class 

ifndef INSIDE_CLOGWINDOW
BEGIN_CLASS CLogWindow, CDlg
END_CLASS
endif
Create@CLogWindow proto
Destroy@CLogWindow proto
;--- static methods
Show@CLogWindow proto dwFlags:DWORD
printf@CLogWindow proto c pszFormat:LPSTR, varargs:VARARG
UpdateLogSwitch@CLogWindow proto

;*** CEventSink class 

ifndef INSIDE_CEVENTSINK
BEGIN_CLASS CEventSink
END_CLASS
endif
Create@CEventSink proto riid:REFIID, :ptr CObjectItem, :LPTYPEINFO

;*** CAdviseSink class 

ifndef INSIDE_CADVISESINK
BEGIN_CLASS CAdviseSink
END_CLASS
endif
Create@CAdviseSink proto


;*** CEnumFORMATETCDlg dialog class 

ifndef INSIDE_CENUMFORMATETCDLG
BEGIN_CLASS CEnumFORMATETCDlg, CDlg
END_CLASS
endif
Create@CEnumFORMATETCDlg proto :LPENUMFORMATETC, :LPDATAOBJECT

;*** CDropTarget

ifndef INSIDE_CDROPTARGET
BEGIN_CLASS CDropTarget
END_CLASS
endif
Create@CDropTarget proto :HWND

;*** CDropSource

ifndef INSIDE_CDROPSOURCE
BEGIN_CLASS CDropSource
END_CLASS
endif
Create@CDropSource proto

;*** CDataObject

ifndef INSIDE_CDATAOBJECT
BEGIN_CLASS CDataObject
END_CLASS
endif
Create@CDataObject proto :LPVOID, :DWORD, :DWORD

;*** CEnumFORMATETC

ifndef INSIDE_CENUMFORMATETC
BEGIN_CLASS CEnumFORMATETC
END_CLASS
endif
Create@CEnumFORMATETC proto :LPUNKNOWN, :DWORD, :ptr FORMATETC

ifdef _DEBUG
Create@CMallocSpy	proto
endif

ifndef INSIDE_CVIEWSTORAGEDLG
BEGIN_CLASS CViewStorageDlg, CDlg
END_CLASS
endif
Create@CViewStorageDlg	proto pStorage:LPSTORAGE, pszFile:LPSTR, :ptr CObjectItem
Show@CViewStorageDlg	proto :ptr CViewStorageDlg, :HWND
Destroy@CViewStorageDlg	proto :ptr CViewStorageDlg

;*** CPropertyBag

ifndef INSIDE_CPROPERTYBAG
BEGIN_CLASS CPropertyBag
END_CLASS
endif
Create@CPropertyBag proto :LPSTORAGE, pFmtId:ptr FMTID, bCreateAlways:BOOL, ppPropertyBag:ptr LPPROPERTYBAG

;*** CPropertyStorageDlg

ifndef INSIDE_CPROPERTYSTORAGEDLG
BEGIN_CLASS CPropertyStorageDlg, CDlg
END_CLASS
endif
Create@CPropertyStorageDlg	proto pPropertyStorage:LPPROPERTYSTORAGE, pPropertySetStorage:LPPROPERTYSETSTORAGE
Show@CPropertyStorageDlg	proto :ptr CPropertyStorageDlg, :HWND

;*** CAmbientPropDlg dialog class

ifndef INSIDE_CAMBIENTPROPDLG
BEGIN_CLASS CAmbientPropDlg, CDlg
END_CLASS
endif
Create@CAmbientPropDlg	proto
Show@CAmbientPropDlg	proto :ptr CAmbientPropDlg, hWnd:HWND
Destroy@CAmbientPropDlg	proto :ptr CAmbientPropDlg

;*** static method

GetAmbientProp			proto :DWORD, :ptr VARIANT

;*** end class definitions ***

