
;*** CEditDlg class dummy definition

KEYPAIR		struct
pszRoot		dd ?
pszKey		dd ?
bExpand		BOOL ?
KEYPAIR		ends

EDITITEMDESC struct
pszKey		LPSTR ?
dwCookie	DWORD ?
EDITITEMDESC ends

LPKEYPAIR	typedef ptr KEYPAIR
ifndef HTREEITEM
HTREEITEM typedef HANDLE
endif

ifndef INSIDE_CEditDlg
CEditDlg struct
CEditDlg ends
endif


;*** public methods

Create@CEditDlg			proto hWndParent:HWND, bModeless:BOOL, bConfirmDelete:BOOL
Destroy@CEditDlg		proto :ptr CEditDlg
Show@CEditDlg			proto :ptr CEditDlg
SetCookie@CEditDlg		proto :ptr CEditDlg, dwCookie:DWORD
SetKeys@CEditDlg		proto :ptr CEditDlg, numPairs:DWORD,:ptr KEYPAIR
SetRoot@CEditDlg		proto :ptr CEditDlg, hRoot:HANDLE

