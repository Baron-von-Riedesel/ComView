
;--- Generated .INC file by COMView
;--- typelib name=MSDATASRC
;--- documentation=Microsoft Data Source Interfaces
;--- GUID={7C0FFAB0-CD84-11D0-949A-00A0C91110ED}
;--- Version 1.0, LCID 0

;--- prototypes too complex for MASM may exist
;--- if such ones reside in a vtable they will be replaced by dummies
;--- search for occurances of ';+++' to have a closer look at these

sTLBID_MSDATASRC textequ <GUID {07C0FFAB0h,0CD84h,011D0h,{094h,09Ah,000h,0A0h,0C9h,011h,010h,0EDh}}>
_MajorVer_MSDATASRC equ 1
_MinorVer_MSDATASRC equ 0

;--- DataMember, Type=ALIAS, GUID={7C0FFAB1-CD84-11D0-949A-00A0C91110ED}

DataMember typedef BSTR

;--- DataSourceListener, Type=INTERFACE, GUID={7C0FFAB2-CD84-11D0-949A-00A0C91110ED}

ifndef DataSourceListenerVtbl
BEGIN_INTERFACE DataSourceListener, IUnknown
	STDMETHOD	dataMemberChanged	, :DataMember
	STDMETHOD	dataMemberAdded	, :DataMember
	STDMETHOD	dataMemberRemoved	, :DataMember
END_INTERFACE
endif

sIID_DataSourceListener textequ <IID {07C0FFAB2h,0CD84h,011D0h,{094h,09Ah,000h,0A0h,0C9h,011h,010h,0EDh}}>

;--- DataSource, Type=INTERFACE, GUID={7C0FFAB3-CD84-11D0-949A-00A0C91110ED}

ifndef DataSourceVtbl
BEGIN_INTERFACE DataSource, IUnknown
	STDMETHOD	getDataMember	, :DataMember,:ptr GUID,:ptr LPUNKNOWN
	STDMETHOD	getDataMemberName	, :SDWORD,:ptr DataMember
	STDMETHOD	getDataMemberCount	, :ptr SDWORD
	STDMETHOD	addDataSourceListener	, :ptr DataSourceListener
	STDMETHOD	removeDataSourceListener	, :ptr DataSourceListener
END_INTERFACE
endif

sIID_DataSource textequ <IID {07C0FFAB3h,0CD84h,011D0h,{094h,09Ah,000h,0A0h,0C9h,011h,010h,0EDh}}>

;--- end of type library ---
