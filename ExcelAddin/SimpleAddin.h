// SimpleAddin.h : CSimpleAddin ������

#pragma once
#include "resource.h"       // ������



#include "ExcelAddin_i.h"



#if defined(_WIN32_WCE) && !defined(_CE_DCOM) && !defined(_CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA)
#error "Windows CE ƽ̨(�粻�ṩ��ȫ DCOM ֧�ֵ� Windows Mobile ƽ̨)���޷���ȷ֧�ֵ��߳� COM ���󡣶��� _CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA ��ǿ�� ATL ֧�ִ������߳� COM ����ʵ�ֲ�����ʹ���䵥�߳� COM ����ʵ�֡�rgs �ļ��е��߳�ģ���ѱ�����Ϊ��Free����ԭ���Ǹ�ģ���Ƿ� DCOM Windows CE ƽ̨֧�ֵ�Ψһ�߳�ģ�͡�"
#endif

using namespace ATL;

# include "CAppEvents.h"
# include "CApplication.h"
# include "CWorkbook.h"
# include "CWorksheet.h"

/*
//����VC++�ṩ�Ĺؼ���__uuidof
//������������һ�����ӣ�
//class __declspec(uuid("B372C9F6-1959-4650-960D-73F20CD479BA")) Class;
//struct __declspec(uuid("B372C9F6-1959-4650-960D-73F20CD479BB")) Interface;
//void test()
//{
//CLSID clsid = __uuidof(Class);
//IID iid = __uuidof(Interface);
//...
//}

class ATL_NO_VTABLE CSimpleAddin;

struct DECLSPEC_UUID("00024412-0000-0000-C000-000000000046") IExcelAppEvent;
IID IID_ExcelAppEvent = __uuidof(IExcelAppEvent);

*/

#import "C:\Program Files (x86)\Common Files\microsoft shared\OFFICE15\MSO.DLL" \
	no_implementation \
	rename("RGB", "ExclRGB") \
	rename("DocumentProperties", "ExclDocumentProperties") \
	rename("SearchPath", "ExclSearchPath")

#import "C:\Program Files (x86)\Common Files\microsoft shared\VBA\VBA6\VBE6EXT.OLB" \
	no_implementation

#import "C:\Program Files (x86)\Microsoft Office\Office15\EXCEL.EXE" \
	rename("DialogBox", "ExclDialogBox") \
	rename("RGB", "ExclRGB") \
	rename("CopyFile", "ExclCopyFile") \
	rename("ReplaceText", "ExclReplaceText") \
	exclude("IFont", "IPicture")

_ATL_FUNC_INFO WorkbookOpenInfo = { CC_STDCALL, VT_EMPTY, 1,{VT_DISPATCH | VT_BYREF}};
_ATL_FUNC_INFO WorkbookBeforeCloseInfo = { CC_STDCALL, VT_EMPTY, 2, {{VT_DISPATCH | VT_BYREF}, {VT_BYREF | VT_BOOL}} };


// CSimpleAddin

class ATL_NO_VTABLE CSimpleAddin :
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CSimpleAddin, &CLSID_SimpleAddin>,
	public IDispatchImpl<ISimpleAddin, &IID_ISimpleAddin, &LIBID_ExcelAddinLib, /*wMajor =*/ 1, /*wMinor =*/ 0>,
	public IDispatchImpl<_IDTExtensibility2, &__uuidof(_IDTExtensibility2), &LIBID_AddInDesignerObjects, /* wMajor = */ 1>,
	public IDispEventSimpleImpl<10, CSimpleAddin, &__uuidof(Excel::AppEvents)>,
	public IDispEventSimpleImpl<11, CSimpleAddin, &__uuidof(Excel::AppEvents)>
{
public:
	CSimpleAddin()
	{
	}
	
	DECLARE_REGISTRY_RESOURCEID(IDR_SIMPLEADDIN)


	BEGIN_COM_MAP(CSimpleAddin)
		COM_INTERFACE_ENTRY(ISimpleAddin)
		COM_INTERFACE_ENTRY2(IDispatch, _IDTExtensibility2)
		COM_INTERFACE_ENTRY(_IDTExtensibility2)
	END_COM_MAP()



	DECLARE_PROTECT_FINAL_CONSTRUCT()

	HRESULT FinalConstruct()
	{
		return S_OK;
	}

	void FinalRelease()
	{
	}

public:




	// _IDTExtensibility2 Methods
public:
	STDMETHOD(OnConnection)(LPDISPATCH Application, ext_ConnectMode ConnectMode, LPDISPATCH AddInInst, SAFEARRAY * * custom);
	STDMETHOD(OnDisconnection)(ext_DisconnectMode RemoveMode, SAFEARRAY * * custom);
	STDMETHOD(OnAddInsUpdate)(SAFEARRAY * * custom);
	STDMETHOD(OnStartupComplete)(SAFEARRAY * * custom);
	STDMETHOD(OnBeginShutdown)(SAFEARRAY * * custom);

public:
	typedef IDispEventSimpleImpl<10, CSimpleAddin, &__uuidof(Excel::AppEvents)> WorkbookOpenEvent;//��Excel�¼�
	typedef IDispEventSimpleImpl<11, CSimpleAddin, &__uuidof(Excel::AppEvents)> WorkbookBeforeCloseEvent;//�ر�Excel�¼�

	LPDISPATCH m_Application;

	void _stdcall WorkbookOpen(_In_ LPDISPATCH Wb);
	void _stdcall WorkbookBeforeClose(_In_ LPDISPATCH Wb, _Inout_ VARIANT_BOOL* Cancel);    
	void WorkbookBeforeSave (
        struct _Workbook * Wb,
        VARIANT_BOOL SaveAsUI,
        VARIANT_BOOL * Cancel );

	BEGIN_SINK_MAP(CSimpleAddin)
		SINK_ENTRY_INFO(10, __uuidof(Excel::AppEvents), 0x0000061f, WorkbookOpen, &WorkbookOpenInfo)
		SINK_ENTRY_INFO(11, __uuidof(Excel::AppEvents), 0x00000622, WorkbookBeforeClose, &WorkbookBeforeCloseInfo)
	END_SINK_MAP()
};

OBJECT_ENTRY_AUTO(__uuidof(SimpleAddin), CSimpleAddin)
