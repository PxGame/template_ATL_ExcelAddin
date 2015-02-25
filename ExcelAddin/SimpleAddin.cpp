// SimpleAddin.cpp : CSimpleAddin µÄÊµÏÖ

#include "stdafx.h"
#include "SimpleAddin.h"


// CSimpleAddin

STDMETHODIMP CSimpleAddin::OnConnection(LPDISPATCH Application, ext_ConnectMode ConnectMode, LPDISPATCH AddInInst, SAFEARRAY * * custom)
{
	OutputDebugString(TEXT("Excel OnConnection"));
	m_Application = Application;
	WorkbookOpenEvent::DispEventAdvise(m_Application, &__uuidof(Excel::AppEvents));
	WorkbookBeforeCloseEvent::DispEventAdvise(m_Application, &__uuidof(Excel::AppEvents));
	return S_OK;
}
STDMETHODIMP CSimpleAddin::OnDisconnection(ext_DisconnectMode RemoveMode, SAFEARRAY * * custom)
{
	OutputDebugString(TEXT("Excel OnDisconnection"));
	WorkbookOpenEvent::DispEventUnadvise(m_Application, &__uuidof(Excel::AppEvents));
	WorkbookBeforeCloseEvent::DispEventUnadvise(m_Application, &__uuidof(Excel::AppEvents));
	return S_OK;
}
STDMETHODIMP CSimpleAddin::OnAddInsUpdate(SAFEARRAY * * custom)
{
	OutputDebugString(TEXT("Excel OnAddInsUpdate"));
	return S_OK;
}
STDMETHODIMP CSimpleAddin::OnStartupComplete(SAFEARRAY * * custom)
{
	OutputDebugString(TEXT("Excel OnStartupComplete"));
	return S_OK;
}
STDMETHODIMP CSimpleAddin::OnBeginShutdown(SAFEARRAY * * custom)
{
	OutputDebugString(TEXT("Excel OnBeginShutdown"));
	return S_OK;
}

void  CSimpleAddin::WorkbookOpen(_In_ LPDISPATCH Wb)
{
	OutputDebugString(TEXT("Excel WorkbookOpen"));
}
void  CSimpleAddin::WorkbookBeforeClose(_In_ LPDISPATCH Wb, _Inout_ VARIANT_BOOL* Cancel)
{
	OutputDebugString(TEXT("Excel WorkbookBeforeClose"));
}