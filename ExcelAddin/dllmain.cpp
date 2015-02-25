// dllmain.cpp : DllMain 的实现。

#include "stdafx.h"
#include "resource.h"
#include "ExcelAddin_i.h"
#include "dllmain.h"
#include "xdlldata.h"

CExcelAddinModule _AtlModule;

class CExcelAddinApp : public CWinApp
{
public:

// 重写
	virtual BOOL InitInstance();
	virtual int ExitInstance();

	DECLARE_MESSAGE_MAP()
};

BEGIN_MESSAGE_MAP(CExcelAddinApp, CWinApp)
END_MESSAGE_MAP()

CExcelAddinApp theApp;

BOOL CExcelAddinApp::InitInstance()
{
#ifdef _MERGE_PROXYSTUB
	if (!PrxDllMain(m_hInstance, DLL_PROCESS_ATTACH, NULL))
		return FALSE;
#endif
	return CWinApp::InitInstance();
}

int CExcelAddinApp::ExitInstance()
{
	return CWinApp::ExitInstance();
}
