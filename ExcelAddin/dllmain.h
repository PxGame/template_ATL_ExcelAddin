// dllmain.h : 模块类的声明。

class CExcelAddinModule : public ATL::CAtlDllModuleT< CExcelAddinModule >
{
public :
	DECLARE_LIBID(LIBID_ExcelAddinLib)
	DECLARE_REGISTRY_APPID_RESOURCEID(IDR_EXCELADDIN, "{5DBE0B9E-6AA4-4718-89E9-4B64196821D4}")
};

extern class CExcelAddinModule _AtlModule;
