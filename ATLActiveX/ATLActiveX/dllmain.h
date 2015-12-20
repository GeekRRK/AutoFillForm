// dllmain.h : 模块类的声明。

class CATLActiveXModule : public ATL::CAtlDllModuleT< CATLActiveXModule >
{
public :
	DECLARE_LIBID(LIBID_ATLActiveXLib)
	DECLARE_REGISTRY_APPID_RESOURCEID(IDR_ATLACTIVEX, "{E955CFC7-9028-4C56-B8DE-164E99EA9C80}")
};

extern class CATLActiveXModule _AtlModule;
