// IEControl.cpp : CIEControl 的实现
#include "stdafx.h"
#include "IEControl.h"

// CIEControl
#include <comdef.h>
#include <string>

//STDMETHODIMP CIEControl::test(void)
//{
//	CallJavaScript(L"alert(\"hello\")");
//
//	// TODO: 在此添加实现代码
//	CComPtr<IServiceProvider> isp;
//    CComPtr<IWebBrowser2> ppBrowser;
//    CComPtr<IDispatch> pDispDoc;
//    CComPtr<IHTMLDocument2> pDocument2;
//    HRESULT hr = S_OK;
//    DISPID dispid;
//    CComVariant avarParams[1];
//    avarParams[0].vt = VT_UI1;
//    DISPPARAMS dispparams = {avarParams, NULL, 1, 0};
//    EXCEPINFO excepInfo;
//    memset(&excepInfo, 0, sizeof excepInfo);
//    CComVariant vaResult;
//    UINT nArgErr = (UINT)-1;// initialize to invalid arg
//
//    //获取站点
//	CComPtr<IOleClientSite> spClientSite;
//	hr = GetClientSite(&spClientSite); 
//
//    //获取服务提供者
//    hr = spClientSite->QueryInterface(IID_IServiceProvider,reinterpret_cast<void **>(&isp));  
//
//    //获取IE控件
//    hr = isp->QueryService(IID_IWebBrowserApp,   IID_IWebBrowser2,   reinterpret_cast<void **>(&ppBrowser));  
//
//    //获取当前页面对象
//    hr = ppBrowser->get_Document(&pDispDoc);  
//
//    //获取页面内容
//    hr = pDispDoc->QueryInterface(IID_IHTMLDocument2 ,reinterpret_cast<void **>(&pDocument2));
//
//	CComQIPtr< IHTMLFramesCollection2> spFrameCollection;
//	hr = pDocument2->get_frames(&spFrameCollection);
//
//    long nFrameCount=0;//取得表单数目    
//    hr = spFrameCollection->get_length( &nFrameCount );    
//
//	for(long i = 0; i < nFrameCount; ++i)
//	{
//		CComVariant vDispWin2; //取得子框架的自动化接口   
//        hr = spFrameCollection->item(&CComVariant(i), &vDispWin2);    
//        if ( FAILED( hr ) )continue;    
//          
//		CComQIPtr<IHTMLWindow2>spWin2 = vDispWin2.pdispVal;
//		if (!spWin2) continue; //取得子框架的   IHTMLWindow2   接口
//		CComPtr <IHTMLLocation> spLoc;
//		spWin2->get_location(&spLoc);//获取frame的页面地址
//		BSTR bHref;
//		spLoc->get_href(&bHref);//获取链接地址
//
//		MessageBox(bHref);
//
//		if(!wcscmp(bHref, L"http://219.148.31.135:8181/examples/iframe.html"))
//		{ 
//			//效能提升门户中间框架middleFrame的页面地址
//			CComPtr <IHTMLDocument2> spDoc2;   
//			spWin2->get_document(&spDoc2); //取得子框架的   IHTMLDocument2   接口
//			CComQIPtr< IHTMLElementCollection > spElementCollection;    
//			hr = spDoc2->get_forms( &spElementCollection );//取得表单集合  
//
//			long nFormCount=0;//取得表单数目    
//			hr = spElementCollection->get_length( &nFormCount );
//
//			for(long i=0; i<nFormCount; i++)  //遍历表单  
//			{    
//				
//				IDispatch *pDisp = NULL;//取得第 i 项表单    
//				hr = spElementCollection->item( CComVariant( i ), CComVariant(), &pDisp );    
//				if ( FAILED( hr ) )continue;    
//
//				CComQIPtr< IHTMLFormElement > spFormElement = pDisp;    
//				pDisp->Release();    
//
//				long nElemCount=0;//取得表单中 域 的数目    
//				hr = spFormElement->get_length( &nElemCount );    
//				if ( FAILED( hr ) )continue;    
//
//				for(long j=0; j<nElemCount; j++)    
//				{    
//					CComDispatchDriver spInputElement;//取得第 j 项表单域    
//					CComVariant vName,vVal,vType;//取得表单域的名，值，类型    
//					hr = spFormElement->item( CComVariant( j ), CComVariant(), &spInputElement );    
//					if ( FAILED( hr ) )continue;    
//					hr = spInputElement.GetPropertyByName(L"name", &vName);    
//					if(vName == (CComVariant)"user")  
//					{  
//						vVal = "admin";  
//						spInputElement.PutPropertyByName(L"value",&vVal);  
//					}  
//					if(vName == (CComVariant)"password")  
//					{  
//						vVal = "123";  
//						spInputElement.PutPropertyByName(L"value",&vVal);  
//					}
//				}
//			} 
//			break;
//		}
//	}
//
//	return S_OK;
//}
//
//LRESULT CIEControl::CallJavaScript( LPCWSTR lpJavaScriptCode )
//{
//	CComPtr<IServiceProvider> isp;
//    CComPtr<IWebBrowser2> ppBrowser;
//    CComPtr<IDispatch> pDispDoc;
//    CComPtr<IHTMLDocument2> pDocument2;
//    HRESULT hr = S_OK;
//    DISPID dispid;
//    CComVariant avarParams[1];
//    avarParams[0].vt = VT_UI1;
//    DISPPARAMS dispparams = {avarParams, NULL, 1, 0};
//    EXCEPINFO excepInfo;
//    memset(&excepInfo, 0, sizeof excepInfo);
//    CComVariant vaResult;
//    UINT nArgErr = (UINT)-1;// initialize to invalid arg
//
//    //获取站点
//	CComPtr<IOleClientSite> spClientSite;
//	hr = GetClientSite(&spClientSite); 
//
//    //获取服务提供者
//    hr = spClientSite->QueryInterface(IID_IServiceProvider,reinterpret_cast<void **>(&isp));  
//
//    //获取IE控件
//    hr = isp->QueryService(IID_IWebBrowserApp,   IID_IWebBrowser2,   reinterpret_cast<void **>(&ppBrowser));  
//
//    //获取当前页面对象
//    hr = ppBrowser->get_Document(&pDispDoc); 
//
//	//获取页面内容
//    hr = pDispDoc->QueryInterface(IID_IHTMLDocument2 ,reinterpret_cast<void **>(&pDocument2));
//
//    if (!pDocument2)
//        return E_FAIL;
//    CComPtr<IHTMLWindow2> pWindow;
//    if (FAILED(pDocument2->get_parentWindow(&pWindow)))
//    {
//        return E_FAIL;
//    }
// 
//    CComVariant bvResult;
//    static CComBSTR javascript = L"javascript";
// 
//	char strSafeJsCode[200] = {0};
//	sprintf(strSafeJsCode, "%s", "try{ %s } catch(e){ }");
//	lpJavaScriptCode = CComBSTR(strSafeJsCode);
//
//    CComBSTR bstrSaveJsCode(lpJavaScriptCode);
//    hr = pWindow->execScript(bstrSaveJsCode, javascript, &bvResult);
//
//    return hr;
//}
//
//STDMETHODIMP CIEControl::CallWebJs(VARIANT scriptCallback)
//{
//	// TODO: 在此添加实现代码
//
//	CComPtr<IDispatch> spCallback;
//	if (scriptCallback.vt == VT_DISPATCH)
//		spCallback = scriptCallback.pdispVal;
//
//	CComVariant avarParams[1];
//
//	avarParams[0] = "hhheeee"; //指定回调函数的参数
//
//	DISPPARAMS params = { avarParams, NULL, 1, 0 };
//
//	if(spCallback)
//
//		spCallback->Invoke(0, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
//
//	return S_OK;
//}

//STDMETHODIMP CIEControl::test(void)
//{
//	// TODO: 在此添加实现代码
//	CComPtr<IServiceProvider> isp;
//    CComPtr<IWebBrowser2> ppBrowser;
//    CComPtr<IDispatch> pDispDoc;
//    CComPtr<IHTMLDocument2> pDocument2;
//    HRESULT hr = S_OK;
//    DISPID dispid;
//    CComVariant avarParams[1];
//    avarParams[0].vt = VT_UI1;
//    DISPPARAMS dispparams = {avarParams, NULL, 1, 0};
//    EXCEPINFO excepInfo;
//    memset(&excepInfo, 0, sizeof excepInfo);
//    CComVariant vaResult;
//    UINT nArgErr = (UINT)-1;// initialize to invalid arg
//
//    //获取站点
//	CComPtr<IOleClientSite> spClientSite;
//	hr = GetClientSite(&spClientSite); 
//
//    //获取服务提供者
//    hr = spClientSite->QueryInterface(IID_IServiceProvider,reinterpret_cast<void **>(&isp));  
//
//    //获取IE控件
//    hr = isp->QueryService(IID_IWebBrowserApp,   IID_IWebBrowser2,   reinterpret_cast<void **>(&ppBrowser));  
//
//    //获取当前页面对象
//    hr = ppBrowser->get_Document(&pDispDoc);  
//
//    //获取页面内容
//    hr = pDispDoc->QueryInterface(IID_IHTMLDocument2 ,reinterpret_cast<void **>(&pDocument2));
//
//	CComQIPtr< IHTMLFramesCollection2> spFrameCollection;
//	hr = pDocument2->get_frames(&spFrameCollection);
//
//    long nFrameCount=0;//取得iframe数目    
//    hr = spFrameCollection->get_length( &nFrameCount );
//
//	for(long i = 0; i < nFrameCount; ++i)
//	{
//		CComVariant vDispWin2; //取得子框架的自动化接口   
//        hr = spFrameCollection->item(&CComVariant(i), &vDispWin2);    
//        if ( FAILED( hr ) )continue;    
//          
//		CComQIPtr<IHTMLWindow2>spWin2 = vDispWin2.pdispVal; //取得子框架的   IHTMLWindow2   接口
//		if (!spWin2)continue;
//
//		CComPtr <IHTMLLocation> spLoc = 0;
//		spWin2->get_location(&spLoc);//获取frame的页面地址
//		if(spLoc == 0)
//		{
//			MessageBox(L"没有获得iframe的页面地址！");
//		}
//
//		BSTR bHref = 0;
//		spLoc->get_href(&bHref);//获取链接地址
//		if(bHref == 0)
//		{
//			MessageBox(L"没有获得iframe的链接地址！");
//		}
//		MessageBox(bHref);
//
//		if(/*!wcscmp(bHref, L"http://wsbs.he-n-tax.gov.cn/")*/true)
//		{ 
//			CComPtr <IHTMLDocument2> spDoc2 = 0;   
//			spWin2->get_document(&spDoc2);
//			if(spDoc2 == 0)
//			{
//				MessageBox(L"没有获得iframe的document！");
//			}
//
//			CComQIPtr< IHTMLElementCollection > spElementCollection;    
//			hr = spDoc2->get_all( &spElementCollection );//取得元素集合
//
//			long nCount=0;//取得元素数目    
//			hr = spElementCollection->get_length( &nCount );
//
//			char buf[20] = {0};
//			itoa(nCount, buf, 10);
//			MessageBox(CComBSTR(buf));
//
//			for(long i=0; i<nCount; i++)  //遍历元素
//			{  
//				IDispatch *pDisp = NULL;//取得第 i 项表单    
//				hr = spElementCollection->item(CComVariant(i),CComVariant(0),&pDisp); //获取单个元素
//				if(SUCCEEDED(hr)) 
//				{
//					//CComQIPtr <IHTMLElement, &IID_IHTMLElement> pElement(pDisp);
//					CComQIPtr<IHTMLElement, &IID_IHTMLElement> pElement;
//					pDisp->QueryInterface(&pElement);
//					BSTR bId = 0;
//					pElement->get_id(&bId);//可以获取其他特征，根据具体元素而定_
//					if(bId == 0)
//					{
//						continue;
//					}
//
//					if(!wcscmp(bId, L"nsrsbh$text"))//根据id是主叫号码获取值或作其他处理
//					{
//						MessageBox(L"nsrsbh$text");
//						IHTMLInputTextElement* input;
//						pDisp->QueryInterface(IID_IHTMLInputTextElement,(void**)&input);
//						BSTR bVal;
//						input->put_value(L"123");
//					}
//				}
//			} 
//			break;
//		}
//	}
//
//	return S_OK;
//}

STDMETHODIMP CIEControl::test(void)
{
	// TODO: 在此添加实现代码
	CComPtr<IServiceProvider> isp;
    CComPtr<IWebBrowser2> ppBrowser;
    CComPtr<IDispatch> pDispDoc;
    CComPtr<IHTMLDocument2> pDocument2;
    HRESULT hr = S_OK;
    DISPID dispid;
    CComVariant avarParams[1];
    avarParams[0].vt = VT_UI1;
    DISPPARAMS dispparams = {avarParams, NULL, 1, 0};
    EXCEPINFO excepInfo;
    memset(&excepInfo, 0, sizeof excepInfo);
    CComVariant vaResult;
    UINT nArgErr = (UINT)-1;// initialize to invalid arg

    //获取站点
	CComPtr<IOleClientSite> spClientSite;
	hr = GetClientSite(&spClientSite); 

    //获取服务提供者
    hr = spClientSite->QueryInterface(IID_IServiceProvider,reinterpret_cast<void **>(&isp));  

    //获取IE控件
    hr = isp->QueryService(IID_IWebBrowserApp,   IID_IWebBrowser2,   reinterpret_cast<void **>(&ppBrowser));  

    //获取当前页面对象
    hr = ppBrowser->get_Document(&pDispDoc);  

    //获取页面内容
    hr = pDispDoc->QueryInterface(IID_IHTMLDocument2 ,reinterpret_cast<void **>(&pDocument2));

	EnumFrame(pDocument2, "");

	return S_OK;
}


void CIEControl::EnumFrame(CComPtr<IHTMLDocument2> spDoc2, CComBSTR innerCode)
{
	HRESULT hr = 0;
	CComQIPtr< IHTMLElementCollection > spElementCollection;    
	spDoc2->get_all( &spElementCollection );//取得元素集合

	long nCount=0;//取得元素数目    
	spElementCollection->get_length( &nCount );

	char buf[20] = {0};
	itoa(nCount, buf, 10);

	for(long i=0; i<nCount; i++)  //遍历元素
	{  
		IDispatch *pDisp = NULL;//取得第 i 项表单    
		spElementCollection->item(CComVariant(i),CComVariant(0),&pDisp); //获取单个元素
		if(SUCCEEDED(hr)) 
		{
			//CComQIPtr <IHTMLElement, &IID_IHTMLElement> pElement(pDisp);
			CComQIPtr<IHTMLElement, &IID_IHTMLElement> pElement;
			pDisp->QueryInterface(&pElement);
			BSTR bId = 0;
			pElement->get_id(&bId);//可以获取其他特征，根据具体元素而定_
			if(bId == 0)
			{
				continue;
			}

			if(!wcscmp(bId, L"taxnum"))//根据id是主叫号码获取值或作其他处理
			{
				IHTMLInputTextElement* input;
				pDisp->QueryInterface(IID_IHTMLInputTextElement,(void**)&input);
				BSTR bVal;
				input->get_value(&bVal);

				CComPtr<IHTMLFramesCollection2> pFrames;
				hr = spDoc2->get_frames(&pFrames);
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

						//CComPtr <IHTMLLocation> spLoc = 0;
						//spWnd2->get_location(&spLoc);//获取frame的页面地址
						//if(spLoc == 0)
						//{
						//	MessageBox(L"没有获得iframe的页面地址！");
						//}

						//BSTR bHref = 0;
						//spLoc->get_href(&bHref);//获取链接地址
						//if(bHref == 0)
						//{
						//	MessageBox(L"没有获得iframe的链接地址！");
						//}
						//MessageBox(bHref);

						CComPtr<IHTMLDocument2> pSubDoc2;
						hr = spWnd2->get_document(&pSubDoc2);
						if (hr == S_OK) {
							//ExecJs(pSubDoc2, innerCode);
							//EnumFrame(pSubDoc2, innerCode);
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
						//ExecJs(pSubDoc2, innerCode);
						//EnumFrame(pSubDoc2, innerCode);

						CComQIPtr< IHTMLElementCollection > spElementCollection;    
						hr = pSubDoc2->get_all( &spElementCollection );//取得元素集合

						long nCount=0;//取得元素数目    
						hr = spElementCollection->get_length( &nCount );

						char buf[20] = {0};
						itoa(nCount, buf, 10);

						for(long i=0; i<nCount; i++)  //遍历元素
						{  
							IDispatch *pDisp = NULL;//取得第 i 项表单    
							hr = spElementCollection->item(CComVariant(i),CComVariant(0),&pDisp); //获取单个元素
							if(SUCCEEDED(hr)) 
							{
								//CComQIPtr <IHTMLElement, &IID_IHTMLElement> pElement(pDisp);
								CComQIPtr<IHTMLElement, &IID_IHTMLElement> pElement;
								pDisp->QueryInterface(&pElement);
								BSTR bId = 0;
								pElement->get_id(&bId);//可以获取其他特征，根据具体元素而定
								if(bId == 0)
								{
									continue;
								}

								if(!wcscmp(bId, L"nsrsbhCa$text"))//根据id是主叫号码获取值或作其他处理
								{
									IHTMLInputTextElement* input;
									pDisp->QueryInterface(IID_IHTMLInputTextElement,(void**)&input);
									input->put_value(bVal);

									return;
								}
							}
						} 
					} 
				}
			}
		}
	}

	return;
}

