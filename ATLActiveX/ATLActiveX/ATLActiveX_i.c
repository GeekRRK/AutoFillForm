

/* this ALWAYS GENERATED file contains the IIDs and CLSIDs */

/* link this file in with the server and any clients */


 /* File created by MIDL compiler version 7.00.0555 */
/* at Mon Jun 01 09:55:40 2015
 */
/* Compiler settings for ATLActiveX.idl:
    Oicf, W1, Zp8, env=Win32 (32b run), target_arch=X86 7.00.0555 
    protocol : dce , ms_ext, c_ext, robust
    error checks: allocation ref bounds_check enum stub_data 
    VC __declspec() decoration level: 
         __declspec(uuid()), __declspec(selectany), __declspec(novtable)
         DECLSPEC_UUID(), MIDL_INTERFACE()
*/
/* @@MIDL_FILE_HEADING(  ) */

#pragma warning( disable: 4049 )  /* more than 64k source lines */


#ifdef __cplusplus
extern "C"{
#endif 


#include <rpc.h>
#include <rpcndr.h>

#ifdef _MIDL_USE_GUIDDEF_

#ifndef INITGUID
#define INITGUID
#include <guiddef.h>
#undef INITGUID
#else
#include <guiddef.h>
#endif

#define MIDL_DEFINE_GUID(type,name,l,w1,w2,b1,b2,b3,b4,b5,b6,b7,b8) \
        DEFINE_GUID(name,l,w1,w2,b1,b2,b3,b4,b5,b6,b7,b8)

#else // !_MIDL_USE_GUIDDEF_

#ifndef __IID_DEFINED__
#define __IID_DEFINED__

typedef struct _IID
{
    unsigned long x;
    unsigned short s1;
    unsigned short s2;
    unsigned char  c[8];
} IID;

#endif // __IID_DEFINED__

#ifndef CLSID_DEFINED
#define CLSID_DEFINED
typedef IID CLSID;
#endif // CLSID_DEFINED

#define MIDL_DEFINE_GUID(type,name,l,w1,w2,b1,b2,b3,b4,b5,b6,b7,b8) \
        const type name = {l,w1,w2,{b1,b2,b3,b4,b5,b6,b7,b8}}

#endif !_MIDL_USE_GUIDDEF_

MIDL_DEFINE_GUID(IID, IID_IComponentRegistrar,0xa817e7a2,0x43fa,0x11d0,0x9e,0x44,0x00,0xaa,0x00,0xb6,0x77,0x0a);


MIDL_DEFINE_GUID(IID, IID_IIEControl,0xB222902F,0x02B5,0x48F8,0x81,0x7F,0xF5,0x89,0xB5,0xE1,0x7A,0x2C);


MIDL_DEFINE_GUID(IID, LIBID_ATLActiveXLib,0x95FA6526,0x05F8,0x4C7B,0xAE,0x9B,0x5E,0x5D,0x99,0xBC,0x6B,0x82);


MIDL_DEFINE_GUID(CLSID, CLSID_CompReg,0xBE76B85A,0x1493,0x4248,0x8E,0x09,0x8F,0xA5,0xE0,0x14,0x2A,0x38);


MIDL_DEFINE_GUID(IID, DIID__IIEControlEvents,0x753C80F8,0xA84B,0x410F,0xB7,0xA0,0xE9,0xF4,0x17,0xB2,0xF7,0x7F);


MIDL_DEFINE_GUID(CLSID, CLSID_IEControl,0x42CC3B43,0xC167,0x458C,0x8E,0x9A,0x7C,0xCD,0x0E,0x12,0xC5,0x9A);

#undef MIDL_DEFINE_GUID

#ifdef __cplusplus
}
#endif



