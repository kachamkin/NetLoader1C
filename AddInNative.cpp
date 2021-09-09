
#include "stdafx.h"

#ifdef __linux__
#include <unistd.h>
#include <stdlib.h>
#include <signal.h>
#endif

#include <wchar.h>
#include <string>
#include "AddInNative.h"
#include <clocale>

static const wchar_t* g_MethodNames[] = {
    L"CreateObject",
    L"GetObjectType",
    L"InvokeMember"
};

static const wchar_t* g_MethodNamesRu[] = {
    L"СоздатьОбъект",
    L"ПолучитьТипОбъекта",
    L"ВызватьМетодИлиСвойство"
};

static const wchar_t g_kClassNames[] = L"CAddInNative"; //|OtherClass1|OtherClass2";

uint32_t convToShortWchar(WCHAR_T** Dest, const wchar_t* Source, uint64_t len = 0);
uint32_t convFromShortWchar(wchar_t** Dest, const WCHAR_T* Source, uint32_t len = 0);
uint32_t getLenShortWcharStr(const WCHAR_T* Source) noexcept;
static AppCapabilities g_capabilities = eAppCapabilitiesInvalid;
static WcharWrapper s_names(g_kClassNames);

//---------------------------------------------------------------------------//
AppCapabilities SetPlatformCapabilities(const AppCapabilities capabilities)
{
    g_capabilities = capabilities;
    return eAppCapabilitiesLast;
}

//---------------------------------------------------------------------------//
long long GetClassObject(const wchar_t* wsName, IComponentBase** pInterface)
{
    if(!*pInterface)
    {
        *pInterface = new CAddInNative();
        return (long long)*pInterface;
    }
    return 0;
}
//---------------------------------------------------------------------------//
long DestroyObject(IComponentBase** pIntf)
{
   if(!*pIntf)
      return -1;

   delete *pIntf;
   *pIntf = 0;
   return 0;
}
//---------------------------------------------------------------------------//
const WCHAR_T* GetClassNames()
{
    return s_names;
}
//---------------------------------------------------------------------------//
//CAddInNative
CAddInNative::CAddInNative() noexcept
{
    m_iMemory = 0;
    m_iConnect = 0;
}
//---------------------------------------------------------------------------//
CAddInNative::~CAddInNative()
{
}
//---------------------------------------------------------------------------//
bool CAddInNative::Init(void* pConnection)
{
    DWORD dwRetCode = 0;
    
    m_iConnect = static_cast<IAddInDefBase*>(pConnection);

	return m_iConnect != NULL && LoadAssembly();
}
//---------------------------------------------------------------------------//
long CAddInNative::GetInfo() noexcept
{ 
    return 2000; 
}
//---------------------------------------------------------------------------//
void CAddInNative::Done() noexcept
{
}
//---------------------------------------------------------------------------//
bool CAddInNative::RegisterExtensionAs(WCHAR_T** wsExtensionName)
{ 
    const wchar_t* wsExtension = L"NetLoader";
    const size_t iActualSize = ::wcslen(wsExtension) + 1;

    if (m_iMemory)
    {
        if (m_iMemory->AllocMemory((void**)wsExtensionName, (unsigned)iActualSize * sizeof(WCHAR_T)))
            ::convToShortWchar(wsExtensionName, wsExtension, iActualSize);
        return true;
    }

    return false;

}
//---------------------------------------------------------------------------//
long CAddInNative::GetNProps() noexcept
{ 
    return eLastProp;
}
//---------------------------------------------------------------------------//
long CAddInNative::FindProp(const WCHAR_T* wsPropName) noexcept
{ 
    return -1;
}
//---------------------------------------------------------------------------//
const WCHAR_T* CAddInNative::GetPropName(long lPropNum, long lPropAlias) noexcept
{ 
    return 0;
}
//---------------------------------------------------------------------------//
bool CAddInNative::GetPropVal(const long lPropNum, tVariant* pvarPropVal) noexcept
{ 
    return false;
}
//---------------------------------------------------------------------------//
bool CAddInNative::SetPropVal(const long lPropNum, tVariant* varPropVal) noexcept
{ 
    return false;
}
//---------------------------------------------------------------------------//
bool CAddInNative::IsPropReadable(const long lPropNum) noexcept
{ 
    return false;
}
//---------------------------------------------------------------------------//
bool CAddInNative::IsPropWritable(const long lPropNum) noexcept
{
    return false;
}
//---------------------------------------------------------------------------//
long CAddInNative::GetNMethods() noexcept
{ 
    return eMethLast;
}
//---------------------------------------------------------------------------//
long CAddInNative::FindMethod(const WCHAR_T* wsMethodName)
{ 
    long plMethodNum = -1;
    wchar_t* name = 0;

    ::convFromShortWchar(&name, wsMethodName);

    plMethodNum = findName(g_MethodNames, name, eMethLast);

    if (plMethodNum == -1)
        plMethodNum = findName(g_MethodNamesRu, name, eMethLast);

    delete[] name;

    return plMethodNum;
}
//---------------------------------------------------------------------------//
const WCHAR_T* CAddInNative::GetMethodName(const long lMethodNum, 
                            const long lMethodAlias)
{ 
    if (lMethodNum >= eMethLast)
        return NULL;

    wchar_t* wsCurrentName = NULL;
    WCHAR_T* wsMethodName = NULL;
    size_t iActualSize = 0;

    switch (lMethodAlias)
    {
    case 0: // First language
        wsCurrentName = (wchar_t*)g_MethodNames[lMethodNum];
        break;
    case 1: // Second language
        wsCurrentName = (wchar_t*)g_MethodNamesRu[lMethodNum];
        break;
    default:
        return 0;
    }

    iActualSize = wcslen(wsCurrentName) + 1;

    if (m_iMemory && wsCurrentName)
    {
        if (m_iMemory->AllocMemory((void**)&wsMethodName, (unsigned)iActualSize * sizeof(WCHAR_T)))
            ::convToShortWchar(&wsMethodName, wsCurrentName, iActualSize);
    }

    return wsMethodName;

}
//---------------------------------------------------------------------------//
long CAddInNative::GetNParams(const long lMethodNum) noexcept
{ 
    switch (lMethodNum)
    {
    case eMethCreateObject:
        return MAXARGS + 2;
    case eMethInvokeMember:
        return MAXARGS + 4;
    case eMethGetObjectType:
        return 1;
    default:
        return 0;
    }

    return 0;

}
//---------------------------------------------------------------------------//
bool CAddInNative::GetParamDefValue(const long lMethodNum, const long lParamNum,
	tVariant* pvarParamDefValue) noexcept
{
	switch (lMethodNum)
	{
	case eMethCreateObject:
		switch (lParamNum)
		{
		case 0:
            return false;
		case 1:
			TV_VT(pvarParamDefValue) = VTYPE_PWSTR;
			pvarParamDefValue->pwstrVal = L"";
			return true;
		default:
            TV_VT(pvarParamDefValue) = VTYPE_NULL;
			return true;
		}
    case eMethInvokeMember:
        if (lParamNum < 2)
            return false;
        else if (lParamNum == 2)
        {
            TV_VT(pvarParamDefValue) = VTYPE_PWSTR;
            pvarParamDefValue->pwstrVal = L"";
            return true;
        }
        else if (lParamNum == 3)
        {
            TV_VT(pvarParamDefValue) = VTYPE_BOOL;
            pvarParamDefValue->bVal = false;
            return true;
        }
        else
            TV_VT(pvarParamDefValue) = VTYPE_NULL;
            return true;
	default:
		return false;
	}

	return false;

}
//---------------------------------------------------------------------------//
bool CAddInNative::HasRetVal(const long lMethodNum) noexcept
{ 
    switch (lMethodNum)
    {
    case eMethCreateObject:
        return true;
    case eMethGetObjectType:
        return true;
    case eMethInvokeMember:
        return true;
    default:
        return false;
    }

    return false;

}
//---------------------------------------------------------------------------//
bool CAddInNative::CallAsProc(const long lMethodNum,
                    tVariant* paParams, const long lSizeArray) 
{ 
    return false;
}
//---------------------------------------------------------------------------//
bool CAddInNative::CallAsFunc(const long lMethodNum,
                tVariant* pvarRetValue, tVariant* paParams, const long lSizeArray)
{ 
    
    sVariant** pargs = NULL;
    sVariant* res = NULL;
    wchar_t* errText = NULL;
    wchar_t* error = NULL;

    switch (lMethodNum)
    {
    case eMethCreateObject:

        if (lSizeArray != MAXARGS + 2 || !paParams)
            return false;
        
        if (TV_VT(paParams) != VTYPE_PWSTR || TV_VT(paParams + 1) != VTYPE_PWSTR)
        {
            addError(ADDIN_E_VERY_IMPORTANT, L"NETLoader", L"CreateObject: parameter type mismatch.", -1);
            return false;
        }

        TV_VT(pvarRetValue) = VTYPE_I4;

        pargs = new sVariant* [lSizeArray - 2];
        for (int i = 0; i < lSizeArray - 2; i++)
            pargs[i] = ToSVariant(paParams + i + 2);

        pvarRetValue->lVal = CreateObject(paParams->pwstrVal, (paParams + 1)->pwstrVal, pargs, &error);
        if (!pvarRetValue->lVal)
        {
            errText = new wchar_t[wcslen(L"Failed to create object ") + wcslen(paParams->pwstrVal) + 1 + (error ? wcslen(error) + 3 : 0)];
            wcscpy(errText, L"Failed to create object ");
            addError(ADDIN_E_VERY_IMPORTANT, L"NETLoader", wcscat(wcscat(wcscat(errText, paParams->pwstrVal), L":\r\n"), (error ? error : (wchar_t*)L"")), -1);

            delete errText;

            return false;
        }

        delete pargs;

        return true;
    
    case eMethInvokeMember:

        if (lSizeArray != MAXARGS + 4 || !paParams)
            return false;

        if (TV_VT(paParams) != VTYPE_I4 || TV_VT(paParams + 1) != VTYPE_PWSTR || TV_VT(paParams + 2) != VTYPE_PWSTR)
        {
            addError(ADDIN_E_VERY_IMPORTANT, L"NETLoader", L"InvokeNETObjectMember: parameter type mismatch.", -1);
            return false;
        }

        pargs = new sVariant* [lSizeArray - 4];
        for (int i = 0; i < lSizeArray - 4; i++)
            pargs[i] = ToSVariant(paParams + i + 4);

        res = InvokeNETObjectMember(paParams->lVal, (paParams + 1)->pwstrVal, (paParams + 2)->pwstrVal, (paParams + 3)->vt == VTYPE_BOOL ? (paParams + 3)->bVal : false, pargs, &error);
        if (!res || res->vt == VTYPE_ERROR)
        {
            errText = new wchar_t[wcslen(L"Failed to invoke member ") + wcslen((paParams + 1)->pwstrVal) + 1 + (error ? wcslen(error) + 3 : 0)];
            wcscpy(errText, L"Failed to invoke member ");
            addError(ADDIN_E_VERY_IMPORTANT, L"NETLoader", wcscat(wcscat(wcscat(errText, (paParams + 1)->pwstrVal), L":\r\n"), (error ? error : (wchar_t*)L"")), -1);

            delete errText;

            return false;
        }

        FromSVariant(res, pvarRetValue);

        delete pargs;

        return true;

    case eMethGetObjectType:

        if (lSizeArray != 1 || !paParams)
            return false;

        if (TV_VT(paParams) != VTYPE_I4)
        {
            addError(ADDIN_E_VERY_IMPORTANT, L"NETLoader", L"GetNETObjectType: parameter type mismatch.", -1);
            return false;
        }

        TV_VT(pvarRetValue) = VTYPE_PWSTR;
        pvarRetValue->pwstrVal = GetNETObjectType(paParams->lVal, &error);

        if (pvarRetValue->pwstrVal)
        {
            pvarRetValue->wstrLen = (uint32_t)wcslen(pvarRetValue->pwstrVal);
            return true;
        }
        else
        {
            errText = new wchar_t[wcslen(L"Failed to get object type ") + 1 + (error ? wcslen(error) + 3 : 0)];
            wcscpy(errText, L"Failed to get object type ");
            addError(ADDIN_E_VERY_IMPORTANT, L"NETLoader", wcscat(wcscat(errText, L":\r\n"), (error ? error : (wchar_t*)L"")), -1);

            delete errText;

            return false;
        }


    default:
        return false;
    }        
        
     return true;
}

void CAddInNative::FromSVariant(sVariant* psv, tVariant* pv) noexcept
{
    pv->vt = psv->vt;

    if (pv->vt == VTYPE_PWSTR)
    {
        pv->pwstrVal = psv->pwstrVal;
        pv->wstrLen = (uint32_t)wcslen(pv->pwstrVal);
    }
   
    if (pv->vt == VTYPE_TM)
        pv->tmVal = *psv->pdate;

    if (pv->vt == VTYPE_R8)
        pv->dblVal = psv->dblVal;

    if (pv->vt == VTYPE_BOOL)
        pv->bVal = psv->bVal;

    if (pv->vt == VTYPE_I4)
        pv->lVal = psv->lVal;

}

sVariant* CAddInNative::ToSVariant(tVariant* ptVar)
{
    sVariant* psv = new sVariant;

    if (ptVar->vt < 2)
        psv->vt = VTYPE_NULL;
    else if (ptVar->vt < 4 || ptVar->vt > 12 && ptVar->vt < 21)
    {
        psv->vt = VTYPE_I4;
        psv->lVal = ptVar->lVal;
    }
    else if (ptVar->vt == 22)
    {
        psv->vt = VTYPE_PWSTR;
        psv->pwstrVal = ptVar->pwstrVal;
    }
    else if (ptVar->vt == 11)
    {
        psv->vt = VTYPE_BOOL;
        psv->bVal = ptVar->bVal;
    }
    else if (ptVar->vt < 6)
    {
        psv->vt = VTYPE_R8;
        psv->dblVal = ptVar->dblVal;
    }
    else if (ptVar->vt == 7)
    {
        psv->vt = VTYPE_DATE;
        psv->pdate = &(ptVar->tmVal);  
    }

    return psv;
}


//---------------------------------------------------------------------------//
void CAddInNative::SetLocale(const WCHAR_T* loc) noexcept
{
#if !defined( __linux__ ) && !defined(__APPLE__)
    _wsetlocale(LC_ALL, loc);
#else
    //We convert in char* char_locale
    //also we establish locale
    //setlocale(LC_ALL, char_locale);
#endif
}
//---------------------------------------------------------------------------//
bool CAddInNative::setMemManager(void* mem) noexcept
{
    m_iMemory = (IMemoryManager*)mem;
    return m_iMemory != 0;
}
//---------------------------------------------------------------------------//
void CAddInNative::addError(uint32_t wcode, const wchar_t* source,
    const wchar_t* descriptor, long code)
{
    if (m_iConnect)
    {
        WCHAR_T* err = 0;
        WCHAR_T* descr = 0;

        ::convToShortWchar(&err, source);
        ::convToShortWchar(&descr, descriptor);

        m_iConnect->AddError(wcode, err, descr, code);
        delete[] err;
        delete[] descr;
    }
}
//---------------------------------------------------------------------------//
long CAddInNative::findName(const wchar_t* names[], const wchar_t* name,
    const uint32_t size) const noexcept
{
    long ret = -1;
    for (uint32_t i = 0; i < size; i++)
    {
        if (!wcscmp(names[i], name))
        {
            ret = i;
            break;
        }
    }
    return ret;
}
//---------------------------------------------------------------------------//
uint32_t convToShortWchar(WCHAR_T** Dest, const wchar_t* Source, uint64_t len)
{
    if (!len)
        len = ::wcslen(Source)+1;

    if (!*Dest)
        *Dest = new WCHAR_T[len];

    WCHAR_T* tmpShort = *Dest;
    wchar_t* tmpWChar = (wchar_t*) Source;
    uint32_t res = 0;

    ::memset(*Dest, 0, len*sizeof(WCHAR_T));
    do
    {
        *tmpShort++ = (WCHAR_T)*tmpWChar++;
        ++res;
    }
    while (len-- && *tmpWChar);

    return res;
}
//---------------------------------------------------------------------------//
uint32_t convFromShortWchar(wchar_t** Dest, const WCHAR_T* Source, uint32_t len)
{
    if (!len)
        len = getLenShortWcharStr(Source)+1;

    if (!*Dest)
        *Dest = new wchar_t[len];

    wchar_t* tmpWChar = *Dest;
    WCHAR_T* tmpShort = (WCHAR_T*)Source;
    uint32_t res = 0;

    ::memset(*Dest, 0, len*sizeof(wchar_t));
    do
    {
        *tmpWChar++ = (wchar_t)*tmpShort++;
        ++res;
    }
    while (len-- && *tmpShort);

    return res;
}
//---------------------------------------------------------------------------//
uint32_t getLenShortWcharStr(const WCHAR_T* Source) noexcept
{
    uint32_t res = 0;
    WCHAR_T *tmpShort = (WCHAR_T*)Source;

    while (*tmpShort++)
        ++res;

    return res;
}
//---------------------------------------------------------------------------//

#ifdef LINUX_OR_MACOS
WcharWrapper::WcharWrapper(const WCHAR_T* str) : m_str_WCHAR(NULL),
m_str_wchar(NULL)
{
    if (str)
    {
        int len = getLenShortWcharStr(str);
        m_str_WCHAR = new WCHAR_T[len + 1];
        memset(m_str_WCHAR, 0, sizeof(WCHAR_T) * (len + 1));
        memcpy(m_str_WCHAR, str, sizeof(WCHAR_T) * len);
        ::convFromShortWchar(&m_str_wchar, m_str_WCHAR);
    }
}
#endif
//---------------------------------------------------------------------------//
WcharWrapper::WcharWrapper(const wchar_t* str) :
#ifdef LINUX_OR_MACOS
    m_str_WCHAR(NULL),
#endif 
    m_str_wchar(NULL)
{
    if (str)
    {
        const size_t len = wcslen(str);
        m_str_wchar = new wchar_t[len + 1];
        memset(m_str_wchar, 0, sizeof(wchar_t) * (len + 1));
        memcpy(m_str_wchar, str, sizeof(wchar_t) * len);
#ifdef LINUX_OR_MACOS
        ::convToShortWchar(&m_str_WCHAR, m_str_wchar);
#endif
    }

}
//---------------------------------------------------------------------------//
WcharWrapper::~WcharWrapper()
{
#ifdef LINUX_OR_MACOS
    if (m_str_WCHAR)
    {
        delete[] m_str_WCHAR;
        m_str_WCHAR = NULL;
    }
#endif
    if (m_str_wchar)
    {
        delete[] m_str_wchar;
        m_str_wchar = NULL;
    }
}
//---------------------------------------------------------------------------//

