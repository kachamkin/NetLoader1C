
#ifndef __ADDINNATIVE_H__
#define __ADDINNATIVE_H__

#ifndef __linux__
#include <wtypes.h>
#endif //__linux__

#include "ComponentBase.h"
#include "AddInDefBase.h"
#include "IMemoryManager.h"

//������������ ����� ���������� ������������ ������ ��� ������������
#define MAXARGS 5

// ����������� ��� VARIANT
struct _sVariant
{
    wchar_t* pwstrVal;
    ENUMVAR vt;
    int32_t lVal;
    bool bVal;
    double dblVal;
    tm* pdate;
};
typedef _sVariant sVariant;

// ������������� ����������� ���� ��������.
extern "C" bool LoadAssembly();

// �������� ������� ���� pType;
//pargTypes � ���� ���������� ������������ ������� � ������������ ";".
//pargs � ��������� �� ������ ���������� �� ��������� ������������.
extern "C" int CreateObject(wchar_t* pType, wchar_t* pargTypes, sVariant** pargs, wchar_t** error);

//��������� ���������� ������������� ��� ���� ������� (fully qualified, ��� � Type.FullName).
//objectNum � ����� ������ � ���������� ���.
//������������ ��������� LPWSTR
extern "C" wchar_t* GetNETObjectType(int objectNum, wchar_t** error);

//����� ������ ������������ ������� ��� ���������/��������� �������� ��� ��������.
//objectNum � ����� ������� � ���������� ����.
//pname � ��������� LPWSTR �� ������ � ������ ������/��������.
//pargTypes � ���� ���������� ������������ ������� � ������������ ";". ��������� LPWSTR.
//pargs � ��������� �� ������ ���������� �� ��������� ������������ � ������� sVariant.
//setProperty � ���� true, �������� ��������� �������� ��������. �������� ���������� � ������ �������� parg. � ��������� ������ �� �������� ����� ��� �������� �������� ��������.
extern "C" sVariant* InvokeNETObjectMember(int objectNum, wchar_t* pname, wchar_t* pargTypes, bool setProperty, sVariant** pargs, wchar_t** error);

///////////////////////////////////////////////////////////////////////////////
// class CAddInNative
class CAddInNative : public IComponentBase
{
public:
    enum Props
    {
        eLastProp      // Always last
    };

    enum Methods
    {
        
        eMethCreateObject = 0,
        eMethGetObjectType,
        eMethInvokeMember,
        eMethLast      // Always last
    };

    CAddInNative(void) noexcept;
    virtual ~CAddInNative();
    // IInitDoneBase
    virtual bool ADDIN_API Init(void*);
    virtual bool ADDIN_API setMemManager(void* mem) noexcept;
    virtual long ADDIN_API GetInfo() noexcept;
    virtual void ADDIN_API Done() noexcept;
    // ILanguageExtenderBase
    virtual bool ADDIN_API RegisterExtensionAs(WCHAR_T** wsLanguageExt);
    virtual long ADDIN_API GetNProps() noexcept;
    virtual long ADDIN_API FindProp(const WCHAR_T* wsPropName) noexcept;
    virtual const WCHAR_T* ADDIN_API GetPropName(long lPropNum, long lPropAlias) noexcept;
    virtual bool ADDIN_API GetPropVal(const long lPropNum, tVariant* pvarPropVal) noexcept;
    virtual bool ADDIN_API SetPropVal(const long lPropNum, tVariant* varPropVal) noexcept;
    virtual bool ADDIN_API IsPropReadable(const long lPropNum) noexcept;
    virtual bool ADDIN_API IsPropWritable(const long lPropNum) noexcept;
    virtual long ADDIN_API GetNMethods() noexcept;
    virtual long ADDIN_API FindMethod(const WCHAR_T* wsMethodName);
    virtual const WCHAR_T* ADDIN_API GetMethodName(const long lMethodNum, 
                            const long lMethodAlias);
    virtual long ADDIN_API GetNParams(const long lMethodNum) noexcept;
    virtual bool ADDIN_API GetParamDefValue(const long lMethodNum, const long lParamNum,
                            tVariant *pvarParamDefValue) noexcept;
    virtual bool ADDIN_API HasRetVal(const long lMethodNum) noexcept;
    virtual bool ADDIN_API CallAsProc(const long lMethodNum,
                    tVariant* paParams, const long lSizeArray);
    virtual bool ADDIN_API CallAsFunc(const long lMethodNum,
                tVariant* pvarRetValue, tVariant* paParams, const long lSizeArray);
    operator IComponentBase*() noexcept { return this; }
    // LocaleBase
    virtual void ADDIN_API SetLocale(const WCHAR_T* loc) noexcept;
private:
    long findName(const wchar_t* names[], const wchar_t* name, const uint32_t size) const noexcept;
    void addError(uint32_t wcode, const wchar_t* source,
        const wchar_t* descriptor, long code);

    // ���������� ���� VARIANT � ������������
    sVariant* CAddInNative::ToSVariant(tVariant* ptVar);
    // ��������� tVariant �� sVariant
    void FromSVariant(sVariant* psv, tVariant* pv) noexcept;

    // Attributes
    IAddInDefBase* m_iConnect;
    IMemoryManager* m_iMemory;
    
};

class WcharWrapper
{
public:
#ifdef LINUX_OR_MACOS
    WcharWrapper(const WCHAR_T* str);
#endif
    WcharWrapper(const wchar_t* str);
    ~WcharWrapper();
#ifdef LINUX_OR_MACOS
    operator const WCHAR_T* () { return m_str_WCHAR; }
    operator WCHAR_T* () { return m_str_WCHAR; }
#endif
    operator const wchar_t* () { return m_str_wchar; }
    operator wchar_t* () { return m_str_wchar; }
private:
    WcharWrapper& operator = (const WcharWrapper& other) { return *this; }
    WcharWrapper(const WcharWrapper& other) { }
private:
#ifdef LINUX_OR_MACOS
    WCHAR_T* m_str_WCHAR;
#endif
    wchar_t* m_str_wchar;
};

#endif //__ADDINNATIVE_H__
