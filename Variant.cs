using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace _1CAddIn
{
    public enum ENUMVAR
    {
        VTYPE_EMPTY = 0,
        VTYPE_NULL,
        VTYPE_I2,                   //int16_t
        VTYPE_I4,                   //int32_t
        VTYPE_R4,                   //float
        VTYPE_R8,                   //double
        VTYPE_DATE,                 //DATE (double)
        VTYPE_TM,                   //struct tm
        VTYPE_PSTR,                 //struct str    string
        VTYPE_INTERFACE,            //struct iface
        VTYPE_ERROR,                //int32_t errCode
        VTYPE_BOOL,                 //bool
        VTYPE_VARIANT,              //struct _tVariant *
        VTYPE_I1,                   //int8_t
        VTYPE_UI1,                  //uint8_t
        VTYPE_UI2,                  //uint16_t
        VTYPE_UI4,                  //uint32_t
        VTYPE_I8,                   //int64_t
        VTYPE_UI8,                  //uint64_t
        VTYPE_INT,                  //int   Depends on architecture
        VTYPE_UINT,                 //unsigned int  Depends on architecture
        VTYPE_HRESULT,              //long hRes
        VTYPE_PWSTR,                //struct wstr
        VTYPE_BLOB,                 //means in struct str binary data contain
        VTYPE_CLSID,                //UUID
        VTYPE_STR_BLOB = 0xfff,
        VTYPE_VECTOR = 0x1000,
        VTYPE_ARRAY = 0x2000,
        VTYPE_BYREF = 0x4000,    //Only with struct _tVariant *
        VTYPE_RESERVED = 0x8000,
        VTYPE_ILLEGAL = 0xffff,
        VTYPE_ILLEGALMASKED = 0xfff,
        VTYPE_TYPEMASK = 0xfff
    };

    public struct tm
    {
        public int tm_sec;   // seconds after the minute - [0, 60] including leap second
        public int tm_min;   // minutes after the hour - [0, 59]
        public int tm_hour;  // hours since midnight - [0, 23]
        public int tm_mday;  // day of the month - [1, 31]
        public int tm_mon;   // months since January - [0, 11]
        public int tm_year;  // years since 1900
        public int tm_wday;  // days since Sunday - [0, 6]
        public int tm_yday;  // days since January 1 - [0, 365]
        public int tm_isdst; // daylight savings time flag
    };

    // Укороченный VARIANT.
    // Аргументы этого типа передаются в управляемые методы из внешней компоненты. 
    public struct sVariant
    {
        public long pwstrVal;
        public ENUMVAR vt;
        public Int32 lVal; //целое число или номер объекта в шлобальном кэше (в зависимости от типа аргумента)
        public bool bVal;
        public double dblVal;
        public long pdate; //указатель на tm
    };

}