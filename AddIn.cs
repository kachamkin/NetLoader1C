using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Reflection;

namespace _1CAddIn
{
    public static class AddIn
    {
        //Максимальное число аргументов управляемого метода или конструктора
        const int MAXARGS = 5;
                
        //Глобальный кэш объектов 
        static List<object> global;

        static bool isEnabled = false;

        //Инициализация. Вызывается из внешней компоненты при ее загрузке.
        [DllExport]
        public static bool LoadAssembly()
        {
            if (isEnabled)
                return false;
            
            isEnabled = true;
            global = new List<object>();

            return true;
        }

        //Создание объекта типа ptype и помещение его в глобальный кэш; возвращается номер объекта в кэше.
        //pargTypes — типы аргументов конструктора строкой с разделителем ";". Указатель LPWSTR.
        //parg — указатель на массив указателей на аргументы конструктора в формате sVariant.
        //Вызывается из внешней компоненты.
        [DllExport]
        public static int CreateObject(long ptype, long pargTypes, long parg, ref long perror)
        {

            if (!isEnabled)
            {
                perror = (long)Marshal.StringToCoTaskMemAuto("CreateObject: assembly is not initialized!");
                return 0;
            }

            try
            {

                string type = Marshal.PtrToStringAuto((IntPtr)ptype);
                string argTypes = Marshal.PtrToStringAuto((IntPtr)pargTypes);

                Type t = Type.GetType(type);
                if (!IsObject(t))
                {
                    perror = (long)Marshal.StringToCoTaskMemAuto("CreateObject: type \"" + type + "\" is not an object type and cannot be instantiated!");
                    return 0;
                }

                Type[] t_argTypes = ParseTypes(argTypes);
                if (t_argTypes == null)
                {
                    perror = (long)Marshal.StringToCoTaskMemAuto("CreateObject: failed to parse argument types string \"" + argTypes + "\"!");
                    return 0;
                }
                if (t_argTypes.Length > MAXARGS)
                {
                    perror = (long)Marshal.StringToCoTaskMemAuto("CreateObject: number of arguments is greater than max possible (max = " + MAXARGS + "!");
                    return 0;
                }

                long[] pargs = new long[t_argTypes.Length];
                Marshal.Copy((IntPtr)parg, pargs, 0, t_argTypes.Length);

                object[] args = new object[t_argTypes.Length];
                for (int i = 0; i < t_argTypes.Length; i++)
                    args[i] = FromSVariant(pargs[i], t_argTypes[i]);

                ConstructorInfo constructor = t.GetConstructor(t_argTypes);

                if (constructor == null)
                {
                    perror = (long)Marshal.StringToCoTaskMemAuto("CreateObject: failed to get constructor for type \"" + type + "\" with arguments of types \"" + argTypes + "\"!");
                    return 0;
                }

                global.Add(constructor.Invoke(args));
                return global.Count;

            }
            catch (Exception ex)
            {
                perror = (long)Marshal.StringToCoTaskMemAuto("CreateObject: exception occured!\r\n\r\n" + ex);
                return 0;
            }

        }

        //Вызов метода управляемого объекта или получение/установка значения его свойства.
        //objectNum — номер объекта в глобальном кэше.
        //pName — указатель LPWSTR на строку с именем метода/свойства.
        //pargTypes — типы аргументов конструктора строкой с разделителем ";". Указатель LPWSTR.
        //parg — указатель на массив указателей на аргументы конструктора в формате sVariant.
        //setProperty — если true, вызываем установку значения свойства. Значение передается в первом элементе parg. В противном случае мы вызываем метод или получаем значение свойства.
        //Вызывается из внешней компоненты.
        [DllExport]
        public static long InvokeNETObjectMember(int objectNum, long pname, long pargTypes, bool setProperty, long parg, ref long perror)
        {
            if (!isEnabled)
            {
                perror = (long)Marshal.StringToCoTaskMemAuto("InvokeNETObjectMember: assembly is not initialized!");
                return 0;
            }

            sVariant v = new sVariant();
            v.vt = ENUMVAR.VTYPE_ERROR;

            try
            {

                string argTypes = Marshal.PtrToStringAuto((IntPtr)pargTypes);
                Type[] t_argTypes = ParseTypes(argTypes);
                if (t_argTypes == null)
                {
                    perror = (long)Marshal.StringToCoTaskMemAuto("InvokeNETObjectMember: failed to parse argument types string \"" + argTypes + "\"!");
                    return 0;
                }
                if (t_argTypes.Length > MAXARGS)
                {
                    perror = (long)Marshal.StringToCoTaskMemAuto("InvokeNETObjectMember: number of arguments is greater than max possible (max = " + MAXARGS + "!");
                    return 0;
                }

                long[] pargs = new long[t_argTypes.Length];
                Marshal.Copy((IntPtr)parg, pargs, 0, t_argTypes.Length);

                object[] args = new object[t_argTypes.Length];
                for (int i = 0; i < t_argTypes.Length; i++)
                    args[i] = FromSVariant(pargs[i], t_argTypes[i]);

                string name = Marshal.PtrToStringAuto((IntPtr)pname);

                object obj = global[objectNum - 1];
                object res = obj.GetType().InvokeMember(name, BindingFlags.DeclaredOnly | BindingFlags.Instance | BindingFlags.Public | (setProperty ? BindingFlags.SetProperty : BindingFlags.GetProperty | BindingFlags.InvokeMethod), null, obj, args);
                if (res != null && IsObject(res.GetType()))
                    global.Add(res);

                ToSVariant(ref v, res);

                IntPtr ptr = Marshal.AllocCoTaskMem(Marshal.SizeOf(v));
                Marshal.StructureToPtr(v, ptr, true);

                return (long)ptr;

            }
            catch (Exception ex)
            {
                perror = (long)Marshal.StringToCoTaskMemAuto("InvokeNETObjectMember: exception occured!\r\n\r\n" + ex);
                return 0;
            }
        }

        //Приведение управляемого типа к sVariant.
        static void ToSVariant(ref sVariant v, object obj)
        {
            v.vt = ENUMVAR.VTYPE_ERROR;

            if (obj == null)
            {
                v.vt = ENUMVAR.VTYPE_NULL;
                return;
            }

            Type t = obj.GetType();
            if (IsObject(t))
            {
                int index = global.FindIndex(new Predicate<object>(o => o.Equals(obj)));

                if (index >= 0)
                {
                    v.vt = ENUMVAR.VTYPE_I4;
                    v.lVal = index + 1;
                }
            }
            else if (t == typeof(short) || t == typeof(int) || t == typeof(long))
            {
                v.vt = ENUMVAR.VTYPE_I4;
                v.lVal = (int)obj;
            }
            else if (t == typeof(double) || t == typeof(float))
            {
                v.vt = ENUMVAR.VTYPE_R8;
                v.dblVal = (double)obj;
            }
            else if (t == typeof(bool))
            {
                v.vt = ENUMVAR.VTYPE_BOOL;
                v.bVal = (bool)obj;
            }
            else if (t == typeof(DateTime))
            {
                v.vt = ENUMVAR.VTYPE_TM;

                DateTime dt = (DateTime)obj;
                tm tm = new tm();
                tm.tm_year = dt.Year - 1900;
                tm.tm_mon = dt.Month - 1;
                tm.tm_mday = dt.Day;
                tm.tm_yday = dt.DayOfYear;
                tm.tm_wday = (int)dt.DayOfWeek;
                tm.tm_hour = dt.Hour;
                tm.tm_min = dt.Minute;
                tm.tm_sec = dt.Second;

                IntPtr ptr = Marshal.AllocCoTaskMem(Marshal.SizeOf(tm));
                Marshal.StructureToPtr(tm, ptr, true);

                v.pdate = (long)ptr;
            }
            else if (t == typeof(string))
            {
                v.vt = ENUMVAR.VTYPE_PWSTR;
                v.pwstrVal = (long)Marshal.StringToCoTaskMemAuto((string)obj);
            }
        }

        //Получение строкового представления типа объекта (fully qualified, как в Type.FullName).
        //objectNum — номер объекта в глобальном кэше.
        //Возвращается указатель LPWSTR
        //Вызывается из внешней компоненты.
        [DllExport]
        public static long GetNETObjectType(int objectNum, ref long perror)
        {
            try
            {
                return (long)Marshal.StringToCoTaskMemAuto(global[objectNum - 1].GetType().FullName);

            }
            catch (Exception ex)
            {
                perror = (long)Marshal.StringToCoTaskMemAuto("GetNETObjectType: exception occured!\r\n\r\n" + ex);
                return 0;
            }
        }

        //Является ли тип объектным (не null, не число, не строка, не дата, не булево).
        static bool IsObject(Type type)
        {
            return type != null && type != typeof(string) && type != typeof(DateTime) && type.GetConstructors().Length > 0;  
        }

        //Получение управляемых данных из указателя на укороченный VARIANT, который передан из внешней компоненты. 
        static object FromSVariant(long pVariant = 0, Type t = null)
        {
            if (pVariant == 0)
                return null;
            
            sVariant v = (sVariant)Marshal.PtrToStructure((IntPtr)pVariant, typeof(sVariant));
            switch (v.vt)
            {
                case ENUMVAR.VTYPE_NULL:
                    return null;
                case ENUMVAR.VTYPE_I4:
                    {
                        if (t == null || !IsObject(t))
                            return v.lVal;
                        else
                            return global[v.lVal - 1]; //в этом случае lVal — номер объекта в кєше, возвращается хранимый объект, а не число
                    }
                case ENUMVAR.VTYPE_R8:
                    return v.dblVal;
                case ENUMVAR.VTYPE_BOOL:
                    return v.bVal;
                case ENUMVAR.VTYPE_PWSTR:
                    return Marshal.PtrToStringAuto((IntPtr)v.pwstrVal);
                case ENUMVAR.VTYPE_DATE:
                    {
                        tm tm = (tm)Marshal.PtrToStructure((IntPtr)v.pdate, typeof(tm));
                        return new DateTime(tm.tm_year + 1900, tm.tm_mon + 1, tm.tm_mday, tm.tm_hour, tm.tm_min, tm.tm_sec);
                    }
                default:
                    return null; 
            }
        }
     
        //Получения списка типов по строке с разделителем ";".
        //Следует указывать fully qualified имена типов (как в Type.FullName).
        static Type[] ParseTypes(string types)
        {
            string[] typesArray = types.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
            Type[] retTypes = new Type[typesArray.Length];

            for (int i = 0; i < typesArray.Length; i++)
            {
                Type t = Type.GetType(typesArray[i].Trim());
                if (t == null)
                    return null;
                retTypes[i] = t;
            }

            return retTypes;
        }
    }
}