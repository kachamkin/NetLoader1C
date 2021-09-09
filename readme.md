Еще один загрузчик .NET, позволяющий создавать объекты .NET и обращаться к их методам и свойствам
непосредственно из "1С".

Внешняя компонета, написанная по технологии Native API, служит просто прокси между "1С" и управляемой сборкой
"11CAddin.dll", которая, собственно, и реализует весь функционал. Работает экономно с точки зрения памяти и производительности,
поскольку не требует запуска отдельного runtime-хоста .NET. Конечно, поскольку речь идет о разработке под Windows,
можно было бы обойтись единственной dll, — внешней COM-компонентой, — но хотелось обойтись без COM и регистрации.

Из "1С" могут быть вызваны три метода: 

CreateObject/СоздатьОбъект;
InvokeMember/ВызватьМетодИлиСвойство;
GetObjectType/ПолучитьТипОбъекта (этот метод сводится к вызовам первых двух, но сохранен для удобства).

Аргументы и возвращаемые значения подробно описаны в комментариях к коду.

Пример вызовов из "1С":

//Подключение<br>
ПодключитьВнешнююКомпоненту("C:\Users\kacha\source\repos\VNCOMP83\template\x64\Debug\AddInNative.dll", "Comp", ТипВнешнейКомпоненты.Native);<br> 
ВК = Новый("AddIn.Comp.NetLoader");<br>

//Создание объекта "List"<br>
СсылкаНаОбъект = ВК.СоздатьОбъект("System.Collections.Generic.List`1[[System.Object, mscorlib, Version = 4.0.0.0, Culture = neutral, PublicKeyToken = b77a5c561934e089]]");<br>

//Добавление в объект "List" элемента типа "Целое число" со значением "1"<br>
ВК.ВызватьМетодИлиСвойство(СсылкаНаОбъект, "Add", "System.Int32",, 1);<br>   

//Получение типа созданного объекта "List"<br>
Сообщить(ВК.ПолучитьТипОбъекта(СсылкаНаОбъект));<br>     

//Получение числа элементов в объекте "List"<br>
Сообщить(ВК.ВызватьМетодИлиСвойство(СсылкаНаОбъект, "Count"));
