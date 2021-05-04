using System;
using System.Reflection;
using System.Runtime.InteropServices;

namespace SoftBalance.check_reg
{

    class Program
    {
        private const string componentGuid = "CD968F5F-1776-4942-8884-573CFAE224E3";
        private const string progID = "Softbalance.OneC2RMQcom";

        static void Main(string[] args)
        {
            Console.WriteLine("Тест работы COM-объекта внешней компоненты OneC2RMQcom.");
            Console.WriteLine("\nИнформация о системе:");
            Console.WriteLine($"* Операционная система: {Environment.OSVersion.VersionString} ({RuntimeInformation.ProcessArchitecture})");
            Console.WriteLine($"* Пользователь ОС: {Environment.UserName}");
            Console.WriteLine($"* .Net реализация: {RuntimeInformation.FrameworkDescription}");
            Console.WriteLine($"* Версия CLR: {Environment.Version}");
            Console.WriteLine($"* Путь CLR: {RuntimeEnvironment.GetRuntimeDirectory()}");
            Console.WriteLine("\nТестируемый COM Объект (предопределенные значения):");
            Console.WriteLine($" * GUID: {componentGuid}");
            Console.WriteLine($" * ProgID: {progID}");
            Console.WriteLine("");

            bool successful = TestCOMCreation();

            if (!successful)
            {
                //Console.WriteLine("");
                //Console.WriteLine("Проверяем регистрацию COM в реестре.");
            }

            if (System.Diagnostics.Debugger.IsAttached)
            {
                Console.WriteLine("\nPress enter...");
                Console.ReadLine();
            }
        }

        private static bool TestCOMCreation()
        {
            Type _type;
            Object obj;
            Object[] prms = new Object[] { }; // Массив пустой, т.к. мы запросим метод Get, принимающий 0 параметров
            bool successful = false;

            try
            {
                Console.WriteLine("Попытка создания COM объекта.");

                Console.Write($"Получаем тип на основании ProgID '{progID}'... ");
                _type = Type.GetTypeFromProgID(progID);
                Guid _typeGuid = _type.GUID;
                string guidEqualityString = _typeGuid == Guid.Parse(componentGuid) ? "совпадает" : "НЕ СОВПАДАЕТ";
                Console.WriteLine($"Готово. Полученный GUID: {_typeGuid} ({guidEqualityString})");
                

                Console.Write("Создаем экземпляр объекта... ");
                obj = Activator.CreateInstance(_type);
                Console.WriteLine("Готово.");

                Console.Write("Обращаемся к свойствам компоненты... ");
                string version = (string)_type.InvokeMember("Version", BindingFlags.GetProperty, null, obj, prms);
                string buildInfo = (string)_type.InvokeMember("BuildInfo", BindingFlags.GetProperty, null, obj, null);
                Console.WriteLine("Готово.");

                successful = true;
                Console.WriteLine($"От компоненты получена информация: Version: '{version}', BuildInfo: '{buildInfo}'.");
            }
            catch (Exception e)
            {
                Console.WriteLine("Возникла ошибка:");
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(e.ToString());
                Console.ResetColor();
            }

            return successful;
        }
    }

}
