using System;
using System.Reflection;
using System.Runtime.InteropServices;

namespace SoftBalance.check_reg
{

    class Program
    {
        private const string GUID = "CD968F5F-1776-4942-8884-573CFAE224E3";
        private const string ProgID = "Softbalance.OneC2RMQcom";

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
            Console.WriteLine($" * GUID: {GUID}");
            Console.WriteLine($" * ProgID: {ProgID}");
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

                Console.Write($"Получаем тип на основании ProgID '{ProgID}'...");
                _type = Type.GetTypeFromProgID(ProgID);
                var _typeGuid = _type.GUID;
                var guidEqualityString = _typeGuid == Guid.Parse(GUID) ? "совпадает" : "НЕ СОВПАДАЕТ";
                Console.WriteLine($" Готово. Полученный GUID: {_typeGuid} ({guidEqualityString})");

                Console.WriteLine("Создаем экземпляр объекта...");
                obj = Activator.CreateInstance(_type);

                string version = (string)_type.InvokeMember("Version", BindingFlags.GetProperty, null, obj, prms);
                string buildInfo = (string)_type.InvokeMember("BuildInfo", BindingFlags.GetProperty, null, obj, null);

                successful = true;
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("Попытка успешно завершена.");
                Console.ResetColor();
                Console.WriteLine($"От компоненты получена информация: version: '{version}', buildinfo: '{buildInfo}'.");
            }
            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(e.ToString());
                Console.ResetColor();
                Console.WriteLine("Попытка провалена.");
            }

            return successful;
        }
    }

}
