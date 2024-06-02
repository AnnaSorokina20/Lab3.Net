using System;
using System.Reflection;

namespace OfficeApp
{
    class Program
    {
        static void Main(string[] args)
        {
            string assemblyPath = "OfficeLib.dll"; // Укажите правильный путь к сборке OfficeLib.dll
            Assembly assembly = Assembly.LoadFrom(assemblyPath);

            // Load ExcelDocument dynamically
            Type excelType = assembly.GetType("OfficeLib.ExcelDocument");
            dynamic excelInstance = Activator.CreateInstance(excelType);

            // Use ExcelDocument
            excelInstance["A1"] = "Dynamic Cell A1";
            excelInstance.SaveAs("DynamicTestDocument.xlsx");

            // Load WordDocument dynamically
            Type wordType = assembly.GetType("OfficeLib.WordDocument");
            dynamic wordInstance = Activator.CreateInstance(wordType);

            // Use WordDocument
            wordInstance[1] = "Dynamic Paragraph 1";
            wordInstance.SaveAs("DynamicTestWordDocument.docx");

            Console.WriteLine("Documents created successfully.");
        }
    }
}
