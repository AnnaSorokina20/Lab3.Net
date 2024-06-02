using OfficeLib;
using System.Diagnostics;
using System.IO;

namespace TestOffice
{
    [TestClass]
    public class OfficeTest
    {
        static int c1 = -1, c3 = -1;
        string documentsPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

        [TestMethod("Start / stop excel com server")]
        public void Test01()
        {
            bool Result;
            int c2;
            c1 = Process.GetProcessesByName("excel").Length;
            using (var x = new ExcelDocument())
            {
                c2 = Process.GetProcessesByName("excel").Length;
            }
            Thread.Sleep(3000);
            Result = (c2 - c1 == 1);
            Assert.IsTrue(Result);
        }

        [TestMethod("Create new excel file")]
        public void Test02()
        {
            string FileName = "TestDocument.xlsx";
            string FullName = Path.Combine(documentsPath, FileName);
            if (File.Exists(FullName))
                File.Delete(FullName);
            using (var x = new ExcelDocument())
            {
                x["A1"] = "Cell A1";
                x["A2"] = "Cell A2";
                x[2, 2] = "Cell 2,2";
                x[3, 3] = "100.0";
                x[4, 4] = "100.5";
                x[5, 5] = "100,5";

                x.SaveAs(FullName);
                Assert.IsTrue(File.Exists(FullName));
            }
        }

        [TestMethod("Check content")]
        public void Test03()
        {
            string FileName = "TestDocument.xlsx";
            string FullName = Path.Combine(documentsPath, FileName);
            bool Result = File.Exists(FullName);
            using (var x = new ExcelDocument(FullName))
            {
                string valueA1 = x["A1"];
                string valueA2 = x["A2"];
                string value22 = x[2, 2];
                string value33 = x[3, 3];
                string value44 = x[4, 4];
                string value55 = x[5, 5];

                Console.WriteLine($"A1: {valueA1}, A2: {valueA2}, 2,2: {value22}, 3,3: {value33}, 4,4: {value44}, 5,5: {value55}");

                Result &= valueA1 == "Cell A1";
                Result &= valueA2 == "Cell A2";
                Result &= value22 == "Cell 2,2";
                Result &= value33 == "100";
                Result &= value44 == "100.5";
                Result &= value55 == "100,5";
            }
            Assert.IsTrue(Result, $"Expected content does not match in {FullName}");
        }

        [TestMethod("Check Garbage collector")]
        public void Test04()
        {
            Thread.Sleep(3000);
            c3 = Process.GetProcessesByName("excel").Length;
            Assert.IsTrue(c1 == c3);
        }

        [TestMethod("Create new word file")]
        public void Test05()
        {
            string FileName = "TestWordDocument.docx";
            string FullName = Path.Combine(documentsPath, FileName);
            if (File.Exists(FullName))
                File.Delete(FullName);
            using (var x = new WordDocument())
            {
                x[1] = "Paragraph 1";
                x[2] = "Paragraph 2";
                x[3] = "Paragraph 3";

                x.SaveAs(FullName);
                Assert.IsTrue(File.Exists(FullName));
            }
        }

        [TestMethod("Check Word content")]
        public void Test06()
        {
            string FileName = "TestWordDocument.docx";
            string FullName = Path.Combine(documentsPath, FileName);
            bool Result = File.Exists(FullName);
            using (var x = new WordDocument(FullName))
            {
                string value1 = x[1];
                string value2 = x[2];
                string value3 = x[3];

                Console.WriteLine($"1: {value1}, 2: {value2}, 3: {value3}");

                Result &= value1 == "Paragraph 1\r\n";
                Result &= value2 == "Paragraph 2\r\n";
                Result &= value3 == "Paragraph 3\r\n";
            }
            Assert.IsTrue(Result, $"Expected content does not match in {FullName}");
        }

        [TestMethod("Check Garbage collector for Word")]
        public void Test07()
        {
            Thread.Sleep(3000);
            int wordProcesses = Process.GetProcessesByName("WINWORD").Length;
            Assert.IsTrue(wordProcesses == 0);
        }
    }
}
