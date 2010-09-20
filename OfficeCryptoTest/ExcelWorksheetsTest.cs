using System;
using System.Collections.Generic;
using System.Text;
using NUnit.Framework;
using OfficeOpenXml;
using System.IO;

namespace OfficeOpenXmlCrypto.Test
{
    [TestFixture]
    public class ExcelWorksheetsTest
    {
        protected ExcelPackage package;
        protected ExcelWorksheets doc;

        protected const String Filename = "testWs.xlsx";

        [SetUp]
        public void SetUp()
        {
            if (File.Exists(Filename)) { File.Delete(Filename); }
            package = new ExcelPackage(new FileInfo(Filename));
            doc = package.Workbook.Worksheets;
        }

        [TearDown]
        public void TearDown()
        {
            if (package != null)
            {
                if (package.Workbook.Worksheets.Count > 0)
                {
                    package.Save();
                }
                package.Dispose();
                package = null;
            }
        }

        [Test]
        public void AddSheets()
        {
            ExcelWorksheet one = doc.Add("one");
            ExcelWorksheet two = doc.Add("two");

            Assert.AreEqual(one, doc[1]);
            Assert.AreEqual(two, doc[2]);
            Assert.AreEqual(one, doc["one"]);
            Assert.AreEqual(two, doc["two"]);
        }

        [Test]
        public void RenameSheet()
        {
            ExcelWorksheet one = doc.Add("one");
            Assert.AreEqual(one, doc[1]);
            one.Name = "jedan";
            Assert.AreEqual(one, doc["jedan"]);
        }

        [Test]
        public void DeleteSheets()
        {
            ExcelWorksheet one = doc.Add("one");
            ExcelWorksheet two = doc.Add("two");
            ExcelWorksheet three = doc.Add("three");
            Assert.AreEqual(3, doc.Count);

            doc.Delete(2);
            Assert.AreEqual(2, doc.Count);
            Assert.AreEqual(one, doc[1]);
            Assert.AreEqual(three, doc[2]);
            Assert.AreEqual(null, doc["two"]);

            doc.Delete(1);
            Assert.AreEqual(1, doc.Count);
            Assert.AreEqual(three, doc[1]);
            Assert.AreEqual(null, doc["one"]);
        }

        [Test]
        public void DeleteThenAddSheets()
        {
            ExcelWorksheet one = doc.Add("one");
            ExcelWorksheet two = doc.Add("two");
            Assert.AreEqual(2, doc.Count);

            doc.Delete(1);
            Assert.AreEqual(1, doc.Count);
            Assert.AreEqual(two, doc[1]);

            ExcelWorksheet anotherOne = doc.Add("one");
            Assert.AreEqual(2, doc.Count);
            Assert.AreEqual(anotherOne, doc["one"]);
            Assert.AreNotEqual(anotherOne, one);            
        }

    }
}
