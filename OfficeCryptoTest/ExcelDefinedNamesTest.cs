using System;
using System.Collections.Generic;
using System.Text;
using NUnit.Framework;
using System.Diagnostics;

namespace OfficeOpenXmlCrypto.Test
{
    [TestFixture]
    public class ExcelDefinedNamesTest : ExcelWorksheetsTest
    {
        const int Count = 5;

        [Test]
        public void aCreateDefinedNames()
        {
            for(int i=0; i<Count; i++)
            {
                Assert.IsFalse(doc.DefinedNames.Contains(Name(i)));
            }
            CreateElements();
            for(int i=0; i<Count; i++)
            {
                Assert.IsTrue(doc.DefinedNames.Contains(Name(i)));      
            }
        }

        [Test]
        [Explicit]
        public void zCreatePermanentDoc()
        {
            CreateElements();
            doc.Add("sheet1");
            Process.Start("."); // open current dir
        }

        void CreateElements()
        {
            for (int i = 0; i < Count; i++)
            {
                doc.DefinedNames.Add(Name(i), RangeRef(i));
            }
        }

        static string RangeRef(int i)
        {
            return "sheet1!$A$1:$D$" + (i + 10);
        }

        static string Name(int i)
        {
            return "TestDefinedName" + i;
        }

        [Test]
        public void RangeRefsReturned()
        {
            CreateElements();
            for (int i = 0; i < Count; i++)
            {
                Assert.AreEqual(RangeRef(i), doc.DefinedNames[Name(i)] = RangeRef(i));
            }
        }

        [Test]
        public void Remove()
        {
            CreateElements();
            for (int i = 0; i < Count; i++)
            {
                doc.DefinedNames.Remove(Name(i));
            }
            for (int i = 0; i < Count; i++)
            {
                Assert.IsFalse(doc.DefinedNames.Contains(Name(i)));
            }
        }

    }
}
