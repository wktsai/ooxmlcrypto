using System;
using System.Collections.Generic;
using System.Text;
using NUnit.Framework;
using OfficeOpenXml;

namespace OfficeOpenXmlCrypto.Test
{
    [TestFixture]
    public class ExcellCellTest
    {
        [Test]
        public void GetColumnLetter()
        {
            AssertColName("A", 1);
            AssertColName("M", 13);
            AssertColName("Z", 26);
            AssertColName("AA", 27);
            AssertColName("MA", 339);
            AssertColName("MA", 339);
            AssertColName("ZZ", 702);
            AssertColName("AAA", 703);
            AssertColName("ABC", 731);
            AssertColName("CUP", 2590);
            AssertColName("TTT", 14060);
            AssertColName("XFD", 16384);
        }

        void AssertColName(string colName, int colNum) 
        {
            Assert.AreEqual(colName, ExcelCell.GetColumnLetter(colNum), "Wrong at col #" + colNum); 
        }

    }
}
