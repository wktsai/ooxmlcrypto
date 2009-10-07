using System;
using System.Collections.Generic;
using System.Text;
using NUnit.Framework;
using System.IO;
using OfficeOpenXml;

namespace OfficeOpenXmlCrypto.Test
{
    [TestFixture]
    public class OfficeCryptoStreamTest
    {
        String TestFile = "test.xlsx";

        [SetUp]
        public void TestSetUp()
        {
            if (File.Exists(TestFile)) { File.Delete(TestFile); }
        }

        [Test]
        public void PlaintextFile()
        {
            CreateTestWorkbook(null);
            AssertFileCorrect(null);
        }

        [Test]
        public void EncryptedFile()
        {
            CreateTestWorkbook("foo");
            AssertFileCorrect("foo");
        }

        [Test]
        // This works -- if file is saved later, 
        // we encrypt it and apply the password.
        public void PlaintextFileUsingPassword()
        {
            CreateTestWorkbook(null);
            AssertFileCorrect("bar");
        }

        [Test]
        [ExpectedException(typeof(InvalidPasswordException))]
        public void EncryptedFileEmptyPassword()
        {
            CreateTestWorkbook("foo");
            AssertFileCorrect(null);
        }

        [Test]
        public void TryOpen()
        {
            CreateTestWorkbook("foo");
            OfficeCryptoStream ocs;
            Assert.IsFalse(OfficeCryptoStream.TryOpen(TestFile, null, out ocs));
            Assert.IsFalse(OfficeCryptoStream.TryOpen(TestFile, "bar", out ocs));
            Assert.IsTrue(OfficeCryptoStream.TryOpen(TestFile, "foo", out ocs));
            ocs.Close();
        }

        [Test]
        [ExpectedException(typeof(InvalidPasswordException))]
        public void EncryptedFileInvalidPassword()
        {
            CreateTestWorkbook("foo");
            AssertFileCorrect("bar");
        }

        [Test]
        public void PasswordChange()
        {
            CreateTestWorkbook("foo");
            ChangePassword("foo", "bar");
            AssertFileCorrect("bar");
        }

        [Test]
        public void PasswordAddToPlaintext()
        {
            CreateTestWorkbook(null);
            ChangePassword(null, "bar");
            AssertFileCorrect("bar");
        }

        [Test]
        public void PasswordRemoveFromEncrypted()
        {
            CreateTestWorkbook("foo");
            ChangePassword("foo", null);
            AssertFileCorrect(null);
        }

        void AssertFileCorrect(String password)
        {
            using (OfficeCryptoStream s = OfficeCryptoStream.Open(TestFile, password))
            {
                using (ExcelPackage p = new ExcelPackage(s))
                {
                    Assert.IsNotNull(p, "Cannot create package.");
                    ExcelWorksheet ws = p.Workbook.Worksheets["Test"];
                    Assert.IsNotNull(ws, "No Test worksheet.");
                    String cval = ws.Cell(1, 1).Value;
                    Assert.AreEqual("Test Cell", cval, "First cell value incorrect.");
                }
            }
        }

        void ChangePassword(String oldPassword, String newPassword)
        {
            using (OfficeCryptoStream s = OfficeCryptoStream.Open(TestFile, oldPassword))
            {
                s.Password = newPassword;
                s.Save();
            }
        }

        void CreateTestWorkbook(String password)
        {
            using (OfficeCryptoStream s = OfficeCryptoStream.Create(TestFile))
            {
                s.Password = password;
                using (ExcelPackage p = new ExcelPackage(s))
                {
                    ExcelWorksheet ws = p.Workbook.Worksheets["Test"];
                    if (ws == null) { ws = p.Workbook.Worksheets.Add("Test"); }
                    ws.Cell(1, 1).Value = "Test Cell";
                    p.Save();
                }
                s.Save();
            }
        }

        /*
         *         private void button1_Click(object sender, EventArgs e)
        {
            OfficeCrypto oc = new OfficeCrypto();
            byte[] contents = oc.DecryptToArray(@"C:\temp\mww_crypto_foo.xlsx", "foo");

            File.WriteAllBytes(@"C:\temp\z_decrypted.xlsx", contents);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            byte[] contents = File.ReadAllBytes(@"C:\temp\mww_plain.xlsx");
            
            OfficeCrypto oc = new OfficeCrypto();
            oc.EncryptToFile(contents, "bar", @"C:\temp\z_encrypted_bar.xlsx");
        }
*/

    }
}
