using System;
using System.Collections.Generic;
using System.Text;
using NUnit.Framework;
using System.IO;

namespace OfficeOpenXmlCrypto.Test
{
    [TestFixture]
    public class OfficeCryptoStreamTest
    {
        [Test]
        public void SaveNewPlaintextFile()
        {
            String plainFile = "plain.xlsx";
            if (File.Exists(plainFile)) { File.Delete(plainFile); }

            OfficeCryptoStream s = new OfficeCryptoStream(plainFile, FileMode.CreateNew, null);

            
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
