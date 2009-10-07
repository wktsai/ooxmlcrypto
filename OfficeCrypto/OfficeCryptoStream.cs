using System;
using System.Collections.Generic;
using System.Text;
using System.IO;

/*
 * Wrapper coded by Danilo Mirkovic, Oct 2009
 * License: Open source, GPL 
 *  
 * Note: 
 * - OfficeCrypto class is LGPL2/Apache license.
 *   http://www.lyquidity.com/devblog/?p=35
 * - NPOI is Apache 2.0 license
 *   http://npoi.codeplex.com/
 */
namespace OfficeOpenXmlCrypto
{
    /// <summary>
    /// TODO: document
    /// provide examples using Package
    /// </summary>
    public class OfficeCryptoStream : MemoryStream
    {
        String _password = null;

        // Encrypted or plaintext stream (of the underlying storage file)
        readonly Stream Storage;

        /// <summary>
        /// Create based on a file.
        /// </summary>
        /// <param name="file"></param>
        /// <param name="mode"></param>
        /// <param name="password">Password. Pass null for plaintext.</param>
        public OfficeCryptoStream(String file, FileMode mode, String password)
            : this (new FileStream(file, mode), password) { }

        /// <summary>
        /// Create based on a stream.
        /// </summary>
        /// <param name="stream">Storage stream, usually FileStream</param>
        /// <param name="password">Password. Pass null for plaintext.</param>
        public OfficeCryptoStream(Stream stream, String password)
        {
            Storage = stream;
            Password = password;

            if (stream.Length == 0) 
            {
                // No need to decrypt, stream is already 0-length 
                return;
            }

            byte[] contents;
            if (Encrypted)
            {
                // Decrypt
                OfficeCrypto oc = new OfficeCrypto();
                contents = oc.DecryptToArray(stream, password);
            }
            else
            {
                // Read plaintext
                contents = new byte[stream.Length];
                stream.Read(contents, 0, contents.Length);
            }

            base.Write(contents, 0, contents.Length);
            base.Flush();
            base.Position = 0;

            // TODO: detect wrong password, throw a well-defined exception
            // TODO: detect wrong file format
        }

        /// <summary>
        /// True if stream is encrypted (has a password), false otherwise.
        /// </summary>
        public bool Encrypted
        {
            get { return !String.IsNullOrEmpty(_password); } 
        }

        /// <summary>
        /// Gets or sets the password. Set to null for plaintext (no encryption).
        /// Throws InvalidOperationException if stream is read-only or does 
        /// not support seeking.
        /// </summary>
        public String Password
        {
            get { return _password; }
            set 
            {
                // Throw exception if closed
                if (!base.CanWrite || !base.CanSeek)
                {
                    throw new InvalidOperationException("Cannot set password. Underlying stream does not support seek or write. Make sure it was not closed.");
                }
                _password = value; 
            }
        }

        /// <summary>
        /// Close the stream and perform encryption if needed.
        /// </summary>
        public override void Close()
        {
            base.Close();

            Storage.Seek(0, SeekOrigin.Begin);
            Storage.SetLength(0);
            Storage.Position = 0;

            if (Encrypted)
            {
                // Encrypt this to the storage stream
                OfficeCrypto oc = new OfficeCrypto();
                oc.EncryptToStream(base.ToArray(), Password, Storage); 
            }
            else
            {
                // Just write the contents to storage stream
                base.WriteTo(Storage);
            }

            Storage.Close();
        }
    }
}
