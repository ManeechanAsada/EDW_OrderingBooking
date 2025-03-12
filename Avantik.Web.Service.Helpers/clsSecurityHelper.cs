using System;
using System.Collections.Generic;
using System.Text;
using System.Configuration;
using System.IO;
using System.IO.Compression;
using System.Security.Cryptography;

namespace Avantik.Web.Service.Helpers
{
    public static class SecurityHelper
    {
        public static string EncryptString(string Message, string Passphrase)
        {
            string strResult = string.Empty;
            byte[] Results = null;
            System.Text.UTF8Encoding UTF8 = new System.Text.UTF8Encoding();

            // Step 1. We hash the passphrase using MD5
            // We use the MD5 hash generator as the result is a 128 bit byte array
            // which is a valid length for the TripleDES encoder we use below

            using (System.Security.Cryptography.MD5CryptoServiceProvider HashProvider = new System.Security.Cryptography.MD5CryptoServiceProvider())
            {
                byte[] TDESKey = HashProvider.ComputeHash(UTF8.GetBytes(Passphrase));

                // Step 2. Create a new TripleDESCryptoServiceProvider object
                using (System.Security.Cryptography.TripleDESCryptoServiceProvider TDESAlgorithm = new System.Security.Cryptography.TripleDESCryptoServiceProvider())
                {
                    // Step 3. Setup the encoder
                    TDESAlgorithm.Key = TDESKey;
                    TDESAlgorithm.Mode = System.Security.Cryptography.CipherMode.ECB;
                    TDESAlgorithm.Padding = System.Security.Cryptography.PaddingMode.PKCS7;

                    // Step 4. Convert the input string to a byte[]
                    byte[] DataToEncrypt = UTF8.GetBytes(Message);

                    // Step 5. Attempt to encrypt the string
                    try
                    {
                        using (System.Security.Cryptography.ICryptoTransform Encryptor = TDESAlgorithm.CreateEncryptor())
                        {
                            Results = Encryptor.TransformFinalBlock(DataToEncrypt, 0, DataToEncrypt.Length);
                        }
                    }
                    catch
                    { strResult = string.Empty; }
                }

            }

            // Step 6. Return the encrypted string as a base64 encoded string
            return Convert.ToBase64String(Results);
        }

        public static string DecryptString(string Message, string Passphrase)
        {
            string strResult = string.Empty;
            byte[] Results;
            System.Text.UTF8Encoding UTF8 = new System.Text.UTF8Encoding();

            // Step 1. We hash the passphrase using MD5
            // We use the MD5 hash generator as the result is a 128 bit byte array
            // which is a valid length for the TripleDES encoder we use below

            using (System.Security.Cryptography.MD5CryptoServiceProvider HashProvider = new System.Security.Cryptography.MD5CryptoServiceProvider())
            {
                byte[] TDESKey = HashProvider.ComputeHash(UTF8.GetBytes(Passphrase));

                // Step 2. Create a new TripleDESCryptoServiceProvider object
                using (System.Security.Cryptography.TripleDESCryptoServiceProvider TDESAlgorithm = new System.Security.Cryptography.TripleDESCryptoServiceProvider())
                {
                    // Step 3. Setup the decoder
                    TDESAlgorithm.Key = TDESKey;
                    TDESAlgorithm.Mode = System.Security.Cryptography.CipherMode.ECB;
                    TDESAlgorithm.Padding = System.Security.Cryptography.PaddingMode.PKCS7;

                    try
                    {
                        // Step 4. Convert the input string to a byte[]
                        byte[] DataToDecrypt = Convert.FromBase64String(Message);

                        // Step 5. Attempt to decrypt the string
                        using (System.Security.Cryptography.ICryptoTransform Decryptor = TDESAlgorithm.CreateDecryptor())
                        {
                            Results = Decryptor.TransformFinalBlock(DataToDecrypt, 0, DataToDecrypt.Length);
                        }
                        strResult = UTF8.GetString(Results);
                    }
                    catch
                    { strResult = string.Empty; }
                }
            }

            // Step 6. Return the decrypted string in UTF8 format
            return strResult;
        }
        public static string EncryptStringSHA1(string strToEncryp)
        {
            string hashString = string.Empty;
            using (System.Security.Cryptography.SHA1CryptoServiceProvider sha1 = new System.Security.Cryptography.SHA1CryptoServiceProvider())
            {

                try
                {
                    System.Text.UTF8Encoding ue = new System.Text.UTF8Encoding();

                    byte[] bytes = ue.GetBytes(strToEncryp);
                    byte[] hashBytes = sha1.ComputeHash(bytes);
                    // Convert the encrypted bytes back to a string (base 16)             

                    for (int i = 0; i < hashBytes.Length; i++)
                    {
                        hashString += Convert.ToString(hashBytes[i], 16).PadLeft(2, '0');
                    }
                }
                catch
                { }
            }

            //Return encrypted value.
            if (string.IsNullOrEmpty(hashString) == false)
            {
                return hashString.PadLeft(32, '0');
            }
            else
            {
                return string.Empty;
            }
        }
        public static string MD5Encrypt(string parameters)
        {
            //MD5 Encrypt
            MD5CryptoServiceProvider md5 = new MD5CryptoServiceProvider();
            UTF8Encoding enc = new UTF8Encoding();
            byte[] input = null;
            byte[] output = null;
            System.Text.StringBuilder hash = new StringBuilder();

            input = enc.GetBytes(parameters);
            output = md5.ComputeHash(input);

            foreach (byte byt in output)
            {
                hash.Append(byt.ToString("x2"));
            }

            return hash.ToString();
        }
        public static string CompressString(string text)
        {
            byte[] buffer = Encoding.UTF8.GetBytes(text);
            using (MemoryStream ms = new MemoryStream())
            {
                using (GZipStream zip = new GZipStream(ms, CompressionMode.Compress, true))
                {
                    zip.Write(buffer, 0, buffer.Length);
                }

                ms.Position = 0;

                byte[] compressed = new byte[ms.Length];
                ms.Read(compressed, 0, compressed.Length);

                byte[] gzBuffer = new byte[compressed.Length + 4];
                System.Buffer.BlockCopy(compressed, 0, gzBuffer, 4, compressed.Length);
                System.Buffer.BlockCopy(BitConverter.GetBytes(buffer.Length), 0, gzBuffer, 0, 4);

                return Convert.ToBase64String(gzBuffer);
            }
        }
        public static string DecompressString(string compressedText)
        {
            byte[] gzBuffer = Convert.FromBase64String(compressedText);
            using (MemoryStream ms = new MemoryStream())
            {
                int msgLength = BitConverter.ToInt32(gzBuffer, 0);
                ms.Write(gzBuffer, 4, gzBuffer.Length - 4);

                byte[] buffer = new byte[msgLength];

                ms.Position = 0;
                using (GZipStream zip = new GZipStream(ms, CompressionMode.Decompress))
                {
                    zip.Read(buffer, 0, buffer.Length);
                }

                return Encoding.UTF8.GetString(buffer);
            }
        }
    }
}
