using System;
using System.IO;
using System.Security.Cryptography;
using Aspose.Pdf;

namespace ExFormOfficeAddInBAL
{
    public class ExFormsAsposePdf
    {
        public static Document GetDocument(string pdfPath)
        {
            return new Document(pdfPath);
        }
        public static void LoadLicenseFromFile(string licenseFilePath)
        {
            License license = new License();
            license.SetLicense(licenseFilePath);
        }
        public static void LoadLicenseFromStream(byte[] decryptedLicense)
        {
            MemoryStream licenseStream = new MemoryStream(decryptedLicense);

            License license = new License();
            license.SetLicense(licenseStream);
        }
        public static byte[] EncryptDecryptLicense(string licenseFilePath, string myDir)
        {
            string encryptedFilePath = myDir + @"\EncryptedLicense.txt";
            byte[] licBytes = File.ReadAllBytes(licenseFilePath);
            byte[] key = GenerateKey(licBytes.Length);

            File.WriteAllBytes(encryptedFilePath, EncryptDecryptLicense(licBytes, key));

            byte[] decryptedLicense = EncryptDecryptLicense(File.ReadAllBytes(encryptedFilePath), key);
            return decryptedLicense;
        }

        private static byte[] EncryptDecryptLicense(byte[] licBytes, byte[] key)
        {
            byte[] output = new byte[licBytes.Length];

            for (int i = 0; i < licBytes.Length; i++)
                output[i] = Convert.ToByte(licBytes[i] ^ key[i]);

            return output;
        }

        private static byte[] GenerateKey(long size)
        {
            RNGCryptoServiceProvider rng = new RNGCryptoServiceProvider();
            byte[] strongBytes = new Byte[size];
            rng.GetBytes(strongBytes);

            return strongBytes;
        }
    }
}
