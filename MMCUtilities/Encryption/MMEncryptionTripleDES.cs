using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace MMUtilities {
    public class MMEncryptionTripleDES : IDisposable {


        #region Implementing Dispose
        private IntPtr handle;
        private Component component = new Component();
        private bool disposed = false;
        public MMEncryptionTripleDES( IntPtr handle) {
            this.handle = handle;
        }
        public void Dispose() {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
        private void Dispose(bool disposing) {
            if (!this.disposed) {
                if (disposing) {
                    component.Dispose();
                }
                CloseHandle(handle);
                handle = IntPtr.Zero;
            }
            disposed = true;
        }
        [System.Runtime.InteropServices.DllImport("Kernel32")]
        private extern static Boolean CloseHandle(IntPtr handle);
        ~MMEncryptionTripleDES() {
            Dispose(false);
        }
        #endregion


        private string encryDecryptionKey = string.Empty;
        public string EncriptionKey {
            get { return encryDecryptionKey; }
        }

        public MMEncryptionTripleDES() {
            this.encryDecryptionKey = string.Empty;
        }

        /// <summary>
        /// Generate an encryption key of 128 Bits.
        /// </summary>
        /// <param name="keySize"></param>
        /// <returns></returns>
        public string GenerateKeysOf128Bits() {
            string key = string.Empty;
            RNGCryptoServiceProvider rng = new RNGCryptoServiceProvider();
            byte[] buff = new byte[16];
            rng.GetBytes(buff);
            key = Convert.ToBase64String(buff);
            return key;
        }

        /// <summary>
        /// This method allow the user to encrypt information based on the key provided
        /// </summary>
        /// <param name="dataToEncrypt"></param>
        /// <returns></returns>
        public string EncryptData(string dataToEncrypt, string encryDecryptionKey )
        {
            this.encryDecryptionKey = encryDecryptionKey;
            if (this.encryDecryptionKey == string.Empty)
                throw new Exception("The encription key can not be empty");

            string encryptResult = string.Empty;
            byte[] inputArray = UTF8Encoding.UTF8.GetBytes(dataToEncrypt);
            TripleDESCryptoServiceProvider tripleDES = new TripleDESCryptoServiceProvider();
            tripleDES.Key     = UTF8Encoding.UTF8.GetBytes(this.encryDecryptionKey);
            tripleDES.Mode    = CipherMode.ECB;
            tripleDES.Padding = PaddingMode.PKCS7;
            ICryptoTransform cTransform = tripleDES.CreateEncryptor();
            byte[] resultArray = cTransform.TransformFinalBlock(inputArray, 0, inputArray.Length);
            tripleDES.Clear();
            encryptResult = Convert.ToBase64String(resultArray, 0, resultArray.Length);
            return encryptResult;
        }

        /// <summary>
        /// This method allow tho the user to decrypt information based on the key provided
        /// Key should be either of 128 bit or of 192 bit  
        /// </summary>
        /// <param name="dataToEncrypt"></param>
        /// <returns></returns>
        public string DecryptData( string dataToDecrypt, string encryDecryptionKey ) {
            this.encryDecryptionKey = encryDecryptionKey;
            if (this.encryDecryptionKey == string.Empty)
                throw new Exception("The encription key can not be empty");

            string decryptResult = string.Empty;
            byte[] inputArray = Convert.FromBase64String(dataToDecrypt);
            TripleDESCryptoServiceProvider tripleDES = new TripleDESCryptoServiceProvider();
            tripleDES.Key = UTF8Encoding.UTF8.GetBytes(this.encryDecryptionKey);
            tripleDES.Mode = CipherMode.ECB;
            tripleDES.Padding = PaddingMode.PKCS7;
            ICryptoTransform cTransform = tripleDES.CreateDecryptor();
            byte[] resultArray = cTransform.TransformFinalBlock(inputArray, 0, inputArray.Length);
            tripleDES.Clear();
            decryptResult = UTF8Encoding.UTF8.GetString(resultArray);
            return decryptResult;
        }



    }
}
