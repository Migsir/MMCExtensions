using MMCSirUtilities;
using MMUtilities;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace MMCUtilities.Test {
    class Program {

        public void DecryptData() {
            string encryptionKey = "10ndxY52WiogaRJbEAnv5Q==";
            using ( MMEncryptionTripleDES objEncryption = new MMEncryptionTripleDES() ) {
                string decryptedData = objEncryption.DecryptData("9v8dsY/kqLTkIp6aOS1U/fWV52QRpnDahhmAUNborySSWeUy3V+AH0K", encryptionKey);
            }
        }



        private static void testEncryption() {
            //10ndxY52WiogaRJbEAnv5Q==
            //string key = Convert.ToBase64String(buff);
            //string key = "10ndxY52WiogaRJbEAnv5Q==";
            //string key = "8jdODYrycSzTitEss2papg==";

            MMEncryptionTripleDES objEncription = new MMEncryptionTripleDES();

            string key = objEncription.GenerateKeysOf128Bits();
            //objEncription.EncriptionKey = "D6YluF7S9izIXbEvPp8S3g=="; //key;
            //objEncription.EncriptionKey = key;
            //string decryption = objEncription.DecryptData("9v8dsY/kqLTkIp6aOS1U/fWV52QRpnDahhmAUNborySSWeUy3V+AH0KrYRUIaKyfT0pAPlano6ETWvSHye1rgt1EK0WdCdWdmciHrXWg/9vZZRUAB+zzDa8xmFPbnTtv374fXe7SCzXBjKy+508hPTI7xb6o+q1Bv6Q/QTTIOi95yFWhvZiCP07nCn2VfmBSqwpKAv/nxdJFcMJLXnbgMEv24HNDGV6TlTTlL+Ux3sZA/TEP4eq5iz1wCdh+BdVD76869ZbPnUyfop8dOFyQzLMz1W1JBD91");
            string encrypt = objEncription.EncryptData("Data Source=MXGDLM4SMESQL02; Initial Catalog=MasterData; User ID=sapgerenicuser;Password=3l3m3nt0g3n3r1c;Data Source=MXGDLM4SMESQL02; Initial Catalog=MasterData; User ID=sapgerenicuser;Password=3l3m3nt0g3n3r1c", key);
        }

        static void Main(string[] args) {
            //TestListToDatatable();
            //TestCreateExcelFile();
            //TestCreateCSVFileFromDT();
            //TestListToCSV();
            testEncryption();
        }
    }
}
