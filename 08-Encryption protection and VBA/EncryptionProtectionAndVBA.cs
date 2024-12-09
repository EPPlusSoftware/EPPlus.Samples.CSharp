using EPPlusSamples.EncryptionProtectionAndVba;
using System;
namespace EPPlusSamples
{
    public static class EncryptionProtectionAndVBASample
    {
        public static void Run()
        {
            //Sample 8.1 Swedish Quiz : Shows Encryption, workbook- and worksheet protection.
            EncryptionAndProtection.Run();
            
            //Sample 8.2 - Shows how to work with macro-enabled workbooks(VBA) and how to sign the code with a certificate.
            Console.WriteLine("Running sample 8.2-VBA");
            WorkingWithVbaSample.Run();
            SigningYourVBAProject.Run();

            //Sample 8.3 shows how to sign workbooks
            DigitalSignatureSample.Run();

            Console.WriteLine("Sample 8.2 created {0}", FileUtil.OutputDir.Name);
            Console.WriteLine();
        }
    }
}
