using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusSamples
{
    public static class SampleSettings
    {
        public static string ConnectionString { get; } = "Data Source=SampleDb\\EPPlusSampleDb.db;Version=3;";
        //Set the output directory to the SampleApp folder where the app is running from. 
        public static DirectoryInfo OutputDir { get; } = new DirectoryInfo($"{AppDomain.CurrentDomain.BaseDirectory}SampleApp");

    }
}
