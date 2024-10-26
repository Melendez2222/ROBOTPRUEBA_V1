using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ROBOTPRUEBA_V1.CONFIG
{
    internal class GlobalSettings
    {
        
        public static string TxtFileName => "extracted_digits.txt";
        public static List<string> ExtractedDigitsList { get; set; } = new List<string>();
        public static string logFilePath => @"C:\Users\cristhian.melendez\Desktop\PRUEBAROBOT\logs\log.txt";
        public static string ExcelFileManifestSunat { get; set; }
        public static string DownloadDirectory { get; private set; }
        public static string ConsultaURL { get; private set; }
        public static List<string> Codes { get; private set; }
        public static void Initialize(IConfiguration configuration)
        {
            ExcelFileManifestSunat = configuration["FilePaths:ExcelFileManifestSunat"];
            DownloadDirectory = configuration["FilePaths:DownloadDirectory"];
            ConsultaURL = configuration["ConsultaURL"];
            Codes = configuration.GetSection("Codes").Get<List<string>>();
        }
    }
}
