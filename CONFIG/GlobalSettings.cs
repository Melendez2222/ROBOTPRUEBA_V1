using DocumentFormat.OpenXml.Office2010.ExcelAc;
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
		public static Dictionary<string, string> ExtractedDigitsList { get; set; } = new Dictionary<string, string>();
		public static string logFilePath { get; set; }
        public static string ExcelFileManifestSunat { get; set; }
        public static string DownloadDirectory { get; private set; }
        public static string ConsultaURL { get; private set; }
        public static List<string> Codes { get; set; }
        public static List<string> DigitsListFiltered { get; set; } = new List<string>();

        public static Dictionary<string, List<(string Text, string Href)>> Detalleurls = new Dictionary<string, List<(string Text, string Href)>>();
        public static string FechaInicio { get; set; }
        public static string FechaFin { get; set; }
        public static string NumSemana { get; set; }
		public static List<string> CodeAduanaList { get; set; } = new List<string>();
		public static Dictionary<string, List<string>> GlobalProductDictionary { get; set; } = new Dictionary<string, List<string>>();
		public static Dictionary<string, List<DetalleInfo>> DetalleData = new Dictionary<string, List<DetalleInfo>>();
        public static string CurrentCode { get; set; }

        public static void Initialize(IConfiguration configuration)
        {
            ExcelFileManifestSunat = configuration["FilePaths:ExcelFileManifestSunat"];
            DownloadDirectory = configuration["FilePaths:DownloadDirectory"];
            ConsultaURL = configuration["ConsultaURL"];
            Codes = configuration.GetSection("Codes").Get<List<string>>();
        }
        
    }
    public class DetalleInfo
    {
        public string DetalleNum { get; set; }
        public string DescripcionText { get; set; }
        public string TamanoText { get; set; }
        public List<string> NumeroValues { get; set; } = new List<string>(); 
    }
    public class CategoryProductsDto
    {
        public string Category { get; set; }
        public List<ProductoDTO> Products { get; set; }
    }
	public class CodeAduana
	{
		public string codeAduana { get; set; }
	}
	public class ProductoDTO
	{
		public int IdProductStandarized { get; set; }  
		public string Category { get; set; }           
		public string ProductName { get; set; }         
		public DateTime DateRegister { get; set; }      
		public DateTime? DateUpdate { get; set; }
	}
}
