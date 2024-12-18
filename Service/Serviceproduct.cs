using Microsoft.Extensions.Configuration;
using ROBOTPRUEBA_V1.CONFIG;
using ROBOTPRUEBA_V1.FILES.LOG;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Json;
using System.Text;
using System.Threading.Tasks;

namespace ROBOTPRUEBA_V1.Service
{
    internal class Serviceproduct
    {
        private readonly IConfiguration _configuration;
        private static readonly HttpClient _httpClient = new HttpClient();
        private static Dictionary<string, List<string>> _globalProductDictionary;
        public Serviceproduct()
        {
			var builder = new ConfigurationBuilder()
				.SetBasePath(AppDomain.CurrentDomain.BaseDirectory)
				.AddJsonFile("appsettings.json", optional: true, reloadOnChange: true);


			_configuration = builder.Build();
        }
        public async Task FetchAndStoreProductsAsync()
        {
            WriteLog writeLog = new WriteLog();
            string producturl = _configuration["Navigate:Api"] + "Products/LIST-PRODUCT-CATEGORY";
            try
            {
                var response = await _httpClient.GetFromJsonAsync<List<CategoryProductsDto>>(producturl).ConfigureAwait(false);
                if (response != null)
                {
                    GlobalSettings.GlobalProductDictionary = new Dictionary<string, List<string>>();
                    foreach (var categoryProduct in response)
                    {
                        if (!GlobalSettings.GlobalProductDictionary.ContainsKey(categoryProduct.Category))
                        {
                            GlobalSettings.GlobalProductDictionary[categoryProduct.Category] = new List<string>();
                        }

						GlobalSettings.GlobalProductDictionary[categoryProduct.Category].AddRange(categoryProduct.Products.Select(p => p.ProductName));
					}
                }
			}
            catch (HttpRequestException httpRequestException)
            {
                Console.WriteLine($"Error en la solicitud HTTP: {httpRequestException.Message}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ocurrió un error inesperado: {ex.Message}");
            }
        }
    }
}
