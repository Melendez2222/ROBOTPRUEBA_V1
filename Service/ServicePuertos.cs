using Microsoft.Extensions.Configuration;
using Newtonsoft.Json;
using ROBOTPRUEBA_V1.CONFIG;
using ROBOTPRUEBA_V1.FILES.LOG;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ROBOTPRUEBA_V1.Service
{
	internal class ServicePuertos
	{
		private readonly IConfiguration _configuration;
		public ServicePuertos()
		{
			var builder = new ConfigurationBuilder()
					.SetBasePath(AppDomain.CurrentDomain.BaseDirectory)
					.AddJsonFile("appsettings.json", optional: true, reloadOnChange: true);


			_configuration = builder.Build();
		}
		public async Task LoadCodeAduanaListAsync()
		{
			WriteLog writeLog = new WriteLog();
			string puertosurl = _configuration["Navigate:Api"] + "Puertos";
			using (HttpClient client = new HttpClient())
			{
				HttpResponseMessage response = await client.GetAsync(puertosurl);
				if (response.IsSuccessStatusCode)
				{
					string jsonResponse = await response.Content.ReadAsStringAsync();
					var codeAduanaList = JsonConvert.DeserializeObject<List<CodeAduana>>(jsonResponse);
					GlobalSettings.CodeAduanaList = codeAduanaList.Select(c => c.codeAduana).ToList();
				}
			}
		}
	}
}
