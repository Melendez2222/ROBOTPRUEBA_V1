using Microsoft.Extensions.Configuration;
using ROBOTPRUEBA_V1.FILES.LOG;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

namespace ROBOTPRUEBA_V1.Service
{
	internal class ServiceReportes
	{
		private readonly IConfiguration _configuration;
		private static readonly HttpClient _httpClient = new HttpClient();

		public ServiceReportes()
		{
			var builder = new ConfigurationBuilder()
				.SetBasePath(AppDomain.CurrentDomain.BaseDirectory)
				.AddJsonFile("appsettings.json", optional: true, reloadOnChange: true);

			_configuration = builder.Build();
		}
		public async Task EnviarDatosApiAsync(List<Dictionary<string, object>> datos)
		{
			string apiUrl = _configuration["Navigate:Api"] + "Report_Service/SendData";
			WriteLog writeLog = new WriteLog();
			int maxAttempts = 3;
			int currentAttempt = 0;
			bool success= false;
			while (currentAttempt < maxAttempts)
			{
				try
				{
					currentAttempt++;

					// Serializar los datos a JSON
					var jsonData = JsonSerializer.Serialize(datos);
					var content = new StringContent(jsonData, Encoding.UTF8, "application/json");

					// Enviar los datos a la API mediante una solicitud POST
					var response = await _httpClient.PostAsync(apiUrl, content).ConfigureAwait(false);

					if (response.IsSuccessStatusCode)
					{
						writeLog.Log("Datos enviados correctamente a la API.");
						success = true;
						return;
					}
					else
					{
						writeLog.Log($"Intento {currentAttempt} fallido. Código de estado: {response.StatusCode}");
					}
				}
				catch (HttpRequestException httpRequestException)
				{
					writeLog.Log($"Error en el intento {currentAttempt}: {httpRequestException.Message}");
				}
				catch (Exception ex)
				{
					writeLog.Log($"Error inesperado en el intento {currentAttempt}: {ex.Message}");
				}

				await Task.Delay(2000);
			}
			if (!success)
			{
				writeLog.Log("No se pudo enviar los datos a la API después de varios intentos.");
			}
		}

	}
}
