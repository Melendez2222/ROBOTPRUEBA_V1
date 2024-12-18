using Microsoft.Extensions.Configuration;
using OfficeOpenXml;
using ROBOTPRUEBA_V1.CONFIG;
using ROBOTPRUEBA_V1.FILES.LOG;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http.Headers;
using System.Net.Http.Json;
using System.Text;
using System.Threading.Tasks;

namespace ROBOTPRUEBA_V1.Service
{
	internal class Serviceemail
	{
		private readonly IConfiguration _configuration;
		private static readonly HttpClient _httpClient = new HttpClient();
		public Serviceemail()
		{
			var builder = new ConfigurationBuilder().SetBasePath(AppDomain.CurrentDomain.BaseDirectory).AddJsonFile("appsettings.json", optional: true, reloadOnChange: true); ;

			_configuration = builder.Build();
		}
		public async Task CheckEmailServiceAsync()
		{
			string emialurl = _configuration["Navigate:Api"] + "Email_Service/SendFile";
			WriteLog writeLog = new WriteLog();
			int maxAttempts = 3;
			int currentAttempt = 0;
			while (currentAttempt < maxAttempts)
			{
				try
				{
					currentAttempt++;
					var response = await _httpClient.PostAsync(emialurl, null).ConfigureAwait(false);

					if (response.IsSuccessStatusCode)
					{
						writeLog.Log("Email service is operational.");
						return;
					}
					else
					{
						writeLog.Log($"Attempt {currentAttempt} failed. Status Code: {response.StatusCode}");
					}
				}
				catch (HttpRequestException httpRequestException)
				{
					writeLog.Log($"Error in attempt {currentAttempt}: {httpRequestException.Message}");
				}
				catch (Exception ex)
				{
					writeLog.Log($"Unexpected error in attempt {currentAttempt}: {ex.Message}");
				}
				await Task.Delay(2000);
			}

			writeLog.Log("Failed to check email service after multiple attempts.");
		}
		public async Task SendExcelFilesAsync()
		{
			string emialurl = _configuration["Navigate:Api"] + "Email_Service/SendFile";
			WriteLog writeLog = new WriteLog();
			try
			{

				string file1 = $"Reporte competencia semanal - Callao - semana {GlobalSettings.NumSemana}.xlsx";
				string file2 = $"Reporte competencia semanal Paita semana {GlobalSettings.NumSemana}.xlsx";
				string file3 = $"Reporte competencia semanal Pisco semana {GlobalSettings.NumSemana}.xlsx";
				string downloadDirectory1 = _configuration["FilePaths:ConvertFormatDirectoryPISCO"];
				string downloadDirectory2 = _configuration["FilePaths:ConvertFormatDirectoryPAITA"];
				string downloadDirectory3 = _configuration["FilePaths:ConvertFormatDirectoryCALLAO"];
				var files = new[]
					{
					FindLatestFile(downloadDirectory1, file3),
					FindLatestFile(downloadDirectory2, file2),
					FindLatestFile(downloadDirectory3, file1)
				};

				using var content = new MultipartFormDataContent();

				foreach (var file in files)
				{
					if (file != null)
					{
						var fileContent = new ByteArrayContent(await File.ReadAllBytesAsync(file));
						fileContent.Headers.ContentType = MediaTypeHeaderValue.Parse("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
						content.Add(fileContent, "files", Path.GetFileName(file));
					}
				}

				var response = await _httpClient.PostAsync(emialurl, content).ConfigureAwait(false);

				if (response.IsSuccessStatusCode)
				{
					writeLog.Log("Excel files sent successfully.");
				}
				else
				{
					writeLog.Log($"Failed to send Excel files. Status Code: {response.StatusCode}");
				}
			}
			catch (Exception ex)
			{
				writeLog.Log($"Fallo al intentar enviar correo {ex}");
			}
		}
		public async Task SendErrorMessageAsync(string errorMessage)
		{
			string errorUrl = _configuration["Navigate:Api"] + "Email_Service/SendText";
			WriteLog writeLog = new WriteLog();
			try
			{
				// Crea el contenido del mensaje
				var content = new StringContent(errorMessage, Encoding.UTF8, "text/plain");

				// Enviar la solicitud POST al endpoint
				var response = await _httpClient.PostAsync(errorUrl, content).ConfigureAwait(false);

				if (response.IsSuccessStatusCode)
				{
					writeLog.Log("Error message sent successfully.");
				}
				else
				{
					writeLog.Log($"Failed to send error message. Status Code: {response.StatusCode}");
				}
			}
			catch (Exception ex)
			{
				writeLog.Log($"Failed to send error message. Exception: {ex.Message}");
			}
		}

		private string FindLatestFile(string directory, string fileNamePattern)
		{
			var directoryInfo = new DirectoryInfo(directory);
			var file = directoryInfo.GetFiles(fileNamePattern)
									.OrderByDescending(f => f.LastWriteTime)
									.FirstOrDefault();
			return file?.FullName;
		}
	}
}
