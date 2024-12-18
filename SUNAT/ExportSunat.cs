using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium;
using SeleniumExtras.WaitHelpers;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using ROBOTPRUEBA_V1.FILES.Excel;
using ROBOTPRUEBA_V1.CONFIG;
using Microsoft.Extensions.Configuration;
using ROBOTPRUEBA_V1.FILES.LOG;

namespace ROBOTPRUEBA_V1.SUNAT
{
    internal class ExportSunat
    {
        private readonly IConfiguration configuration;
        public ExportSunat() {
            var builder = new ConfigurationBuilder().SetBasePath(AppDomain.CurrentDomain.BaseDirectory).AddJsonFile("appsettings.json", optional: true, reloadOnChange: true);;

            configuration = builder.Build();
        }
        
        public async Task ExportData(IWebDriver driver, WebDriverWait wait, CancellationToken stoppingToken)
        {
            WriteLog writeLog = new WriteLog();
            driver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(120);
            ManifiestoCarga manifiestoCarga = new ManifiestoCarga();
            ConsultabyFecha consultabyFecha = new ConsultabyFecha();
            await consultabyFecha.ConsultaFecha(driver, wait, stoppingToken);
            Task.Delay(3000).Wait();
			int maxAttempts = 5;
			int attempt = 0;
			IWebElement exportButton = null;

			while (attempt < maxAttempts && exportButton == null)
			{
				try
				{
					WebDriverWait waits = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
					exportButton = waits.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.Id("btnExportarExcel")));
					
				}
				catch (NoSuchElementException)
				{
					attempt++;
					writeLog.Log($"Intento {attempt} fallido. Esperando 5 segundos antes de reintentar Exportar excel reporte competencia semanal...");
					Task.Delay(5000).Wait();
				}
			}

			if (exportButton != null)
			{
				exportButton.Click();
			}
			else
			{
				writeLog.Log("No se pudo encontrar el botón para exportar reporte competencia semanal después de varios intentos.");
			}

			Task.Delay(10000).Wait();
            switch (GlobalSettings.CurrentCode)
            {
                case "CALLAO":
                    manifiestoCarga.Download();
                    break;
                case "PAITA":
                    manifiestoCarga.Download_PAITA();
                    break;
                case "PISCO":
                    manifiestoCarga.Download_PISCO();
                    break;
                default:
                    throw new Exception($"No se encontró un método de descarga para el código {GlobalSettings.CurrentCode}");
            }
        }
    }
}
