using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium;
using ROBOTPRUEBA_V1.CONFIG;
using SeleniumExtras.WaitHelpers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Extensions.Configuration;
using ROBOTPRUEBA_V1.FILES.LOG;

namespace ROBOTPRUEBA_V1.SUNAT
{
    internal class ConsultabyFecha
    {
        private readonly IConfiguration _configuration;
        private readonly Dictionary<string, string> _codeSelec;
        public ConsultabyFecha() {
            var builder = new ConfigurationBuilder().SetBasePath(AppDomain.CurrentDomain.BaseDirectory).AddJsonFile("appsettings.json", optional: true, reloadOnChange: true);;
            _configuration = builder.Build();
            _codeSelec = _configuration.GetSection("CODE_SELEC").Get<Dictionary<string, string>>();
        }
        public async Task ConsultaFecha(IWebDriver driver, WebDriverWait wait, CancellationToken stoppingToken)
        {
            WriteLog writeLog = new WriteLog();
            wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.FrameToBeAvailableAndSwitchToIt("iframeApplication"));
            Task.Delay(2000).Wait();
			int maxAttemptsAduanaSelect = 3;
			int attemptsAduanaSelect = 0; 
			IWebElement aduanaSelect = null;

			// Bucle para intentar encontrar el elemento aduanaSelect
			while (attemptsAduanaSelect < maxAttemptsAduanaSelect)
			{
				try
				{
					// Intenta encontrar el elemento aduanaSelect
					aduanaSelect = driver.FindElement(By.Id("selCodigoAduana"));
					break; // Si se encuentra, salimos del bucle
				}
				catch (NoSuchElementException)
				{
					// Si no se encuentra, incrementa el contador de intentos
					attemptsAduanaSelect++;
					if (attemptsAduanaSelect < maxAttemptsAduanaSelect)
					{
						Task.Delay(3000).Wait(); // Espera el tiempo especificado usando Task.Delay
					}
					else
					{
						writeLog.Log("No se pudo encontrar el elemento 'selCodigoAduana' después de varios intentos.");
						// Aquí puedes manejar el caso en que no se encontró el elemento
					}
				}
			}

			// Si se encontró el elemento aduanaSelect, puedes trabajar con él
			if (aduanaSelect == null)
			{
				writeLog.Log("El elemento 'selCodigoAduana' no fue encontrado.");
				return;
			}
			Task.Delay(1000).Wait();
            var selectElement = new SelectElement(aduanaSelect);
            if (GlobalSettings.CurrentCode != null && _codeSelec.TryGetValue(GlobalSettings.CurrentCode, out string aduanaValue))
            {
                selectElement.SelectByValue(aduanaValue);
            }
            else
            {
                throw new Exception($"No se encontró un valor para el código {GlobalSettings.CurrentCode}");
            }

            Task.Delay(1000).Wait();
            var ViaSelect = driver.FindElement(By.Id("selViaTransporte"));
            Task.Delay(1000).Wait();
            var selectElementVia = new SelectElement(ViaSelect);
            selectElementVia.SelectByValue("1");

            var radioButton = driver.FindElement(By.Id("tipoBusquedaFechaTerminoEmbarque"));
            radioButton.Click();
            Task.Delay(1000).Wait();
            
			try
			{
				
				var fechaInicialInput = driver.FindElement(By.Id("txtFechaInicial"));
				fechaInicialInput.Clear();
				Task.Delay(1000).Wait();
				var ErrorRangoFecha = driver.FindElement(By.Id("dlgBtnAceptarConfirm"));
				ErrorRangoFecha.Click();
				Task.Delay(2000).Wait();
				fechaInicialInput.SendKeys(GlobalSettings.FechaInicio);
				Task.Delay(1000).Wait();
			}
			catch (Exception ex)
			{
				var fechaInicialInput = driver.FindElement(By.Id("txtFechaInicial"));
				fechaInicialInput.Clear();
				Task.Delay(1000).Wait();
				fechaInicialInput.SendKeys(GlobalSettings.FechaInicio);
				Task.Delay(1000).Wait();

			}
			Task.Delay(1000).Wait();
			try
			{
				
				var fechaFinalInput = driver.FindElement(By.Id("txtFechaFinal"));
				fechaFinalInput.Clear();
				Task.Delay(1000).Wait();
				var ErrorRangoFecha = driver.FindElement(By.Id("dlgBtnAceptarConfirm"));
				ErrorRangoFecha.Click();
				Task.Delay(2000).Wait();
				fechaFinalInput.SendKeys(GlobalSettings.FechaFin);
				Task.Delay(2000).Wait();
			}
			catch (Exception ex)
			{
				var fechaFinalInput = driver.FindElement(By.Id("txtFechaFinal"));
				fechaFinalInput.Clear();
				Task.Delay(1000).Wait();
				fechaFinalInput.SendKeys(GlobalSettings.FechaFin);
				Task.Delay(1000).Wait();

			}

			Task.Delay(1000).Wait();


			try
            {
                var ErrorRangoFecha = driver.FindElement(By.Id("dlgBtnAceptarConfirm"));
                ErrorRangoFecha.Click();
                Task.Delay(2000).Wait();
                var consultabtn = driver.FindElement(By.Id("btnBuscar"));
                consultabtn.Click();
            }
            catch (Exception ex)
            {
                var consultabtn = driver.FindElement(By.Id("btnBuscar"));
                consultabtn.Click();

            }

        }

    }
}
