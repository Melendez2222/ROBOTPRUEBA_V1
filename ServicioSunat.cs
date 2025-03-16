using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium;
using ROBOTPRUEBA_V1.SUNAT;
using System.Configuration;
using Microsoft.Extensions.Configuration;
using ROBOTPRUEBA_V1.FILES.LOG;
using ROBOTPRUEBA_V1.ADUANET.CALLAO;
using ROBOTPRUEBA_V1.ADUANET.PAITA;
using ROBOTPRUEBA_V1.CONFIG;
using ROBOTPRUEBA_V1.ADUANET.PISCO;
using ROBOTPRUEBA_V1.Service;
using System.Diagnostics;
namespace ROBOTPRUEBA_V1
{
	internal class ServicioSunat
	{
		private readonly IConfiguration _configuration;
		private readonly Dictionary<string, string> _codeSelec;
		private readonly LoginSunat loginSunat;
		private readonly NavigateSunat navigateSunat;
		private readonly ExportSunat exportSunat;
		private readonly NavigateConsulta navigateConsulta;
		private readonly NavigateConsultaPaita navigateConsultaPaita;
		public readonly NavigateConsultPisco navigateConsultPisco;
		private readonly PRUEBANAVIGATE pRUEBANAVIGATE;
		private readonly Serviceemail serviceEmail;
		private readonly ServicePuertos servicePuertos;
		public ServicioSunat()
		{
			var builder = new ConfigurationBuilder()
				.SetBasePath(AppDomain.CurrentDomain.BaseDirectory)
				.AddJsonFile("appsettings.json", optional: true, reloadOnChange: true);


			_configuration = builder.Build();
			_codeSelec = _configuration.GetSection("CODE_SELEC").Get<Dictionary<string, string>>();
			loginSunat = new LoginSunat();
			navigateSunat = new NavigateSunat();
			exportSunat = new ExportSunat();
			navigateConsulta = new NavigateConsulta();
			navigateConsultaPaita = new NavigateConsultaPaita();
			navigateConsultPisco = new NavigateConsultPisco();
			pRUEBANAVIGATE = new PRUEBANAVIGATE();
			servicePuertos = new ServicePuertos();
			serviceEmail = new Serviceemail();
		}
		public async Task Run(CancellationToken stoppingToken)
		{
			VerifGeckodriver.fingecko();
			GlobalSettings.logFilePath = _configuration["FilePaths:Log"];
			WriteLog writeLog = new WriteLog();
			var serviceProduct = new Serviceproduct();
			var servicioemail = new Serviceemail();
			try
			{

				await serviceProduct.FetchAndStoreProductsAsync();


				await servicePuertos.LoadCodeAduanaListAsync();
			}
			catch (Exception ex)
			{
				writeLog.Log($"Error en llamado api : {ex}");
			}
			var productDictionary = GlobalSettings.GlobalProductDictionary;
			bool firstcode = true;
			Fechas_range.SetFechas();

			string defaultDownloadDirectory = _configuration["FilePaths:DefaultDownloadDirectory"];
			string LoginUrl = _configuration["Navigate:LoginSunat"];
			string driverPath = AppDomain.CurrentDomain.BaseDirectory;
			string firefoxPath = @"C:\Program Files\Mozilla Firefox\firefox.exe";
			string firefoxVersion = "127.0";
			foreach (var code in _codeSelec)
			{
				var service = FirefoxDriverService.CreateDefaultService(driverPath);
				service.HideCommandPromptWindow = false;


				var options = new FirefoxOptions();
				options.BrowserExecutableLocation = firefoxPath;
				options.AddArgument($"--version={firefoxVersion}");
				options.SetPreference("browser.download.folderList", 2);
				options.SetPreference("browser.download.dir", defaultDownloadDirectory);
				options.SetPreference("download.prompt_for_download", false);
				options.SetPreference("marionette.logging", false);
				options.SetPreference("disable-popup-blocking", "true");
				options.AddArgument("--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3");
				options.SetPreference("pdfjs.disabled", true);
				options.SetPreference("browser.helperApps.neverAsk.saveToDisk", "application/vnd.ms-excel");
				options.SetPreference("browser.download.manager.showWhenStarting", false);
				options.SetPreference("browser.download.manager.focusWhenStarting", false);
				options.SetPreference("browser.download.useDownloadDir", true);
				options.SetPreference("browser.helperApps.alwaysAsk.force", false);
				options.SetPreference("browser.download.manager.alertOnEXEOpen", false);
				options.SetPreference("browser.download.manager.closeWhenDone", true);
				options.SetPreference("browser.download.manager.showAlertOnComplete", false);
				options.SetPreference("browser.download.manager.useWindow", false);
				options.SetPreference("services.sync.prefs.sync.browser.download.manager.showWhenStarting", false);
				//options.AddArgument("--headless");
				options.PageLoadStrategy = PageLoadStrategy.Normal;

				try
				{

					using (IWebDriver driver = new FirefoxDriver(service, options))
					{
						try
						{


							writeLog.Log($"SE inicio el proceso para {code.Key}");
							try
							{
								GlobalSettings.CurrentCode = code.Key;

								GlobalSettings.DetalleData.Clear();
								GlobalSettings.Detalleurls.Clear();

								int maxAttemptsSUNAT = 3;
								int attemptSUNAT = 0;
								bool succesSUNAT = false;

								while (attemptSUNAT < maxAttemptsSUNAT && !succesSUNAT)
								{
									try
									{
										attemptSUNAT++;

										driver.Navigate().GoToUrl(LoginUrl);
										succesSUNAT = true;
									}
									catch (WebDriverTimeoutException ex)
									{
										writeLog.Log($"Error en el intento {attemptSUNAT} de navegar al aduanet para el manifiesto : {ex.Message}");

										Task.Delay(5000).Wait();
									}
									catch (Exception ex)
									{
										writeLog.Log($"Error inesperado en el intento {attemptSUNAT}: {ex.Message}");
									}
								}
								if (!succesSUNAT)
								{
									writeLog.Log($"Error en navegar al aduanet para el manifiesto ");
									continue;
								}

								var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(15));
								Task.Delay(6000).Wait();

								await loginSunat.Login(driver);

								navigateSunat.NavigateToExportPage(driver, wait, stoppingToken);
								await exportSunat.ExportData(driver, wait, stoppingToken);
								Task.Delay(6000).Wait();
								switch (code.Key)
								{
									case "CALLAO":
										pRUEBANAVIGATE.NavigateConsultaCodManifiestoNUEVO(driver, wait, stoppingToken);
										break;
									case "PAITA":
										navigateConsultaPaita.NavPaita(driver, wait, stoppingToken);
										break;
									case "PISCO":
										navigateConsultPisco.NavPisco(driver, wait, stoppingToken);
										break;


									default:
										writeLog.Log($"No hay valores de codigo de aduana en appsettings.json o no coinciden.");
										break;
								}
							}
							catch (WebDriverTimeoutException ex)
							{
								writeLog.Log($"Timeout occurred for code {code.Key}: {ex.Message}");
								continue;
							}
							catch (Exception ex)
							{
								writeLog.Log($"Error processing code {code.Key}: {ex.Message}");
								continue;
							}
							firstcode = false;


						}
						catch (Exception ex)
						{
							writeLog.Log($"OCURRIO UN EROR EN EL PROCESO DE SERVICIO SUNAT: {ex}");
							await serviceEmail.SendErrorMessageAsync($"OCURRIO UN EROR EN EL PROCESO DE SERVICIO SUNAT: {ex}");
						}
						finally
						{
							driver?.Quit(); // Intenta cerrar el driver

							VerifGeckodriver.fingecko();
						}
					}
					
				}
				catch (Exception ex)
				{
					writeLog.Log($"ERROR CON EL USING {ex}");
				}
				VerifGeckodriver.fingecko();
			}
			await servicioemail.SendExcelFilesAsync();
		}
	}
}
