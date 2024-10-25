using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium;
using ROBOTPRUEBA_V1.SUNAT;
using System.Configuration;
using Microsoft.Extensions.Configuration;
using ROBOTPRUEBA_V1.ADUANET;
using ROBOTPRUEBA_V1.FILES.LOG;



namespace ROBOTPRUEBA_V1
{

    internal class ServicioSunat
    {
        private readonly IConfiguration _configuration;
        private readonly LoginSunat loginSunat;
        private readonly NavigateSunat navigateSunat;
        private readonly ExportSunat exportSunat;
        private readonly NavigateConsulta navigateConsulta;
        public ServicioSunat()
        {
            var builder = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("jsconfig1.json", optional: false, reloadOnChange: true);

            _configuration = builder.Build();
            loginSunat = new LoginSunat();
            navigateSunat = new NavigateSunat();
            exportSunat = new ExportSunat();
            navigateConsulta = new NavigateConsulta();
        }
        public async Task Run(CancellationToken stoppingToken)
        {

            WriteLog writeLog = new WriteLog();
            writeLog.Log("SE inicio el proceso");
            string defaultDownloadDirectory = _configuration["FilePaths:DefaultDownloadDirectory"];
            string LoginUrl = _configuration["Navigate:LoginSunat"];

            var options = new ChromeOptions();
            options.AddUserProfilePreference("download.default_directory", defaultDownloadDirectory);
            options.AddUserProfilePreference("download.prompt_for_download", false);
            options.AddUserProfilePreference("disable-popup-blocking", "true");

            using (IWebDriver driver = new ChromeDriver(options))
            {
                //driver.Manage().Window.Minimize();
                driver.Navigate().GoToUrl(LoginUrl);
                var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
                await Task.Delay(6000, stoppingToken);

                loginSunat.Login(driver);

                navigateSunat.NavigateToExportPage(driver, wait, stoppingToken);
                exportSunat.ExportData(driver, wait, stoppingToken);
                Thread.Sleep(5000);
                
                await Task.Delay(6000, stoppingToken);
                navigateConsulta.NavigateConsultaCodManifiesto(driver, wait, stoppingToken);

            }
        }






    }
}
