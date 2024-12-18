using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Extensions.Configuration;
using NPOI.SS.Formula.Functions;

namespace ROBOTPRUEBA_V1.SUNAT
{
    internal class Obtenerembarcador
    {
        private readonly IConfiguration _configuration;
        private readonly ConsultabyFecha consultabyFecha;
        private readonly LoginSunat loginSunat;
        private readonly NavigateSunat navigateSunat;
        public Obtenerembarcador() {
            var builder = new ConfigurationBuilder().SetBasePath(AppDomain.CurrentDomain.BaseDirectory).AddJsonFile("appsettings.json", optional: true, reloadOnChange: true);;

            _configuration = builder.Build();
            consultabyFecha = new ConsultabyFecha();
            loginSunat = new LoginSunat();
            navigateSunat = new NavigateSunat();
        }
        public async Task Obtener_embarcador(IWebDriver driver, WebDriverWait wait, CancellationToken stoppingToken)
        {
            driver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(120);
            string consultaManifiestoUrl = _configuration["Navigate:ConsultaManifiestoUrl"];
            string loginsunat = _configuration["Navigate:LoginSunat"];
            driver.Navigate().GoToUrl(loginsunat);
            Task.Delay(5000).Wait();
            loginSunat.Login(driver);
            navigateSunat.NavigateToExportPage(driver, wait, stoppingToken);
            consultabyFecha.ConsultaFecha(driver, wait, stoppingToken);
            Task.Delay(2000).Wait();
        }
    }
}
