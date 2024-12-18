using Microsoft.Extensions.Configuration;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ROBOTPRUEBA_V1.SUNAT
{
    internal class LoginSunat
    {
        private readonly IConfiguration _configuration;

        public LoginSunat()
        {
            var builder = new ConfigurationBuilder().SetBasePath(AppDomain.CurrentDomain.BaseDirectory).AddJsonFile("appsettings.json", optional: true, reloadOnChange: true);;

            _configuration = builder.Build();
        }
        public async Task Login(IWebDriver driver)
        {
            string RUC = _configuration["Credentials:RUC"];
            string USUARIO = _configuration["Credentials:Users"];
            string PASSWORD = _configuration["Credentials:PASSWORD"];
            var rucOption = driver.FindElement(By.Id("btnPorRuc"));
            rucOption.Click();

            var rucInput = driver.FindElement(By.Id("txtRuc"));
            rucInput.SendKeys(RUC);

            var userInput = driver.FindElement(By.Id("txtUsuario"));
            userInput.SendKeys(USUARIO);
            Task.Delay(new Random().Next(1000, 3000)).Wait();
            var passwordInput = driver.FindElement(By.Id("txtContrasena"));
            passwordInput.SendKeys(PASSWORD);
            Task.Delay(new Random().Next(1000, 3000)).Wait();
            var loginButton = driver.FindElement(By.Id("btnAceptar"));
            loginButton.Click();
        }
    }
}
