using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium;
using SeleniumExtras.WaitHelpers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ROBOTPRUEBA_V1.SUNAT
{
    internal class NavigateSunat
    {
        public void NavigateToExportPage(IWebDriver driver, WebDriverWait wait, CancellationToken stoppingToken)
        {

            var opComerext = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("divOpcionServicio3")));
            opComerext.Click();

            var maniCargSal = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("nivel2_28_2")));
            maniCargSal.Click();

            var consulta = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("nivel3_28_2_2")));
            consulta.Click();

            var trazManCarga = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("nivel4_28_2_2_1_1")));
            trazManCarga.Click();


        }
    }
}
