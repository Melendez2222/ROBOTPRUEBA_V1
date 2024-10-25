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

namespace ROBOTPRUEBA_V1.SUNAT
{
    internal class ExportSunat
    {
        
        public void ExportData(IWebDriver driver, WebDriverWait wait, CancellationToken stoppingToken)
        {
            ManifiestoCarga manifiestoCarga = new ManifiestoCarga();
            wait.Until(ExpectedConditions.FrameToBeAvailableAndSwitchToIt("iframeApplication"));

            var aduanaSelect = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("selCodigoAduana")));
            var selectElement = new SelectElement(aduanaSelect);
            selectElement.SelectByValue("118");
            Thread.Sleep(1000);
            var ViaSelect = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("selViaTransporte")));
            var selectElementVia = new SelectElement(ViaSelect);
            selectElementVia.SelectByValue("1");

            var radioButton = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("tipoBusquedaFechaTerminoEmbarque")));
            radioButton.Click();

            var fechaInicialInput = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("txtFechaInicial")));
            fechaInicialInput.Clear();
            fechaInicialInput.SendKeys("09/09/2024");

            var fechaFinalInput = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("txtFechaFinal")));
            fechaFinalInput.Clear();
            fechaFinalInput.SendKeys("09/09/2024");
            Thread.Sleep(10000);
            var consultabtn = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("btnBuscar")));
            consultabtn.Click();
            var waitExport = new WebDriverWait(driver, TimeSpan.FromSeconds(20));
            var exportButton = waitExport.Until(ExpectedConditions.ElementIsVisible(By.Id("btnExportarExcel")));
            exportButton.Click();

            Thread.Sleep(10000);
            manifiestoCarga.Download();
        }
    }
}
