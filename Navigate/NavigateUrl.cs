using DocumentFormat.OpenXml.Bibliography;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ROBOTPRUEBA_V1.Navigate
{
    internal class NavigateUrl
    {
        public async Task NavigUrl(string url, IWebDriver driver) {
            driver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(5);
        }
    }
}
