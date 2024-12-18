using OpenQA.Selenium;
using ROBOTPRUEBA_V1.CONFIG;
using ROBOTPRUEBA_V1.FILES.LOG;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ROBOTPRUEBA_V1.ADUANET.PAITA
{
    internal class Obtenerurldetalle
    {

        public async void ObtenerInfDetalle(IWebDriver driver, string digit)
        {
            driver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(120);
            WriteLog writeLog = new WriteLog();
            int maxAttempts = 3;
            int attempts = 0;
            bool elementFound1 = false;
            IWebElement headerRow=null;
            while (attempts < maxAttempts && !elementFound1)
            {
                try
                {
                    headerRow = driver.FindElement(By.XPath("//table[@cellpadding='1' and @cellspacing='1']//tr[1]"));
                    elementFound1 = true; 
                }
                catch (NoSuchElementException)
                {
                    attempts++;
                    if (attempts < maxAttempts)
                    {
                        driver.Navigate().Refresh();
						Task.Delay(1000).Wait();
					}
                    else
                    {
                        throw; 
                    }
                }
            }

            var headerCells = headerRow.FindElements(By.TagName("th"));

            int detalleColumnIndex = -1;
            for (int i = 0; i < headerCells.Count; i++)
            {
                if (headerCells[i].Text.Trim() == "Detalle")
                {
                    detalleColumnIndex = i;
                    break;
                }
            }

            var detalleCells = driver.FindElements(By.XPath($"//table[@cellpadding='1' and @cellspacing='1']//tr[position() > 1]/td[{detalleColumnIndex + 1}]"));
            string detalleimprimir="";
            for (int i = 0; i < detalleCells.Count; i++)
            {
                try
                {
                    detalleCells = driver.FindElements(By.XPath($"//table[@cellpadding='1' and @cellspacing='1']//tr[position() > 1]/td[{detalleColumnIndex + 1}]"));
					var detalleCell = detalleCells[i];
					if (detalleCell==null)
					{
						writeLog.Log($"No se encontraron cotenido en la celda de detalle numero {i}.");
						continue; // O maneja el error de otra manera
					}

                    var linkElement = detalleCell.FindElement(By.TagName("a"));
                    string DetalleNum = linkElement.Text.Trim();
                    detalleimprimir = linkElement.Text.Trim();
                    linkElement.Click();
                    Task.Delay(2000).Wait(); 

                    bool elementFound = false;
                    IWebElement descripcionHeader = null;
                    for (int attempt = 1; attempt <= 3; attempt++)
                    {
                        try
                        {
                            descripcionHeader = driver.FindElement(By.XPath("//th[contains(text(), 'Descripcion de Mercadería')]"));
                            elementFound = true;
                            break; 
                        }
                        catch (NoSuchElementException)
                        {
                            Task.Delay(2000).Wait();
                        }
                    }

                    if (!elementFound)
                    {
                        writeLog.Log(": No se encontró el elemento 'Descripcion de Mercadería' después de 3 intentos.\n");
                        
                        continue;
                    }


                    var descripcionTable = descripcionHeader.FindElement(By.XPath("ancestor::table"));
                    int descripcionColumnIndex = -1;
                    var descripcionHeaderCells = descripcionTable.FindElements(By.XPath(".//tr[1]/th"));
                    for (int j = 0; j < descripcionHeaderCells.Count; j++)
                    {
                        if (descripcionHeaderCells[j].Text.Trim() == "Descripcion de Mercadería") 
                        {
                            descripcionColumnIndex = j;
                            break;
                        }
                    }
                    string descripcionText = "-";
                    var descripcionCells = descripcionTable.FindElements(By.XPath($".//tr[position() > 1]/td[{descripcionColumnIndex + 1}]"));
                    foreach (var cell in descripcionCells)
                    {
                        if (!string.IsNullOrEmpty(cell.Text.Trim()) && cell.Text.Trim() != "...")
                        {
                            descripcionText = cell.Text.Trim();
                            break;
                        }
                    }
                    var tamanoHeader = driver.FindElement(By.XPath("//th[contains(text(), 'Tamaño')]"));
                    var tamanoTable = tamanoHeader.FindElement(By.XPath("ancestor::table"));

                    int tamanoColumnIndex = -1;
                    var tamanoHeaderCells = tamanoTable.FindElements(By.XPath(".//tr[1]/th"));
                    for (int k = 0; k < tamanoHeaderCells.Count; k++)
                    {
                        if (tamanoHeaderCells[k].Text.Trim() == "Tamaño")
                        {
                            tamanoColumnIndex = k;
                            break;
                        }
                    }
                    int numeroColumnIndex = -1;
                    var numeroHeaderCells = tamanoTable.FindElements(By.XPath(".//tr[1]/th"));
                    for (int m = 0; m < numeroHeaderCells.Count; m++)
                    {
                        if (numeroHeaderCells[m].Text.Trim() == "Número")
                        {
                            numeroColumnIndex = m;
                            break;
                        }
                    }
                    string tamanoText;
                    List<string> numeroValues = new List<string>();
                    string numeroText;
                    try
                    {
                        var tamanoCell = tamanoTable.FindElement(By.XPath($".//tr[position() > 1]/td[{tamanoColumnIndex + 1}]"));
                        tamanoText = !string.IsNullOrEmpty(tamanoCell.Text.Trim()) ? tamanoCell.Text.Trim() : "-";
                        var numeroCells = tamanoTable.FindElements(By.XPath($".//tr[position() > 1]/td[{numeroColumnIndex + 1}]"));
                        foreach (var numeroCell in numeroCells)
                        {
                            numeroText = !string.IsNullOrEmpty(numeroCell.Text.Trim()) ? numeroCell.Text.Trim() : "-";
                            numeroValues.Add(numeroText);
                        }
                    }
                    catch (NoSuchElementException)
                    {
                        tamanoText = "-";
                        numeroText = "-";
                        numeroValues.Add(numeroText);
                    }

                    if (!GlobalSettings.DetalleData.ContainsKey(digit))
                    {
                        GlobalSettings.DetalleData[digit] = new List<DetalleInfo>();
                    }

                    GlobalSettings.DetalleData[digit].Add(new DetalleInfo
                    {
                        DetalleNum = DetalleNum,
                        DescripcionText = descripcionText,
                        TamanoText = tamanoText,
                        NumeroValues = numeroValues 
                    });

                    driver.Navigate().Back();
                    Task.Delay(1500).Wait();
                }
                catch (NoSuchElementException ex)
                {
                    writeLog.Log($"No se encontró el elemento en la celda de detalle: {ex.Message}.-{ digit} --- {detalleimprimir}");
                }
                catch (Exception ex)
                {
                    writeLog.Log($"Ocurrió un error al procesar la celda de detalle: {ex.Message}");
					if (i == detalleCells.Count - 1)
					{
						break;
                    }
                    else
                    {
					    continue;
                    }
				}
            }
        }
    }
}
