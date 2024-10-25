using Microsoft.Extensions.Configuration;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;
using ROBOTPRUEBA_V1.CONFIG;
using ROBOTPRUEBA_V1.FILES.LOG;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium.Support.UI;
using OfficeOpenXml;
using NPOI.HSSF.UserModel;
using System.IO;
using System.Diagnostics;
using SeleniumExtras.WaitHelpers;
using NPOI.SS.UserModel;
using System.Security.Policy;
using HtmlAgilityPack;

namespace ROBOTPRUEBA_V1.ADUANET
{
    internal class NavigateConsulta
    {
        private readonly IConfiguration _configuration;
        private readonly List<string> _codes;

        public NavigateConsulta()
        {
            var builder = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("jsconfig1.json", optional: false, reloadOnChange: true);

            _configuration = builder.Build();
            _codes = _configuration.GetSection("Codes").Get<List<string>>();
        }
        public void NavigateConsultaCodManifiesto(IWebDriver driver, WebDriverWait wait, CancellationToken stoppingToken)
        {
            string consultaManifiestoUrl = _configuration["Navigate:ConsultaManifiestoUrl"];
            WriteLog writeLog = new WriteLog();
            string downloadDirectory = _configuration["FilePaths:DownloadDirectory"];

            if (GlobalSettings.ExtractedDigitsList.Count == 0)
            {
                writeLog.Log("No hay dígitos extraídos disponibles.");
                return;
            }


            try
            {
                using (var fileStream = new FileStream(GlobalSettings.ExcelFileManifestSunat, FileMode.OpenOrCreate, FileAccess.ReadWrite))
                {
                    var workbook = new HSSFWorkbook(fileStream);
                    var sheet = workbook.GetSheet("aduanet") ?? workbook.CreateSheet("aduanet");

                    int currentRow = sheet.LastRowNum + 1;
                    bool isFirstDigit = true;

                    foreach (var digit in GlobalSettings.ExtractedDigitsList)
                    {
                        driver.Navigate().GoToUrl(consultaManifiestoUrl);

                        Thread.Sleep(3000);
                        var inputField = driver.FindElement(By.Name("CMc1_Numero"));
                        inputField.SendKeys(digit);

                        var consultButton = driver.FindElement(By.XPath("//input[@value='Consultar']"));
                        consultButton.Click();
                        Thread.Sleep(10000);

                        string dateXpath = "//td/p/b[contains(text(), 'Fecha de Salida')]/../../following-sibling::td";
                        IWebElement dateElement;
                        try
                        {
                            dateElement = driver.FindElement(By.XPath(dateXpath));
                        }
                        catch (NoSuchElementException)
                        {
                            writeLog.Log($"Fecha de Salida no encontrada para el dígito: {digit}");
                            continue;
                        }
                        string fechaSalida = dateElement.Text.Trim();

                        string xpath = "//table[@width='100%'][contains(., 'Puerto ')]";
                        IWebElement table;
                        try
                        {
                            table = driver.FindElement(By.XPath(xpath));
                        }
                        catch (NoSuchElementException)
                        {
                            writeLog.Log($"Tabla no encontrada para el dígito: {digit}");
                            continue;
                        }

                        List<List<string>> data = new List<List<string>>();

                        IList<IWebElement> rows = table.FindElements(By.TagName("tr"));
                        foreach (IWebElement row in rows)
                        {
                            List<string> rowData = new List<string>();
                            IList<IWebElement> cells = row.FindElements(By.TagName("td"));
                            foreach (IWebElement cell in cells)
                            {
                                rowData.Add(cell.Text);
                            }
                            data.Add(rowData);
                        }

                        List<string> header = data[0];
                        List<List<string>> content = data.Skip(1).ToList();

                        var filteredData = content.Where(row => _codes.Contains(row[0])).ToList();
                        List<(string hrefConocimiento, string hrefDetalle, int rowIndex)> hrefs = new List<(string hrefConocimiento, string hrefDetalle, int rowIndex)>();
                        foreach (var row in filteredData)
                        {
                            var conocimientoCell = row[header.IndexOf("Conocimiento")];
                            var DetalleCell = row[header.IndexOf("Detalle")];
                            try
                            {
                                var conocimientoElement = table.FindElement(By.XPath($"//td/a/b[text()='{conocimientoCell}']"));
                                string hrefConocimiento = conocimientoElement.FindElement(By.XPath("..")).GetAttribute("href");
                                var detalleElement = table.FindElement(By.XPath($"//td/a/b[contains(text(), '{DetalleCell.Trim()}')]"));
                                string hrefDetalle = detalleElement.FindElement(By.XPath("..")).GetAttribute("href");
                                hrefs.Add((hrefConocimiento, hrefDetalle, currentRow));
                            }
                            catch (NoSuchElementException)
                            {
                                writeLog.Log($"No se encontró el enlace para el conocimiento: {conocimientoCell}");
                            }

                            var sheetRow = sheet.CreateRow(currentRow++);
                            for (int j = 0; j < row.Count; j++)
                            {
                                var cell = sheetRow.CreateCell(j);
                                cell.SetCellValue(row[j]);
                            }
                            var cellManifest = sheetRow.CreateCell(row.Count);
                            cellManifest.SetCellValue(digit);
                            var cellFecha = sheetRow.CreateCell(row.Count + 1);
                            cellFecha.SetCellValue(fechaSalida);
                            var P_Origen = sheetRow.CreateCell(row.Count + 2);
                            P_Origen.SetCellValue("CALLAO");
                        }


                        if (isFirstDigit)
                        {
                            var headerRow = sheet.CreateRow(currentRow++);
                            for (int j = 0; j < header.Count; j++)
                            {
                                var headerCell = headerRow.CreateCell(j);
                                headerCell.SetCellValue(header[j]);
                            }
                            var headerManifestCell = headerRow.CreateCell(header.Count);
                            headerManifestCell.SetCellValue("Nº Manifiesto");
                            var headerFechaCell = headerRow.CreateCell(header.Count + 1);
                            headerFechaCell.SetCellValue("Fecha");
                            var headerPOrigenCell = headerRow.CreateCell(header.Count + 2);
                            headerPOrigenCell.SetCellValue("Puerto de Origen");
                            var headerNDUAnCell = headerRow.CreateCell(header.Count + 3);
                            headerNDUAnCell.SetCellValue("Nº DUA");
                            var headerProductonCell = headerRow.CreateCell(header.Count + 4);
                            headerProductonCell.SetCellValue("PRODUCTO");
                            var headerTContenedornCell = headerRow.CreateCell(header.Count + 5);
                            headerTContenedornCell.SetCellValue("TAMAÑO DE CONTENEDOR");
                            var headerNContenedornCell = headerRow.CreateCell(header.Count + 6);
                            headerNContenedornCell.SetCellValue("Nº CONTENEDOR");
                            isFirstDigit = false;
                        }
                        foreach (var (hrefConocimiento, hrefDetalle, rowIndex) in hrefs)
                        {
                            driver.Navigate().GoToUrl(hrefConocimiento);
                            Thread.Sleep(5000);

                            // Encuentra la celda que contiene el encabezado "Numero DUA"
                            string headerXpath = "//table[@border='1' and @width='100%']//tr[1]/td[center/font/b[contains(text(), 'Numero DUA')]]";
                            IWebElement duaElement;
                            string numeroDUA;
                            try
                            {
                                IWebElement headerElement = driver.FindElement(By.XPath(headerXpath));
                                // Encuentra el índice de la columna del encabezado "Numero DUA"
                                int columnIndexD = int.Parse(headerElement.GetAttribute("cellIndex"));

                                // Usa el índice de la columna para encontrar el primer contenido de la columna "Numero DUA"
                                string duaXpath = $"//table[@border='1' and @width='100%']//tr[position()>1]/td[{columnIndexD + 1}]/font";
                                duaElement = driver.FindElement(By.XPath(duaXpath));
                                numeroDUA = duaElement.Text.Trim();
                                writeLog.Log($"Número DUA encontrado: {numeroDUA}");
                            }
                            catch (NoSuchElementException)
                            {
                                numeroDUA = "-";
                                writeLog.Log("Número DUA no encontrado, se guarda un guion '-'.");
                            }

                            var row = sheet.GetRow(rowIndex);
                            var cellDua = row.CreateCell(row.LastCellNum);
                            cellDua.SetCellValue(numeroDUA);


                            driver.Navigate().GoToUrl(hrefDetalle);
                            Thread.Sleep(5000); // Esperar a que la página cargue

                            string headerDescXpath = "//table[@border='1' and @width='100%']//tr[1]/td/font/b[contains(text(), 'Descripción de Mercadería')]";
                            IWebElement detalleElement;
                            string Producto;
                            try
                            {
                                IWebElement headerDescElement = driver.FindElement(By.XPath(headerDescXpath));
                                int columnDescIndex = int.Parse(headerDescElement.GetAttribute("cellIndex"));
                                string productXpath = $"//table[@border='1' and @width='100%']//tr[position()>1]/td[{columnDescIndex + 1}]/font";
                                detalleElement = driver.FindElement(By.XPath(productXpath));
                                Producto = detalleElement.Text.Trim();
                                writeLog.Log($"Detalle encontrado: {Producto}");
                            }
                            catch (NoSuchElementException)
                            {
                                Producto= "";
                                writeLog.Log("Detalle no encontrado, se guarda un guion '-'.");
                            }

                            // Añadir el "Detalle" o "-" a la hoja de trabajo en la fila correspondiente
                            var cellDetalle = row.CreateCell(row.LastCellNum + 1);
                            cellDetalle.SetCellValue(Producto);
                        }
                    }

                    using (var outputStream = new FileStream(GlobalSettings.ExcelFileManifestSunat, FileMode.Create, FileAccess.Write))
                    {
                        workbook.Write(outputStream);
                    }
                }

                writeLog.Log("Consulta realizada y datos exportados con éxito.");
            }
            catch (NoSuchElementException ex)
            {
                writeLog.Log($"El elemento no se encontró en la página: {ex.Message}");
            }

            Process.Start(new ProcessStartInfo(GlobalSettings.ExcelFileManifestSunat) { UseShellExecute = true });

        }
    }
}
