using Microsoft.Extensions.Configuration;
using OpenQA.Selenium;
using ROBOTPRUEBA_V1.CONFIG;
using ROBOTPRUEBA_V1.FILES.LOG;
using OpenQA.Selenium.Support.UI;
using OfficeOpenXml;
using System.Diagnostics;
using System.Data;
using OfficeOpenXml.Table;
using NPOI.SS.Formula.Functions;

namespace ROBOTPRUEBA_V1.ADUANET.CALLAO
{
    internal class NavigateConsulta
    {
        private readonly IConfiguration _configuration;
        private readonly List<string> _codes;

        public NavigateConsulta()
        {
            var builder = new ConfigurationBuilder()
                .SetBasePath(AppDomain.CurrentDomain.BaseDirectory)
                .AddJsonFile("appsettings.json", optional: true, reloadOnChange: true);;

            _configuration = builder.Build();
            _codes = _configuration.GetSection("Codes").Get<List<string>>();
        }
        public void NavigateConsultaCodManifiesto(IWebDriver driver, WebDriverWait wait, CancellationToken stoppingToken)
        {
            string consultaManifiestoUrl = _configuration["Navigate:ConsultaManifiestoUrl"];
            WriteLog writeLog = new WriteLog();
            string downloadDirectory = _configuration["FilePaths:ConvertFormatDirectory"];

            if (GlobalSettings.ExtractedDigitsList.Count == 0)
            {
                writeLog.Log("No hay dígitos extraídos disponibles.");
                return;
            }


            try
            {
                using (var package = new ExcelPackage(new FileInfo(GlobalSettings.ExcelFileManifestSunat)))
                {
                    var workbook = package.Workbook;
                    var sheet = workbook.Worksheets.FirstOrDefault(ws => ws.Name == "aduanet") ?? workbook.Worksheets.Add("aduanet");

                    int currentRow = sheet.Dimension?.End.Row + 1 ?? 1;
                    bool isFirstDigit = currentRow == 1;
                    List<string> header = new List<string>();

                    foreach (var digit in GlobalSettings.ExtractedDigitsList)
                    {
                        driver.Navigate().GoToUrl(consultaManifiestoUrl);

                        Task.Delay(3000).Wait();
                        var inputField = driver.FindElement(By.Name("CMc1_Numero"));
                        inputField.SendKeys(digit.Key);

                        var consultButton = driver.FindElement(By.XPath("//input[@value='Consultar']"));
                        consultButton.Click();
                        Task.Delay(10000).Wait();

                        string dateXpath = "//td/p/b[contains(text(), 'Fecha de Salida')]/../../following-sibling::td";
						IWebElement dateElement = null;
						int maxRetriesdataelement = 5; // Número máximo de intentos
						int retryCountdataelement = 0;
						bool elementFounddataelement = false;

						while (retryCountdataelement < maxRetriesdataelement && !elementFounddataelement)
						{
							try
							{
								// Intentar encontrar el elemento
								dateElement = driver.FindElement(By.XPath(dateXpath));
								elementFounddataelement = true; // Si se encuentra, marcar como éxito
							}
							catch (NoSuchElementException)
							{
								retryCountdataelement++; // Incrementar contador de intentos
								writeLog.Log($"Intento {retryCountdataelement}/{maxRetriesdataelement}: Fecha de Salida no encontrada para el dígito: {digit}");

								// Esperar un tiempo antes del próximo intento
								Task.Delay(1000).Wait(); // 1000 ms = 1 segundo (ajustar según sea necesario)
							}
						}

						// Verificar si el elemento fue encontrado después de los intentos
						if (!elementFounddataelement)
						{
							writeLog.Log($"No se pudo encontrar contenido en la pagina del manifiesto después de {maxRetriesdataelement} intentos para el dígito: {digit}");
							// Continuar con la lógica deseada (ejemplo: continuar con el siguiente dígito)
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

                        Dictionary<string, List<(string Text, string Href)>> data = new Dictionary<string, List<(string Text, string Href)>>();

                        IList<IWebElement> rows = table.FindElements(By.TagName("tr"));

                        foreach (IWebElement row in rows)
                        {
                            IList<IWebElement> cells = row.FindElements(By.TagName("td"));

                            if (header.Count == 0)
                            {
                                foreach (var cell in cells)
                                {
                                    string columnName = cell.Text.Trim();
                                    if (!header.Contains(columnName))
                                    {
                                        header.Add(columnName);
                                    }
                                    else
                                    {
                                        writeLog.Log($"Clave duplicada encontrada en el encabezado: {columnName}");
                                    }
                                }
                                header.Add("N° de Manifiesto");
                                header.Add("Fecha");
                                header.Add("Puerto de Origen");
                                header.Add("Nº DUA");
                                header.Add("Producto");
                                header.Add("Tamaño de Contenedor");
                                header.Add("N° de Contenedor");
                                continue;
                            }

                            for (int i = 0; i < cells.Count; i++)
                            {
                                string cellText = cells[i].Text.Trim();
                                string cellHref = "-";
                                try
                                {
                                    var linkElement = cells[i].FindElement(By.TagName("a"));
                                    cellHref = linkElement.GetAttribute("href");
                                }
                                catch (NoSuchElementException)
                                {
                                }

                                string columnName = header[i];

                                if (!data.ContainsKey(columnName))
                                {
                                    data[columnName] = new List<(string Text, string Href)>();
                                }
                                data[columnName].Add((cellText, cellHref));
                            }
                        }

                        var filteredRows = data["Puerto"]
                            .Select((value, index) => new { value, index })
                            .Where(x => _codes.Contains(x.value.Text))
                            .Select(x => x.index)
                            .ToList();

                        Dictionary<string, List<(string Text, string Href)>> filteredData = new Dictionary<string, List<(string Text, string Href)>>();

                        foreach (var key in data.Keys)
                        {
                            filteredData[key] = filteredRows.Select(index => data[key][index]).ToList();
                        }

                        bool hasData = filteredData.Values.Any(list => list.Count > 0);

                        if (!hasData)
                        {
                            writeLog.Log($"No se encontraron datos filtrados para el digito: {digit}");
                            continue;
                        }


                        if (isFirstDigit)
                        {

                            int headerColumn = 1;
                            foreach (var key in filteredData.Keys)
                            {
                                sheet.Cells[currentRow, headerColumn].Value = key;
                                headerColumn++;
                            }
                            sheet.Cells[currentRow, headerColumn].Value = "N° de Manifiesto";
                            headerColumn++;
                            sheet.Cells[currentRow, headerColumn].Value = "Fecha";
                            headerColumn++;
                            sheet.Cells[currentRow, headerColumn].Value = "Puerto de Origen";
                            headerColumn++;
                            sheet.Cells[currentRow, headerColumn].Value = "Nº DUA";
                            headerColumn++;
                            sheet.Cells[currentRow, headerColumn].Value = "Producto";
                            headerColumn++;
                            sheet.Cells[currentRow, headerColumn].Value = "Tamaño de Contenedor";
                            headerColumn++;
                            sheet.Cells[currentRow, headerColumn].Value = "N° de Contenedor";
                            currentRow++;
                            isFirstDigit = false;
                        }

                        int dataRow = currentRow;
                        try
                        {
                            int maxRows = filteredData.Values.Max(list => list.Count);

                            for (int i = 0; i < maxRows; i++)
                            {
                                int dataColumn = 1;
                                foreach (var key in filteredData.Keys)
                                {
                                    if (i < filteredData[key].Count)
                                    {
                                        var cellData = filteredData[key][i];
                                        sheet.Cells[dataRow, dataColumn].Value = cellData.Text;
                                        if (cellData.Href != "-")
                                        {
                                            sheet.Cells[dataRow, dataColumn].Hyperlink = new ExcelHyperLink(cellData.Href);
                                        }
                                    }
                                    dataColumn++;
                                }
                                sheet.Cells[dataRow, dataColumn].Value = digit;
                                dataColumn++;
                                sheet.Cells[dataRow, dataColumn].Value = fechaSalida;
                                dataColumn++;
                                sheet.Cells[dataRow, dataColumn].Value = "CALLAO";
                                dataRow++;
                            }
                        }
                        catch (Exception ex)
                        {
                            writeLog.Log($"Error al escribir los datos en el archivo Excel: {ex.Message}");
                        }
                        int columnIndexDUA = header.IndexOf("Nº DUA") + 1;
                        int columnIndexProducto = header.IndexOf("Producto") + 1;
                        int columnIndexTamañoContenedor = header.IndexOf("Tamaño de Contenedor") + 1;
                        int columnIndexNumeroContenedor = header.IndexOf("N° de Contenedor") + 1;
                        for (int i = 0; i < filteredData["Conocimiento"].Count; i++)
                        {


                            var detailUrlConocimiento = filteredData["Conocimiento"][i].Href;

                            if (!string.IsNullOrEmpty(detailUrlConocimiento) && detailUrlConocimiento != "-")
                            {
                                driver.Navigate().GoToUrl(detailUrlConocimiento);
                                Task.Delay(5000).Wait();

                                string headerXpath = "//table[@border='1' and @width='100%']//tr[1]/td[center/font/b[contains(text(), 'Numero DUA')]]";
                                IWebElement duaElement;
                                string numeroDUA;
                                try
                                {
                                    IWebElement headerElement = driver.FindElement(By.XPath(headerXpath));
                                    int columnIndexD = int.Parse(headerElement.GetAttribute("cellIndex"));

                                    string duaXpath = $"//table[@border='1' and @width='100%']//tr[position()>1]/td[{columnIndexD + 1}]/font";
                                    duaElement = driver.FindElement(By.XPath(duaXpath));
                                    numeroDUA = duaElement.Text.Trim();
                                }
                                catch (NoSuchElementException)
                                {
                                    numeroDUA = "-";
                                    writeLog.Log("Número DUA no encontrado, se guarda un guion '-'.");
                                }

                                var cellNumeroDUA = sheet.Cells[currentRow + i, columnIndexDUA];
                                cellNumeroDUA.Value = numeroDUA;
                            }
                        }

                        for (int i = 0; i < filteredData["Detalle"].Count; i++)
                        {
                            var detailUrlDetalle = filteredData["Detalle"][i].Href;
                            if (!string.IsNullOrEmpty(detailUrlDetalle) && detailUrlDetalle != "-")
                            {
                                driver.Navigate().GoToUrl(detailUrlDetalle);
                                Task.Delay(5000).Wait();
                                string headerDescXpath = "//table[@border='1' and @width='100%']//tr[1]/td/font/b[contains(text(), 'Descripción de Mercadería')]";
								string headerMarcasXpath = "//table[@border='1' and @width='100%']//tr[1]/td/font/b[contains(text(), 'Marcas y Números')]";
								IWebElement detalleElement;
                                string Producto;
                                try
                                {
                                    IWebElement headerDescElement = driver.FindElement(By.XPath(headerDescXpath));
                                    IList<IWebElement> headerDescCells = driver.FindElements(By.XPath("//table[@border='1' and @width='100%']//tr[1]/td"));
                                    int columnDescIndex = -1;
                                    for (int j = 0; j < headerDescCells.Count; j++)
                                    {
                                        if (headerDescCells[j].Text.Contains("Descripción de Mercadería"))
                                        {
                                            columnDescIndex = j;
                                            break;
                                        }
                                    }
                                    string productXpath = $"//table[@border='1' and @width='100%']//tr[position()>1]/td[{columnDescIndex + 1}]/font";
                                    detalleElement = driver.FindElement(By.XPath(productXpath));
                                    Producto = detalleElement.Text.Trim();
									if (Producto == ". ." || Producto == "..." || Producto == ". . ." || Producto == "..")
									{
										throw new NoSuchElementException();
									}
								}
                                catch (NoSuchElementException)
                                {
									try
									{
										IWebElement headerMarcasElement = driver.FindElement(By.XPath(headerMarcasXpath));
										IList<IWebElement> headerMarcasCells = driver.FindElements(By.XPath("//table[@border='1' and @width='100%']//tr[1]/td"));
										int columnMarcasIndex = -1;
										for (int j = 0; j < headerMarcasCells.Count; j++)
										{
											if (headerMarcasCells[j].Text.Contains("Marcas y Números"))
											{
												columnMarcasIndex = j;
												break;
											}
										}
										string marcasXpath = $"//table[@border='1' and @width='100%']//tr[position()>1]/td[{columnMarcasIndex + 1}]/font";
										detalleElement = driver.FindElement(By.XPath(marcasXpath));
										Producto = detalleElement.Text.Trim();
									}
									catch (NoSuchElementException)
									{
										Producto = "-";
										writeLog.Log("Detalle no encontrado, se guarda un guion '-'.");
									}
								}

                                var cellDetalle = sheet.Cells[currentRow + i, columnIndexProducto];
                                cellDetalle.Value = Producto;

                                string tableXpath = "//table[@width='80%' and @border='']";
                                IWebElement tableElement;
                                try
                                {
                                    tableElement = driver.FindElement(By.XPath(tableXpath));
                                }
                                catch (NoSuchElementException)
                                {
                                    writeLog.Log("Tabla específica no encontrada.");
                                    continue;
                                }

                                IList<IWebElement> tableHeaders = tableElement.FindElements(By.XPath(".//tr[1]/td"));
                                int columnIndexTamaño = -1;
                                int columnIndexNum = -1;
                                for (int j = 0; j < tableHeaders.Count; j++)
                                {
                                    if (tableHeaders[j].Text.Contains("Número"))
                                    {
                                        columnIndexNum = j;
                                    }
                                    else if (tableHeaders[j].Text.Contains("Tamaño"))
                                    {
                                        columnIndexTamaño = j;
                                        break;
                                    }
                                }
                                if (columnIndexTamaño == -1)
                                {
                                    writeLog.Log("Columna 'Tamaño' no encontrada.");
                                    continue;
                                }
                                string tamañoValue = tableElement.FindElement(By.XPath($".//tr[2]/td[{columnIndexTamaño + 1}]")).Text.Trim();
                                var cellTamaño = sheet.Cells[currentRow + i, columnIndexTamañoContenedor];
                                cellTamaño.Value = tamañoValue;

                                IList<IWebElement> tableRows = tableElement.FindElements(By.XPath(".//tr[position()>1]"));
                                int numRows = tableRows.Count;
                                int totalColumns = header.Count;

                                List<string> numeros = new List<string>();
                                for (int j = 0; j < tableRows.Count; j++)
                                {
                                    string numeroValue = tableRows[j].FindElement(By.XPath($".//td[{columnIndexNum + 1}]")).Text.Trim();
                                    numeros.Add(numeroValue);
                                }

                                sheet.Cells[currentRow + i, columnIndexNumeroContenedor].Value = numeros[0];

                                if (numRows > 1)
                                {
                                    for (int j = 1; j < numRows; j++)
                                    {
                                        int targetRow = currentRow + i + j;
                                        sheet.InsertRow(targetRow, 1);

                                        for (int k = 1; k <= totalColumns - 1; k++)
                                        {
                                            sheet.Cells[targetRow, k].Value = sheet.Cells[currentRow + i, k].Value;
                                            if (sheet.Cells[currentRow + i, k].Hyperlink != null)
                                            {
                                                sheet.Cells[targetRow, k].Hyperlink = sheet.Cells[currentRow + i, k].Hyperlink;
                                            }
                                        }
                                        sheet.Cells[targetRow, columnIndexNumeroContenedor].Value = numeros[j];
                                    }
                                }
                                currentRow += numRows - 1;
                            }
                        }

                    }

                    if (sheet.Dimension != null)
                    {
                        var range = sheet.Cells[1, 1, sheet.Dimension.End.Row, sheet.Dimension.End.Column];
                        var table = sheet.Tables.Add(range, "Table1");
                        table.TableStyle = TableStyles.Medium5;
                        for (int col = 1; col <= sheet.Dimension.End.Column; col++)
                        {
                            sheet.Column(col).AutoFit();
                        }

                    }


                    package.Save();
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
