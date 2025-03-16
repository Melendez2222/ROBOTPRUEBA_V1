using Microsoft.Extensions.Configuration;
using OfficeOpenXml;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium;
using ROBOTPRUEBA_V1.CONFIG;
using ROBOTPRUEBA_V1.FILES.LOG;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml.Table;
using System.Drawing;
using SeleniumExtras.WaitHelpers;
using NPOI.SS.Formula.Functions;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using NPOI.HPSF;
using DocumentFormat.OpenXml.Drawing.Diagrams;

namespace ROBOTPRUEBA_V1.SUNAT
{
    internal class PRUEBANAVIGATE
    {
        private readonly IConfiguration _configuration;
        private readonly List<string> _codes;
        private readonly Obtenerembarcador _obtenerembarcador;
        public PRUEBANAVIGATE()
        {
            var builder = new ConfigurationBuilder().SetBasePath(AppDomain.CurrentDomain.BaseDirectory).AddJsonFile("appsettings.json", optional: true, reloadOnChange: true);;

            _configuration = builder.Build();
            _codes = _configuration.GetSection("Codes").Get<List<string>>();
            _obtenerembarcador = new Obtenerembarcador();
        }
        public async void NavigateConsultaCodManifiestoNUEVO(IWebDriver driver, WebDriverWait wait, CancellationToken stoppingToken)
        {
			WriteLog writeLog1 = new WriteLog();

			driver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(120);
            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            string consultaManifiestoUrl = _configuration["Navigate:ConsultaManifiestoUrl"];
            WriteLog writeLog = new WriteLog();
            string downloadDirectory = _configuration["FilePaths:ConvertFormatDirectory"];

            if (GlobalSettings.ExtractedDigitsList.Count == 0)
            {
                writeLog.Log("No hay dígitos extraídos disponibles.");
                return;
            }
            Task.Delay(4000).Wait();
            #region MigrationTabletoExcel
            try
            {
                using (var package = new ExcelPackage(new FileInfo(GlobalSettings.ExcelFileManifestSunat)))
                {
					WebDriverWait waits = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
					var workbook = package.Workbook;
                    var sheet = workbook.Worksheets.FirstOrDefault(ws => ws.Name == "aduanet") ?? workbook.Worksheets.Add("aduanet");
                    int currentRow = sheet.Dimension?.End.Row + 1 ?? 1;
                    bool isFirstDigit = currentRow == 1;
                    List<string> header = new List<string>();
					foreach (var digit in GlobalSettings.ExtractedDigitsList)
                    {
						int maxAttemptsnavADUA = 3;
						int attemptnavADUA = 0;
						bool successnavADUA = false;
                        
                       
						while (attemptnavADUA < maxAttemptsnavADUA && !successnavADUA)
						{
							try
							{
								attemptnavADUA++;

								driver.Navigate().GoToUrl(consultaManifiestoUrl);
								successnavADUA = true;
							}
							catch (WebDriverTimeoutException ex)
							{
								writeLog.Log($"Error en el intento {attemptnavADUA} de navegar al aduanet para el manifiesto {digit}: {ex.Message}");

								Task.Delay(3000).Wait();
							}
							catch (Exception ex)
							{
								writeLog.Log($"Error inesperado en el intento {attemptnavADUA}: {ex.Message}");
							}
						}
						if (!successnavADUA)
						{
							writeLog.Log($"Error en navegar al aduanet para el manifiesto {digit}");
							continue;
						}
						Task.Delay(3000).Wait();
						var inputanno = driver.FindElement(By.Name("CMc1_Anno"));
						inputanno.Clear(); // Limpia el campo
						inputanno.SendKeys(digit.Value); // Envía el nuevo valor
						Task.Delay(3000).Wait();    
                        var inputField = driver.FindElement(By.Name("CMc1_Numero"));
                        inputField.SendKeys(digit.Key);

                        var consultButton = waits.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.XPath("//input[@value='Consultar']")));
                        consultButton.Click();
                        Task.Delay(4000).Wait();
                        
                        string dateXpath = "//td/p/b[contains(text(), 'Fecha de Salida')]/../../following-sibling::td";
                        IWebElement dateElement=null;
						
						int maxAttemptsdataelement = 3; // Número máximo de intentos
						int attemptsdataelement = 0; // Intentos realizados
						bool elementFounddataelement = false; // Variable para verificar si el elemento fue encontrado

						while (attemptsdataelement < maxAttemptsdataelement && !elementFounddataelement)
						{
							try
							{
								// Intentar encontrar el elemento
								dateElement = driver.FindElement(By.XPath(dateXpath));
								elementFounddataelement = true; // Si se encuentra el elemento, marcar como encontrado
								//writeLog.Log($"Fecha de salida encontrada en intento {attemptsdataelement + 1}.");
							}
							catch (NoSuchElementException)
							{
								attemptsdataelement++;
								writeLog.Log($"Fecha de Salida no encontrada para el dígito: {digit}. Intento {attemptsdataelement}/{maxAttemptsdataelement}.");

								if (attemptsdataelement < maxAttemptsdataelement)
								{
									// Espera antes de reintentar (por ejemplo, 5 segundos)
									Task.Delay(5000).Wait();
								}
								else
								{
									// Si se alcanza el número máximo de intentos, se decide si continuar o lanzar excepción
									writeLog.Log($"No se pudo encontrar la fecha de salida después de {maxAttemptsdataelement} intentos.");
									 // O lanzar excepción si es necesario
								}
							}
						}
                    if (dateElement == null) {
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
                            writeLog.Log($"Tabla de puertos no encontrada para el dígito: {digit}");
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
                            .Where(x => GlobalSettings.CodeAduanaList.Contains(x.value.Text))
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
                            writeLog.Log($"No se encontraron datos filtrados para el dígito: {digit}");
                            continue;
                        }
                        GlobalSettings.DigitsListFiltered.Add(digit.Key);

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
                                sheet.Cells[dataRow, dataColumn].Value = digit.Key;
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
                        bool success = false;
                        int retryCount = 0;
                        int maxRetries = 5;
                        int waitTime = 10000;
                        while (!success && retryCount < maxRetries)
                        {
                            try
                            {
								Task.Delay(6000).Wait();
								await _obtenerembarcador.Obtener_embarcador(driver, wait, stoppingToken);
                                success = true; 
                            }
                            catch (Exception ex)
                            {
                                retryCount++;
                                if (retryCount < maxRetries)
                                {
                                    writeLog.Log($"Error: {ex.Message}. Reintentando en {waitTime / 1000} segundos...");
                                    Task.Delay(waitTime).Wait(); 
                                }
                                else
                                {
                                    writeLog.Log("Número máximo de reintentos alcanzado. Operación fallida. Error al obtener datos del embarcador");
                                    throw; 
                                }
                            }
                        }
                        Task.Delay(6000).Wait();
                        try
                        {
							bool elementsFound = false;
							int maxAttemptsmenu = 5;
							int attemptmenu = 0;
							IList<IWebElement>? pageLinks = null;
                            Task.Delay(15000).Wait();
							while (!elementsFound && attemptmenu < maxAttemptsmenu)
							{
								try
								{
									attemptmenu++;
									pageLinks = driver.FindElements(By.XPath("//div[@id='tblLista_paginate']/span/a[contains(@class, 'paginate_button')]"));
									if (pageLinks.Count > 0)
									{
										elementsFound = true;
									}
									else
									{
										throw new NoSuchElementException();
									}
								}
								catch (Exception ex)
								{
                                    writeLog.Log($"ERROR AL BUSCAR EL MANIFIESTO EN SUANT Y BUSCAR EMBARCADOR :{ex}");
									if (attemptmenu < maxAttemptsmenu)
									{
										writeLog.Log($"REINTENTANDO ...... INTENTO NUMERO: {attemptmenu}");
										await Task.Delay(5000); // Espera 5 segundos antes del siguiente intento
									}
								}
							}

							if (!elementsFound)
							{
								writeLog.Log($"No se pudieron encontrar el elemento de navegacion de pagina después de {maxAttemptsmenu}  intentos.");
							}
							int totalPages = pageLinks.Count;
                            bool found = false;
                            for (int i = 1; i <= totalPages && !found; i++)
                            {
                                var manifiestoHeader = driver.FindElement(By.XPath("//table[@id='tblLista']//th[text()='Manifiesto de Carga']"));
                                int manifiestoColumnIndex = manifiestoHeader.GetAttribute("cellIndex") != null ? int.Parse(manifiestoHeader.GetAttribute("cellIndex")) : -1;

                                if (manifiestoColumnIndex == -1)
                                {
                                    writeLog.Log("Encabezado 'Manifiesto de Carga' no encontrado.");
                                    break;
                                }

                                var cells = driver.FindElements(By.XPath($"//table[@id='tblLista']//tr/td[{manifiestoColumnIndex + 1}]"));
                                foreach (var cell in cells)
                                {
                                    var link = cell.FindElement(By.TagName("a"));
                                    string linkText = link.Text;
                                    string pattern = @"- (\d+)$";
                                    System.Text.RegularExpressions.Match match = Regex.Match(linkText, pattern);
                                    if (match.Success)
                                    {
                                        string extractedNumber = match.Groups[1].Value; 
                                        if (extractedNumber == digit.Key)
                                        {
                                            js.ExecuteScript("arguments[0].click();", link);
                                            Task.Delay(6000).Wait(); 

											IWebElement inputBL = null;
											int maxAttempts = 5;
											int attempt = 0;
											bool elementFound = false;

											while (attempt < maxAttempts && !elementFound)
											{
												try
												{
													inputBL = driver.FindElement(By.XPath("//div[@id='tblLista_filter']//label//input[@class='form-control input-sm']"));
													elementFound = true; 
												}
												catch (NoSuchElementException)
												{
													attempt++;
													if (attempt < maxAttempts)
													{
														Task.Delay(6000).Wait();
													}
													else
													{
														writeLog.Log("Elemento input para buscar embarcador no encontrado o no cargado.");
													}
												}
											}
											if (!elementFound)
											{
												throw new NoSuchElementException("Elemento no encontrado después de varios intentos.");
											}
											Task.Delay(3000).Wait();
                                            for (int j = 0; j < filteredData["B/L Master/Directo"].Count; j++)
                                            {
                                                var blMasterDirecto = filteredData["B/L Master/Directo"][j];
                                                var detalleText = filteredData["Detalle"][j].Text;

                                                inputBL.Clear();
                                                js.ExecuteScript("arguments[0].click();", inputBL);
                                                inputBL.SendKeys(blMasterDirecto.Text);
                                                Task.Delay(2000).Wait();

                                                var detalleDirectoHeader = driver.FindElement(By.XPath("//table[@id='tblLista']//th[text() = 'Detalle Directo/Máster']"));
                                                var detalleHijoHeader = driver.FindElement(By.XPath("//table[@id='tblLista']//th[text() = 'Detalle Hijo']"));
                                                var embarcadorHeader = driver.FindElement(By.XPath("//table[@id='tblLista']//th[text() = 'Embarcador']"));
                                                int detalleDirectoColumnIndex = int.Parse(detalleDirectoHeader.GetAttribute("cellIndex"));
                                                int detalleHijoColumnIndex = int.Parse(detalleHijoHeader.GetAttribute("cellIndex"));
                                                int embarcadorColumnIndex = int.Parse(embarcadorHeader.GetAttribute("cellIndex"));

                                                
                                                IWebElement detalleCell = null;

                                                try
                                                {
                                                    detalleCell = driver.FindElement(By.XPath($"//table[@id='tblLista']//tr/td[{detalleDirectoColumnIndex + 1}][text() = '{detalleText}']"));
                                                }
                                                catch (NoSuchElementException)
                                                {
                                                    try
                                                    {
                                                        detalleCell = driver.FindElement(By.XPath($"//table[@id='tblLista']//tr/td[{detalleHijoColumnIndex + 1}][text() = '{detalleText}']"));
                                                    }
                                                    catch (NoSuchElementException)
                                                    {
                                                        detalleCell = null;
                                                    }
                                                }

                                                string embarcadorTexto;
                                                if (detalleCell != null)
                                                {
                                                    var fila = detalleCell.FindElement(By.XPath("./ancestor::tr"));

                                                    var embarcadorCell = fila.FindElement(By.XPath($"./td[{embarcadorColumnIndex + 1}]"));
                                                    embarcadorTexto = embarcadorCell.Text.Trim();
                                                }
                                                else
                                                {
                                                    embarcadorTexto = "Embarcador No Encontrado";
                                                }

                                                
                                                int filaExcel = -1;
                                                for (int rowIndex = currentRow; rowIndex <= sheet.Dimension.End.Row; rowIndex++)
                                                {
                                                    if (sheet.Cells[rowIndex, header.IndexOf("B/L Master/Directo") + 1].Text == blMasterDirecto.Text &&
                                                        sheet.Cells[rowIndex, header.IndexOf("Detalle") + 1].Text == detalleText)
                                                    {
                                                        filaExcel = rowIndex;
                                                    }
                                                }

                                                if (filaExcel != -1)
                                                {
                                                    
                                                    int embarcadorColumnIndexExcel = header.IndexOf("Embarcador") + 1;
                                                    sheet.Cells[filaExcel, embarcadorColumnIndexExcel].Value = embarcadorTexto;
                                                }
                                            }
                                            found = true; 
                                            break;
                                        }
                                    }
                                }

                                if (i < totalPages && !found)
                                {
                                    var nextButton = driver.FindElement(By.Id("tblLista_next"));
                                    if (nextButton.GetAttribute("class").Contains("disabled"))
                                    {
                                        break;
                                    }
                                    nextButton.Click();
                                    Task.Delay(2000).Wait(); 
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            writeLog.Log($"Error en la consulta de manifiesto: {ex.ToString()} en el digito extraido {digit}");
                            continue;
                        }

                        int columnIndexDUA = header.IndexOf("Nº DUA") + 1;
                        int columnIndexProducto = -1;
                        int columnIndexTamañoContenedor = -1;
                        int columnIndexNumeroContenedor = -1;

                        int duanoencontrado = 0;

                        for (int i = 0; i < filteredData["Conocimiento"].Count; i++)
                        {
                            var detailUrlConocimiento = filteredData["Conocimiento"][i].Href;

                            if (!string.IsNullOrEmpty(detailUrlConocimiento) && detailUrlConocimiento != "-")
                            {
								int maxAttempts = 3;
								int attempt = 0;
								bool successnav = false;

								while (attempt < maxAttempts && !successnav)
								{
									try
									{
										attempt++;

										driver.Navigate().GoToUrl(detailUrlConocimiento);
										successnav = true;
									}
									catch (WebDriverTimeoutException ex)
									{
										writeLog.Log($"Error en el intento {attempt} de navegar al Conocimiento url {filteredData["Conocimiento"][i].Text.ToString()} del manifiesto {digit}: {ex.Message}");

										Task.Delay(2000).Wait();
									}
									catch (Exception ex)
									{
										writeLog.Log($"Error inesperado en el intento {attempt}: {ex.Message}");
									}
								}

								if (!successnav)
								{
									if (i == filteredData["Conocimiento"].Count - 1)
									{
										writeLog.Log("No se pudo navegar a la URL del conocimiento después de varios intentos.");
										break;
									}
									continue;
								}
								
                                Task.Delay(3000).Wait();


                                string headerXpath = "//table[@border='1' and @width='100%']//tr[1]/td[center/font/b[contains(text(), 'Numero DUA')]]";
                                HashSet<string> uniqueDUAs = new HashSet<string>();
                                try
                                {
                                    IWebElement headerElement = driver.FindElement(By.XPath(headerXpath));
                                    int columnIndexD = int.Parse(headerElement.GetAttribute("cellIndex"));

                                    string duaXpath = $"//table[@border='1' and @width='100%']//tr[position()>1]/td[{columnIndexD + 1}]/font";
                                    var duaElements = driver.FindElements(By.XPath(duaXpath));

                                    foreach (var element in duaElements)
                                    {
                                        string numeroDUA = element.Text.Trim();
                                        if (!string.IsNullOrEmpty(numeroDUA) && numeroDUA != "-")
                                        {
                                            uniqueDUAs.Add(numeroDUA);
                                        }
                                    }
                                }
                                catch (Exception)
                                {
                                    duanoencontrado++;
                                }

                                List<string> listaDUAs = uniqueDUAs.ToList();

                                EscribirValoresEnExcel(sheet, currentRow + i, columnIndexDUA, listaDUAs);
                            }
                        }
                        if (duanoencontrado > 0) 
                        {
                            writeLog.Log($"Números DUA no encontrado: {duanoencontrado} para el manifiesto {digit}");
                        }
                        for (int col = 1; col <= sheet.Dimension.End.Column; col++)
                        {
                            if (sheet.Cells[1, col].Text.Trim() == "Producto")
                            {
                                columnIndexProducto = col;
                            }
                            if (sheet.Cells[1, col].Text.Trim() == "Tamaño de Contenedor")
                            {
                                columnIndexTamañoContenedor = col;
                            }
                            if (sheet.Cells[1, col].Text.Trim() == "N° de Contenedor")
                            {
                                columnIndexNumeroContenedor = col;
                                break;
                            }
                        }
                        if (columnIndexProducto == -1)
                        {
                            throw new Exception ("No se encontró la columna 'Producto' en la hoja de Excel.");
                        }
                        if (columnIndexTamañoContenedor == -1)
                        {
                            throw new Exception("No se encontró la columna 'Tamño Contenedor' en la hoja de Excel.");
                        }
                        if (columnIndexNumeroContenedor == -1)
                        {
                            throw new Exception("No se encontró la columna 'Numero de Contenedor' en la hoja de Excel.");
                        }
                        for (int i = 0; i < filteredData["Detalle"].Count; i++)
                        {
                            var detailUrlDetalle = filteredData["Detalle"][i].Href;
                            if (!string.IsNullOrEmpty(detailUrlDetalle) && detailUrlDetalle != "-")
                            {
								int maxAttempts = 3;
								int attempt = 0;
								bool successnav = false;

								while (attempt < maxAttempts && !successnav)
								{
									try
									{
										attempt++;

										driver.Navigate().GoToUrl(detailUrlDetalle);
										successnav = true;
									}
									catch (WebDriverTimeoutException ex)
									{
										writeLog.Log($"Error en el intento {attempt} de navegar al detalle url {filteredData["Detalle"][i].Text.ToString()} del manifiesto {digit}: {ex.Message}");

										Task.Delay(5000).Wait();
									}
									catch (Exception ex)
									{
										writeLog.Log($"Error inesperado en el intento {attempt}: {ex.Message}");
									}
								}

								if (!successnav)
								{
                                    if (i == filteredData["Detalle"].Count - 1) 
                                    {
										writeLog.Log("No se pudo navegar a la URL después de varios intentos.");
                                        break;
                                    }
									continue;
								}
								Task.Delay(3000).Wait();

                                string headerDescXpath = "//table[@border='1' and @width='100%']//tr[1]/td/font/b[contains(text(), 'Descripción de Mercadería')]";
								string headerMarcasXpath = "//table[@border='1' and @width='100%']//tr[1]/td/font/b[contains(text(), 'Marcas y Números')]";

								IWebElement detalleElement;
                                string Producto="-";
                                try
                                {
                                    Task.Delay(1500).Wait();
									//IWebElement headerDescElement = driver.FindElement(By.XPath(headerDescXpath));
									int maxAttemptsHeaderDesc = 3; // Número máximo de intentos para headerDescElement
									int currentAttemptHeaderDesc = 0;
									IWebElement headerDescElement = null;

									while (currentAttemptHeaderDesc < maxAttemptsHeaderDesc)
									{
										try
										{
											currentAttemptHeaderDesc++;
											headerDescElement = driver.FindElement(By.XPath(headerDescXpath));
											// Si se encuentra el elemento, salir del bucle
											break;
										}
										catch (NoSuchElementException)
										{
											writeLog.Log($"Intento {currentAttemptHeaderDesc} fallido para HeaderDesc. Reintentando...");
											Task.Delay(5000).Wait(); // Espera antes de intentar de nuevo (opcional)
										}
									}
									if (headerDescElement == null)
									{
										writeLog.Log($"No se pudo encontrar el elemento 'descripcion de mercaderia' después de {maxAttemptsHeaderDesc} intentos.");
                                        continue;
									}

									//IList<IWebElement> headerDescCells = driver.FindElements(By.XPath("//table[@border='1' and @width='100%']//tr[1]/td"));
									IList<IWebElement> headerDescCells = null;
									int maxAttemptshead = 5; // Número máximo de intentos
									int attempthead = 0;     // Contador de intentos

									while (headerDescCells == null && attempthead < maxAttemptshead)
									{
										try
										{
											attempthead++;
											headerDescCells = driver.FindElements(By.XPath("//table[@border='1' and @width='100%']//tr[1]/td"));

											// Verificamos si se encontraron elementos
											if (headerDescCells.Count == 0)
											{
												headerDescCells = null;
											}
										}
										catch (Exception ex)
										{
											// Puedes registrar el error para depuración si es necesario
											writeLog.Log($"Intento {attempthead}: {ex.Message}");
										}

										if (headerDescCells == null)
										{
											// Espera breve entre intentos
											Task.Delay(5000).Wait(); // 1 segundo
										}
									}

									// Validar el resultado
									if (headerDescCells == null)
									{
										writeLog.Log("No se encontro la tabla con el producto.");
									}
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
									if (Producto.Contains(". .") || Producto.Contains("...") || Producto.Contains(". . .") || Producto.Contains(".."))
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
										IList<IWebElement> marcasElements = driver.FindElements(By.XPath(marcasXpath));
										bool foundValidProduct = false;
										foreach (var element in marcasElements)
										{
											string text = element.Text.Trim();
											if (text != ". ." && text != "..." && text != ". . .")
											{
												Producto = text;
												foundValidProduct = true;
												break;
											}
										}
										if (!foundValidProduct)
										{
											throw new NoSuchElementException();
										}
									}
									catch (Exception)
									{
										Producto = "-";
										writeLog.Log("Detalle no encontrado, se guarda un guion '-'.");
									}
								}

                                
								int maxRetriesDESCRIP = 3; // Número máximo de intentos
								int attemptDescrip = 0;
								string tableXpath = "//table[@width='80%' and @border='']";
								IWebElement tableElement = null;
								int columnIndexTamaño = -1;
								int columnIndexNum = -1;
								while (attemptDescrip < maxRetriesDESCRIP && tableElement == null)
								{
									try
									{
										tableElement = driver.FindElement(By.XPath(tableXpath));
										IList<IWebElement> tableHeaders = tableElement.FindElements(By.XPath(".//tr[1]/td"));

										
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
									}
									catch (NoSuchElementException)
									{
										attemptDescrip++;
										if (attemptDescrip < maxRetriesDESCRIP)
										{
											Task.Delay(1500).Wait(); // Espera 1 segundo antes de reintentar
										}
										else
										{
											writeLog.Log($"Contenedor(es) no encontrado para el detalle {filteredData["Detalle"][i].Text.ToString()} del manifiesto {digit}");
											// Manejar el caso de error, por ejemplo, salir del bucle o continuar con la siguiente acción.
										}
									}
								}
                                string tamañoValue = "-";
                                string numeroValue = "-";
								
								
								try
								{
								    tamañoValue = tableElement.FindElement(By.XPath($".//tr[2]/td[{columnIndexTamaño + 1}]")).Text.Trim();
								}
								catch (Exception)
								{
                                    tamañoValue = "-"; 
								}

                                IList<IWebElement> tableRows = null;
								int numRows=0;
								try
								{
									if (tableElement != null)
									{
										tableRows = tableElement.FindElements(By.XPath(".//tr[position()>1]"));
                                        numRows = tableRows.Count;
									}
								}
								catch (Exception)
								{
									tableRows = new List<IWebElement>(); // Asignar una lista vacía para manejar mejor el caso
								}
								 
                                int totalColumns = 0;
								List<string> numeros = new List<string>();
								if (tableRows !=null && tableRows.Count>0)
								{
									for (int col = 1; col <= sheet.Dimension.End.Column; col++)
									{
										if (!string.IsNullOrEmpty(sheet.Cells[1, col].Text.Trim()))
										{
											totalColumns++;
										}
									}

									

									for (int j = 0; j < tableRows.Count; j++)
									{
										try
										{
											numeroValue = tableRows[j].FindElement(By.XPath($".//td[{columnIndexNum + 1}]")).Text.Trim();
											numeros.Add(numeroValue);
										}
										catch (Exception)
										{
											numeros.Add("-"); // Asignar "-" si ocurre una excepción
										}
									}
									// Código restante que sigue utilizando 'numeros'
								}
								else
								{
									// Maneja el caso si no hay filas o si tableRows es nulo.
									numeros.Add("-");
								}

								int filaExcel = -1;
                                for (int rowIndex = currentRow; rowIndex <= sheet.Dimension.End.Row; rowIndex++)
                                {
                                    if (sheet.Cells[rowIndex, header.IndexOf("Detalle") + 1].Text == filteredData["Detalle"][i].Text &&
                                        sheet.Cells[rowIndex, header.IndexOf("N° de Manifiesto") + 1].Text == digit.Key)
                                    {
                                        filaExcel = rowIndex;
                                        break;
                                    }
                                }

                                if (filaExcel != -1)
                                {
                                    
                                    var cellDetalle = sheet.Cells[filaExcel, columnIndexProducto];
                                    cellDetalle.Value = Producto;

                                    var cellTamaño = sheet.Cells[filaExcel, columnIndexTamañoContenedor];
                                    cellTamaño.Value = tamañoValue;

                                    sheet.Cells[filaExcel, columnIndexNumeroContenedor].Value = numeros[0];

                                    if (numRows > 1)
                                    {
                                        for (int j = 1; j < numRows; j++)
                                        {
                                            int targetRow = filaExcel + j;
                                            sheet.InsertRow(targetRow, 1);

                                            for (int k = 1; k <= totalColumns - 1; k++)
                                            {
                                                sheet.Cells[targetRow, k].Value = sheet.Cells[filaExcel, k].Value;
                                                if (sheet.Cells[filaExcel, k].Hyperlink != null)
                                                {
                                                    sheet.Cells[targetRow, k].Hyperlink = sheet.Cells[filaExcel, k].Hyperlink;
                                                }
                                            }
                                            sheet.Cells[targetRow, columnIndexNumeroContenedor].Value = numeros[j];
                                        }
                                    }
                                }

                            }
                        }
                        currentRow = sheet.Dimension.End.Row + 1;

                    }

                    if (sheet != null && sheet.Dimension != null)
                    {
                       
                        int productoColumnIndex = -1;
                        for (int col = 1; col <= sheet.Dimension.End.Column; col++)
                        {
                            if (sheet.Cells[1, col].Text.Equals("Producto", StringComparison.OrdinalIgnoreCase))
                            {
                                productoColumnIndex = col;
                                break;
                            }
                        }

                        if (productoColumnIndex == -1)
                        {
                            throw new Exception("No se encontró la columna 'Producto' en la hoja de Excel.");
                        }
                        int newColumnIndexproducstandar = productoColumnIndex + 1;
                        
                        sheet.InsertColumn(newColumnIndexproducstandar, 1); 
                        sheet.Cells[1, newColumnIndexproducstandar].Value = "Producto Estandarizado"; 

                        int lastRowWithContent = sheet.Dimension.End.Row;
                        for (int row = 2; row <= lastRowWithContent; row++) 
                        {
                            string productoValue = sheet.Cells[row, productoColumnIndex].Text.Trim(); 
                            string productoEstandarizadoValue = "-"; 
                           
                            foreach (var entry in GlobalSettings.GlobalProductDictionary)
                            {
                                foreach (var item in entry.Value)
                                {
                                    if (productoValue.Contains(item, StringComparison.OrdinalIgnoreCase))
                                    {
                                        productoEstandarizadoValue = entry.Key;
                                        
                                        break;
                                    }
                                    
                                }

                                if (productoEstandarizadoValue != "-") break; 
                            }

                            sheet.Cells[row, newColumnIndexproducstandar].Value = productoEstandarizadoValue; 
                        }

                        var range = sheet.Cells[1, 1, lastRowWithContent, sheet.Dimension.End.Column];
                        var table = sheet.Tables.Add(range, "Table1");
                        table.TableStyle = TableStyles.Medium5;
                        for (int col = 1; col <= sheet.Dimension.End.Column; col++)
                        {
                            sheet.Column(col).AutoFit();
                        }

                        package.Save();
                    }
                    else
                    {
                        writeLog.Log("La hoja de cálculo 'aduanet' está vacía o no tiene datos.");
                    }
                }

                writeLog.Log("Consulta realizada y datos exportados con éxito.");

            }
            catch (NoSuchElementException ex)
            {
                writeLog.Log($"El elemento no se encontró en la página: {ex.Message}");
            }
            #endregion

       }
        void EscribirValoresEnExcel(ExcelWorksheet sheet, int row, int columnIndexDUA, List<string> listaDUAs)
        {
            string encabezadoDUA = sheet.Cells[1, columnIndexDUA].Text.Trim();
            int numColumnasDUA = 0;
            for (int col = 1; col <= sheet.Dimension.End.Column; col++)
            {
                if (sheet.Cells[1, col].Text.Trim().Contains("Nº DUA"))
                {
                    numColumnasDUA++;
                }
            }

            if (listaDUAs.Count > numColumnasDUA)
            {
                for (int i = 0; i < listaDUAs.Count - numColumnasDUA; i++)
                {
                    int newColumnIndex = columnIndexDUA + numColumnasDUA + i;
                    sheet.InsertColumn(newColumnIndex, 1);
                    sheet.Cells[1, newColumnIndex].Value = encabezadoDUA + $"_{i + 2}";
                }
            }
            for (int i = 0; i < listaDUAs.Count; i++)
            {
                sheet.Cells[row, columnIndexDUA + i].Value = listaDUAs[i];
            }
        }
    }
}
