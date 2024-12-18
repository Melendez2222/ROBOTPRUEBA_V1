using Microsoft.Extensions.Configuration;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium;
using ROBOTPRUEBA_V1.CONFIG;
using ROBOTPRUEBA_V1.FILES.LOG;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using OfficeOpenXml.Table;
using SeleniumExtras.WaitHelpers;
using System.Diagnostics;
using ROBOTPRUEBA_V1.FILES.Excel;
using NPOI.SS.UserModel;
using ROBOTPRUEBA_V1.Service;

namespace ROBOTPRUEBA_V1.ADUANET.PAITA
{
    internal class NavigateConsultaPaita
    {
        private readonly IConfiguration _configuration;
        private readonly Dictionary<string, string> _codeSelec;
        private readonly List<string> _codes;
        private readonly ConvertFormatExcel convertFormatExcel;
        private readonly Obtenerurldetalle obtenerurldetalle;
        public static bool columnsAdded = false;
        public NavigateConsultaPaita()
        {
            var builder = new ConfigurationBuilder()
                .SetBasePath(AppDomain.CurrentDomain.BaseDirectory)
                .AddJsonFile("appsettings.json", optional: true, reloadOnChange: true);;

            _configuration = builder.Build();
            _codes = _configuration.GetSection("Codes").Get<List<string>>();
            _codeSelec = _configuration.GetSection("CODE_SELEC").Get<Dictionary<string, string>>();
            convertFormatExcel = new ConvertFormatExcel();
            obtenerurldetalle = new Obtenerurldetalle();
        }
        public async void NavPaita(IWebDriver driver, WebDriverWait wait, CancellationToken stoppingToken)
        {
            driver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(120);
            string consultaManifiestoUrl = _configuration["Navigate:C_M_REGIONES"];
			string filePattern = "Reporte competencia semanal Paita semana*.xlsx";
			WriteLog writeLog = new WriteLog();
            string downloadDirectory = _configuration["FilePaths:ConvertFormatDirectoryPAITA"];
			DateTime today = DateTime.Today;
			var files = Directory.GetFiles(downloadDirectory, filePattern)
				.Where(f => File.GetCreationTime(f).Date == today)
				.OrderByDescending(f => File.GetCreationTime(f))
				.ToArray();
			if (files.Length == 0)
			{
				return;
			}
			if (GlobalSettings.ExtractedDigitsList.Count == 0)
            {
                writeLog.Log("No hay dígitos extraídos disponibles.");
                return;
            }
            try
            {
                foreach (var digit in GlobalSettings.ExtractedDigitsList)
                {
					int maxAttemptsnav = 3;
					int attemptnav = 0;
					bool successnav = false;

					while (attemptnav < maxAttemptsnav && !successnav)
					{
						try
						{
							attemptnav++;

							driver.Navigate().GoToUrl(consultaManifiestoUrl);
							successnav = true;
						}
						catch (WebDriverTimeoutException ex)
						{
							writeLog.Log($"Error en el intento {attemptnav} de navegar al aduanet para el manifiesto {digit}: {ex.Message}");

							Task.Delay(3000).Wait();
						}
						catch (Exception ex)
						{
							writeLog.Log($"Error inesperado en el intento {attemptnav}: {ex.Message}");
						}
					}
                    if (!successnav) {
						writeLog.Log($"Error en navegar al aduanet para el manifiesto {digit}");
						continue;
                    }
					
                    Task.Delay(3000).Wait();
                    IWebElement inputManifest = driver.FindElement(By.Name("CMc1_Numero"));
                    inputManifest.SendKeys(digit);
                    Task.Delay(1000).Wait();
                    var CODADUANASelect = driver.FindElement(By.Id("CG_cadu"));
                    var selectElementCA = new SelectElement(CODADUANASelect);
                    Task.Delay(1000).Wait();
                    if (GlobalSettings.CurrentCode != null && _codeSelec.TryGetValue(GlobalSettings.CurrentCode, out string aduanaValue))
                    {
                        selectElementCA.SelectByValue(aduanaValue);
                    }
                    else
                    {
                        throw new Exception($"No se encontró un valor para el código {GlobalSettings.CurrentCode}");
                    }
                    Task.Delay(1000).Wait();
                    var ViaSelect = driver.FindElement(By.Id("viat"));
                    var selectElementVia = new SelectElement(ViaSelect);
                    selectElementVia.SelectByValue("1");
                    Task.Delay(1000).Wait();
                    var consultButton = driver.FindElement(By.Name("btnConsultar"));
                    consultButton.Click();
                    Task.Delay(3000).Wait();
					// var dateElement = driver.FindElement(By.XPath("//td/b[contains(text(), 'Fecha de Zarpe')]/../following-sibling::td"));
					int maxAttemptsDateElement = 3; // Número máximo de intentos
					int attemptsDateElement = 0; // Contador de intentos
					IWebElement dateElement = null; // Inicializa la variable para el elemento

					// Bucle para intentar encontrar el elemento
					while (attemptsDateElement < maxAttemptsDateElement)
					{
						try
						{
							// Intenta encontrar el elemento
							dateElement = driver.FindElement(By.XPath("//td/b[contains(text(), 'Fecha de Zarpe')]/../following-sibling::td"));
							break; // Si se encuentra, salimos del bucle
						}
						catch (NoSuchElementException)
						{
							// Si no se encuentra, incrementa el contador de intentos
							attemptsDateElement++;
							if (attemptsDateElement < maxAttemptsDateElement)
							{
								Task.Delay(3000).Wait(); // Espera el tiempo especificado usando Task.Delay
							}
							else
							{
								writeLog.Log("No se pudo encontrar la fecha de zarpe después de varios intentos.");
								// Aquí puedes manejar el caso en que no se encontró el elemento
							}
						}
					}

					// Si se encontró el elemento, puedes trabajar con él
					if (dateElement == null)
					{
						writeLog.Log("El elemento 'Fecha de Zarpe' no fue encontrado o la pagina no cargó.");
                        continue;
					}
					string fechaSalida = dateElement.Text.Trim();
					int maxAttempts = 3;
					bool elementFound = false;

					for (int attempt = 1; attempt <= maxAttempts; attempt++)
					{
						try
						{
							var exportExcel = driver.FindElement(By.XPath("//*[contains(text(), 'Excel')]"));
							Task.Delay(1000).Wait();
							exportExcel.Click();
							Task.Delay(10000).Wait();

							elementFound = true;
							break; 
						}
						catch (Exception ex)
						{
							if (attempt == maxAttempts)
							{
								writeLog.Log($"El Manifiesto {digit} no contiene información a exportar después de {maxAttempts} intentos.");
							}
							else
							{
								writeLog.Log($"Intento de descarga Excel del manifiesto {digit}, intento N {attempt} fallido. Reintentando...");
								Task.Delay(1000).Wait(); 
							}
						}
					}

					if (!elementFound)
					{
						continue; 
					}
					Task.Delay(10000).Wait();
                    await convertFormatExcel.ConvertConsulManifest(fechaSalida, digit);
                    try
                    {
                        var paginasElement = driver.FindElement(By.XPath("//td[contains(text(), 'Páginas:')]"));

                        var links = paginasElement.FindElements(By.XPath(".//a"));

                        for (int i = 0; i <= links.Count; i++)
                        {
                            obtenerurldetalle.ObtenerInfDetalle(driver, digit);

                            if (i < links.Count)
                            {
                                IWebElement next = driver.FindElement(By.XPath("//td//a[contains(text(), 'Siguiente')]"));
                                next.Click();
                            }
                        }
                    }
                    catch (Exception)
                    {
                       obtenerurldetalle.ObtenerInfDetalle(driver, digit);

                    }
                }
            }
            catch (NoSuchElementException ex)
            {
                writeLog.Log($"Error al consultar Manifiesto de paita : {ex.Message}");
            }
            using (var package = new ExcelPackage(new FileInfo(GlobalSettings.ExcelFileManifestSunat)))
            {
                var workbook = package.Workbook;
                var sheet = workbook.Worksheets.FirstOrDefault(ws => ws.Name == "aduanet");

                if (sheet != null && sheet.Dimension != null)
                {
                    var mergedCells = sheet.MergedCells.ToList();

                    foreach (var mergedCell in mergedCells)
                    {
                        sheet.Cells[mergedCell].Merge = false;
                    }
                    for (int col = sheet.Dimension.End.Column; col >= 1; col--)
                    {
                        bool isEmptyColumn = true;
                        for (int row = 1; row <= sheet.Dimension.End.Row; row++)
                        {
                            if (!string.IsNullOrEmpty(sheet.Cells[row, col].Text))
                            {
                                isEmptyColumn = false;
                                break;
                            }
                        }
                        if (isEmptyColumn)
                        {
                            sheet.DeleteColumn(col);
                        }
                    }

                    for (int row = sheet.Dimension.End.Row; row >= 1; row--)
                    {
                        if (string.IsNullOrEmpty(sheet.Cells[row, 1].Text))
                        {
                            sheet.DeleteRow(row);
                        }
                    }

                    for (int row = sheet.Dimension.End.Row; row >= 2; row--)
                    {
                        if (!GlobalSettings.CodeAduanaList.Contains(sheet.Cells[row, 1].Text))
                        {
                            sheet.DeleteRow(row);
                        }
                    }
                    int detalleColumnIndex = -1;
                    int numManifColumnIndex = -1;

                    for (int col = 1; col <= sheet.Dimension.End.Column; col++)
                    {
                        if (sheet.Cells[1, col].Text.Equals("Detalle", StringComparison.OrdinalIgnoreCase))
                        {
                            detalleColumnIndex = col;
                        }
                        else if (sheet.Cells[1, col].Text.Equals("N° Manifiesto", StringComparison.OrdinalIgnoreCase))
                        {
                            numManifColumnIndex = col;
                        }
                    }

                   
                    int productoColumnIndex = -1;
                    int tamanoContenedorColumnIndex = -1;
                    int numeroContenedorColumnIndex = -1;

                    for (int col = 1; col <= sheet.Dimension.End.Column; col++)
                    {
                        if (sheet.Cells[1, col].Text.Equals("Producto", StringComparison.OrdinalIgnoreCase))
                        {
                            productoColumnIndex = col;
                        }
                        else if (sheet.Cells[1, col].Text.Equals("Tamaño Contenedor", StringComparison.OrdinalIgnoreCase))
                        {
                            tamanoContenedorColumnIndex = col;
                        }
                        else if (sheet.Cells[1, col].Text.Equals("N° de Contenedor", StringComparison.OrdinalIgnoreCase))
                        {
                            numeroContenedorColumnIndex = col;
                        }
                    }

                    int lastRowWithContent = sheet.Dimension.End.Row;
                    for (int row = sheet.Dimension.End.Row; row >= 1; row--)
                    {
                        bool isEmptyRow = true;
                        for (int col = 1; col <= sheet.Dimension.End.Column; col++)
                        {
                            if (!string.IsNullOrEmpty(sheet.Cells[row, col].Text))
                            {
                                isEmptyRow = false;
                                break;
                            }
                        }
                        if (!isEmptyRow)
                        {
                            lastRowWithContent = row;
                            break;
                        }
                    }
                    foreach (var entry in GlobalSettings.DetalleData)
                    {
                        string digit = entry.Key;
                        foreach (var detalleInfo in entry.Value)
                        {
                            int filaExcel = -1;
                            for (int rowIndex = 2; rowIndex <= lastRowWithContent; rowIndex++)
                            {
                                string digittable = sheet.Cells[rowIndex, numManifColumnIndex].Text.Replace(" ", "").Trim();
                                string detallenumtable = sheet.Cells[rowIndex, detalleColumnIndex].Text.Replace(" ", "").Trim();
                                if (digittable == digit.Trim() && detallenumtable == detalleInfo.DetalleNum.Trim())
                                {
                                    filaExcel = rowIndex;
                                    break;
                                }
                            }

                            if (filaExcel == -1)
                            {
                                continue;
                            }
                            sheet.Cells[filaExcel, productoColumnIndex].Value = detalleInfo.DescripcionText;
                            sheet.Cells[filaExcel, tamanoContenedorColumnIndex].Value = detalleInfo.TamanoText;
                            sheet.Cells[filaExcel, numeroContenedorColumnIndex].Value = detalleInfo.NumeroValues[0];
                            int targetRow = filaExcel;
                            if (detalleInfo.NumeroValues.Count > 1)
                            {
                                for (int j = 1; j < detalleInfo.NumeroValues.Count; j++)
                                {
                                    
                                    sheet.InsertRow(targetRow + j, 1);

                                    for (int k = 1; k <= sheet.Dimension.End.Column; k++)
                                    {
                                        var sourceCell = sheet.Cells[filaExcel, k]; 
                                        var targetCell = sheet.Cells[targetRow + j, k]; 
                                        targetCell.Value = sourceCell.Value; 
                                        targetCell.StyleID = sourceCell.StyleID;
                                    }
                                    sheet.Cells[targetRow + j, numeroContenedorColumnIndex].Value = detalleInfo.NumeroValues[j];

                                    filaExcel++;
                                    lastRowWithContent++;
                                }
                            }

                        }
                    }
                    for (int row = sheet.Dimension.End.Row; row >= 1; row--)
                    {
                        bool isEmptyRow = true;
                        for (int col = 1; col <= sheet.Dimension.End.Column; col++)
                        {
                            if (!string.IsNullOrEmpty(sheet.Cells[row, col].Text))
                            {
                                isEmptyRow = false;
                                break;
                            }
                        }
                        if (!isEmptyRow)
                        {
                            lastRowWithContent = row;
                            break;
                        }
                    }
                    int newColumnIndex = productoColumnIndex + 1; 
                    sheet.InsertColumn(newColumnIndex, 1); 

                    sheet.Cells[1, newColumnIndex].Value = "Producto Estandarizado";

                    for (int row = 2; row <= lastRowWithContent; row++)
                    {
                        string productoValue = sheet.Cells[row, productoColumnIndex].Text.Trim();
                        bool found = false;

                        foreach (var entry in GlobalSettings.GlobalProductDictionary)
                        {
                            foreach (var item in entry.Value)
                            {
                                if (productoValue.Contains(item, StringComparison.OrdinalIgnoreCase))
                                {
                                    sheet.Cells[row, newColumnIndex].Value = entry.Key;
                                    found = true;
                                    break;
                                }
                            }

                            if (found) break; 
                        }
                        if (!found)
                        {
                            sheet.Cells[row, newColumnIndex].Value = "-";
                        }
                    }
                    var range = sheet.Cells[1, 1, sheet.Dimension.End.Row, sheet.Dimension.End.Column];
                    var table = sheet.Tables.Add(range, "Table1");
                    table.TableStyle = TableStyles.Medium5;
                    for (int col = 1; col <= sheet.Dimension.End.Column; col++)
                    {
                        sheet.Column(col).AutoFit();
                    }

					//List<Dictionary<string, object>> datosParaEnviar = new List<Dictionary<string, object>>();

					//if (sheet.Dimension != null)
					//{
					//	int totalRows = sheet.Dimension.End.Row;
					//	int totalCols = sheet.Dimension.End.Column;

					//	// Leer los encabezados
					//	var headers = new List<string>();
					//	for (int col = 1; col <= totalCols; col++)
					//	{
					//		headers.Add(sheet.Cells[1, col].Text);
					//	}

					//	// Leer los datos de cada fila
					//	for (int row = 2; row <= totalRows; row++) // Comenzar desde 2 para omitir encabezados
					//	{
					//		var rowData = new Dictionary<string, object>();
					//		for (int col = 1; col <= totalCols; col++)
					//		{
					//			rowData[headers[col - 1]] = sheet.Cells[row, col].Text; // Almacenar el valor de la celda
					//		}
					//		datosParaEnviar.Add(rowData);
					//	}
					//}

					//// Enviar los datos a la API
					//ServiceReportes serviceReportes = new ServiceReportes();
					//await serviceReportes.EnviarDatosApiAsync(datosParaEnviar);


					package.Save();
                }
                else
                {
                    writeLog.Log("La hoja de cálculo 'aduanet' está vacía o no tiene datos.");
                }


            }
            writeLog.Log("SE TERMINO EL PROCESO DE PAITA");
        }
    }
}
