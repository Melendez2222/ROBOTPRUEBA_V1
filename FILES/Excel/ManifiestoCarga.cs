using Microsoft.Extensions.Configuration;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using ROBOTPRUEBA_V1.CONFIG;
using ROBOTPRUEBA_V1.FILES.LOG;
using ROBOTPRUEBA_V1.FILES.TXT;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ROBOTPRUEBA_V1.FILES.Excel
{
    internal class ManifiestoCarga
    {
        private readonly IConfiguration _configuration;

        public ManifiestoCarga()
        {
            var builder = new ConfigurationBuilder().SetBasePath(AppDomain.CurrentDomain.BaseDirectory).AddJsonFile("appsettings.json", optional: true, reloadOnChange: true);;

            _configuration = builder.Build();
        }
        TXTDigitosExtraidos TXTDigitosExtraidos = new TXTDigitosExtraidos();
        ConvertFormatExcel ConvertFormatExcel = new ConvertFormatExcel();
        public async void Download()
        {
            WriteLog writeLog = new WriteLog();
            string defaultDownloadDirectory = _configuration["FilePaths:DefaultDownloadDirectory"];
            string downloadDirectory = _configuration["FilePaths:DownloadDirectoryCALLAO"];
            string filePattern = "repListadoManifiestoTrazabilidad*.xls";
			DateTime today = DateTime.Today;
			string[] files = Array.Empty<string>();

			try
            {
                files = Directory.GetFiles(defaultDownloadDirectory, filePattern)
                .Where(f => File.GetCreationTime(f).Date == today)
                .OrderByDescending(f => File.GetCreationTime(f))
                .ToArray();
			}
            catch (Exception ex) {
                writeLog.Log($"Reporte competencia semanal no encontrado CALLAO.");
            }
            if (files.Length > 0)
            {
                var firstFile = files[0];
                var newFileName = Path.Combine(downloadDirectory, "Reporte competencia semanal - Callao - semana " + GlobalSettings.NumSemana + ".xls");

                try
                {
                    if (File.Exists(newFileName))
                    {
                        File.Delete(newFileName);
                    }
                    File.Move(firstFile, newFileName);
                }
                catch (Exception ex)
                {
                    writeLog.Log($"Error al mover el archivo: {ex.Message}");
                }

                var extractedDigitsList = new List<string>();

                using (var fileStream = new FileStream(newFileName, FileMode.Open, FileAccess.ReadWrite))
                {
                    var workbook = new HSSFWorkbook(fileStream);
                    var sheet = workbook.GetSheetAt(0);


                    for (int rowIndex = 0; rowIndex <= sheet.LastRowNum; rowIndex++)
                    {
                        var row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
                        var newCell = row.CreateCell(2);


                        var previousCell = row.GetCell(1);
                        if (previousCell != null)
                        {
                            newCell.CellStyle = previousCell.CellStyle;
                        }
                    }


                    IRow headerRow = null;
                    int headerRowIndex = -1;
                    for (int rowIndex = 0; rowIndex <= sheet.LastRowNum; rowIndex++)
                    {
                        var row = sheet.GetRow(rowIndex);
                        if (row != null)
                        {
                            var cell = row.GetCell(1);
                            if (cell != null && cell.StringCellValue == "Manifiesto de Carga")
                            {
                                headerRow = row;
                                headerRowIndex = rowIndex;
                                break;
                            }
                        }
                    }

                    if (headerRow != null)
                    {

                        var buscarManifiestoCell = headerRow.CreateCell(2);
                        buscarManifiestoCell.SetCellValue("Buscar Manifiesto");
                        buscarManifiestoCell.CellStyle = headerRow.GetCell(1).CellStyle;

                        for (int rowIndex = headerRowIndex + 1; rowIndex <= sheet.LastRowNum; rowIndex++)
                        {
                            var row = sheet.GetRow(rowIndex);
                            if (row != null)
                            {
                                var cell = row.GetCell(1);
                                if (cell != null)
                                {
                                    var cellValue = cell.StringCellValue;
                                    int lastDashIndex = cellValue.LastIndexOf('-');
                                    if (lastDashIndex != -1 && lastDashIndex + 2 < cellValue.Length)
                                    {
                                        var extractedDigits = cellValue.Substring(lastDashIndex + 2).Trim();
                                        var newCell = row.CreateCell(2);
                                        newCell.CellStyle = cell.CellStyle;
                                        newCell.SetCellValue(extractedDigits);
                                        extractedDigitsList.Add(extractedDigits);

                                    }
                                }

                            }
                        }
                    }
                    using (var outputStream = new FileStream(newFileName, FileMode.Create, FileAccess.Write))
                    {
                        workbook.Write(outputStream);
                    }
                }
                ConvertFormatExcel.ConvertXlsToXlsx(newFileName);
                await TXTDigitosExtraidos.Digitos_Extraidos_TxT(downloadDirectory, extractedDigitsList);

            }
            else {
				writeLog.Log($"no se encontró archivos");
			}
        }
        public async void Download_PAITA()
        {
			WriteLog writeLog = new WriteLog();
			string defaultDownloadDirectory = _configuration["FilePaths:DefaultDownloadDirectory"];
            string downloadDirectory = _configuration["FilePaths:DownloadDirectoryPAITA"];
            string filePattern = "repListadoManifiestoTrazabilidad*.xls";

            DateTime today = DateTime.Today;
            var files = Directory.GetFiles(defaultDownloadDirectory, filePattern)
                .Where(f => File.GetCreationTime(f).Date == today)
                .OrderByDescending(f => File.GetCreationTime(f))
                .ToArray();
            if (files.Length > 0)
            {
                var firstFile = files[0];
                var newFileName = Path.Combine(downloadDirectory, "Reporte competencia semanal Paita semana" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".xls");

                try
                {
                    if (File.Exists(newFileName))
                    {
                        File.Delete(newFileName);
                    }
                    File.Move(firstFile, newFileName);
                }
                catch (Exception ex)
                {
                    writeLog.Log($"Error al mover el archivo: {ex.Message}");
                }

                var extractedDigitsList = new List<string>();

                using (var fileStream = new FileStream(newFileName, FileMode.Open, FileAccess.ReadWrite))
                {
                    var workbook = new HSSFWorkbook(fileStream);
                    var sheet = workbook.GetSheetAt(0);

                    for (int rowIndex = 0; rowIndex <= sheet.LastRowNum; rowIndex++)
                    {
                        var row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
                        var newCell = row.CreateCell(2);

                        var previousCell = row.GetCell(1);
                        if (previousCell != null)
                        {
                            newCell.CellStyle = previousCell.CellStyle;
                        }
                    }

                    IRow headerRow = null;
                    int headerRowIndex = -1;
                    for (int rowIndex = 0; rowIndex <= sheet.LastRowNum; rowIndex++)
                    {
                        var row = sheet.GetRow(rowIndex);
                        if (row != null)
                        {
                            var cell = row.GetCell(1);
                            if (cell != null && cell.StringCellValue == "Manifiesto de Carga")
                            {
                                headerRow = row;
                                headerRowIndex = rowIndex;
                                break;
                            }
                        }
                    }

                    if (headerRow != null)
                    {
                        var buscarManifiestoCell = headerRow.CreateCell(2);
                        buscarManifiestoCell.SetCellValue("Buscar Manifiesto");
                        buscarManifiestoCell.CellStyle = headerRow.GetCell(1).CellStyle;

                        for (int rowIndex = headerRowIndex + 1; rowIndex <= sheet.LastRowNum; rowIndex++)
                        {
                            var row = sheet.GetRow(rowIndex);
                            if (row != null)
                            {
                                var cell = row.GetCell(1);
                                if (cell != null)
                                {
                                    var cellValue = cell.StringCellValue;
                                    int lastDashIndex = cellValue.LastIndexOf('-');
                                    if (lastDashIndex != -1 && lastDashIndex + 2 < cellValue.Length)
                                    {
                                        var extractedDigits = cellValue.Substring(lastDashIndex + 2).Trim();
                                        var newCell = row.CreateCell(2);
                                        newCell.CellStyle = cell.CellStyle;
                                        newCell.SetCellValue(extractedDigits);
                                        extractedDigitsList.Add(extractedDigits);
                                    }
                                }
                            }
                        }
                    }

                    using (var outputStream = new FileStream(newFileName, FileMode.Create, FileAccess.Write))
                    {
                        workbook.Write(outputStream);
                    }
                }
                ConvertFormatExcel.ConvertXlsToXlsx_PAITA(newFileName);
                await TXTDigitosExtraidos.Digitos_Extraidos_TxT(downloadDirectory, extractedDigitsList);

            }
			else
			{
				writeLog.Log("No se encontraron archivos de reporte competencia semanal PISCO");
				return;
			}
		}
        public async void Download_PISCO()
        {
			WriteLog writeLog = new WriteLog();

			string defaultDownloadDirectory = _configuration["FilePaths:DefaultDownloadDirectory"];
            string downloadDirectory = _configuration["FilePaths:DownloadDirectoryPISCO"];
            string filePattern = "repListadoManifiestoTrazabilidad*.xls";

            DateTime today = DateTime.Today;
            var files = Directory.GetFiles(defaultDownloadDirectory, filePattern)
                .Where(f => File.GetCreationTime(f).Date == today)
                .OrderByDescending(f => File.GetCreationTime(f))
                .ToArray();
            if (files.Length > 0)
            {
                var firstFile = files[0];
                var newFileName = Path.Combine(downloadDirectory, "Reporte competencia semanal Pisco semana " + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".xls");


                try
                {
                    if (File.Exists(newFileName))
                    {
                        File.Delete(newFileName);
                    }
                    File.Move(firstFile, newFileName);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error al mover el archivo: {ex.Message}");
                }

                var extractedDigitsList = new List<string>();

                using (var fileStream = new FileStream(newFileName, FileMode.Open, FileAccess.ReadWrite))
                {
                    var workbook = new HSSFWorkbook(fileStream);
                    var sheet = workbook.GetSheetAt(0);

                    for (int rowIndex = 0; rowIndex <= sheet.LastRowNum; rowIndex++)
                    {
                        var row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
                        var newCell = row.CreateCell(2);

                        var previousCell = row.GetCell(1);
                        if (previousCell != null)
                        {
                            newCell.CellStyle = previousCell.CellStyle;
                        }
                    }

                    IRow headerRow = null;
                    int headerRowIndex = -1;
                    for (int rowIndex = 0; rowIndex <= sheet.LastRowNum; rowIndex++)
                    {
                        var row = sheet.GetRow(rowIndex);
                        if (row != null)
                        {
                            var cell = row.GetCell(1);
                            if (cell != null && cell.StringCellValue == "Manifiesto de Carga")
                            {
                                headerRow = row;
                                headerRowIndex = rowIndex;
                                break;
                            }
                        }
                    }

                    if (headerRow != null)
                    {
                        var buscarManifiestoCell = headerRow.CreateCell(2);
                        buscarManifiestoCell.SetCellValue("Buscar Manifiesto");
                        buscarManifiestoCell.CellStyle = headerRow.GetCell(1).CellStyle;

                        for (int rowIndex = headerRowIndex + 1; rowIndex <= sheet.LastRowNum; rowIndex++)
                        {
                            var row = sheet.GetRow(rowIndex);
                            if (row != null)
                            {
                                var cell = row.GetCell(1);
                                if (cell != null)
                                {
                                    var cellValue = cell.StringCellValue;
                                    int lastDashIndex = cellValue.LastIndexOf('-');
                                    if (lastDashIndex != -1 && lastDashIndex + 2 < cellValue.Length)
                                    {
                                        var extractedDigits = cellValue.Substring(lastDashIndex + 2).Trim();
                                        var newCell = row.CreateCell(2);
                                        newCell.CellStyle = cell.CellStyle;
                                        newCell.SetCellValue(extractedDigits);
                                        extractedDigitsList.Add(extractedDigits);
                                    }
                                }
                            }
                        }
                    }

                    using (var outputStream = new FileStream(newFileName, FileMode.Create, FileAccess.Write))
                    {
                        workbook.Write(outputStream);
                    }
                }
                ConvertFormatExcel.ConvertXlsToXlsx_PISCO(newFileName);
                await TXTDigitosExtraidos.Digitos_Extraidos_TxT(downloadDirectory, extractedDigitsList);

            }
            else {
                writeLog.Log("No se encontraron archivos de reporte competencia semanal PISCO");
                return;
            }
        }
    }
}
