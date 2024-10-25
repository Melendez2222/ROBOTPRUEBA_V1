using Microsoft.Extensions.Configuration;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using ROBOTPRUEBA_V1.CONFIG;
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
            var builder = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("jsconfig1.json", optional: false, reloadOnChange: true);

            _configuration = builder.Build();
        }
        TXTDigitosExtraidos TXTDigitosExtraidos = new TXTDigitosExtraidos();
        public async void Download() {
        string defaultDownloadDirectory = _configuration["FilePaths:DefaultDownloadDirectory"];
        string downloadDirectory = _configuration["FilePaths:DownloadDirectory"];
            var files = Directory.GetFiles(defaultDownloadDirectory, "*.xls");
            if (files.Length > 0)
            {
                var firstFile = files[0];
                var newFileName = Path.Combine(downloadDirectory, "Manifiesto_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".xls");
                GlobalSettings.ExcelFileManifestSunat=newFileName;
                File.Move(firstFile, newFileName);

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
               await TXTDigitosExtraidos.Digitos_Extraidos_TxT(downloadDirectory, extractedDigitsList);

            }
        }
    }
}
