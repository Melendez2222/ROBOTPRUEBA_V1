using NPOI.HSSF.UserModel;
using ROBOTPRUEBA_V1.FILES.LOG;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ROBOTPRUEBA_V1.FILES.Excel
{
    internal class Excelwriter
    {
        private readonly string _filePath;
        private readonly WriteLog _writeLog;
        private bool _isFirstWrite = true;

        public Excelwriter(string filePath, WriteLog writeLog)
        {
            _filePath = filePath;
            _writeLog = writeLog;
        }

        public void WriteData(string digit, string fechaSalida, List<string> header, List<List<string>> data)
        {
            try
            {
                using (var fileStream = new FileStream(_filePath, FileMode.OpenOrCreate, FileAccess.ReadWrite))
                {
                    var workbook = new HSSFWorkbook(fileStream);
                    var sheet = workbook.GetSheet("aduanet") ?? workbook.CreateSheet("aduanet");

                    int currentRow = sheet.LastRowNum + 1;

                    if (_isFirstWrite)
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
                        _isFirstWrite = false;
                    }

                    foreach (var rowData in data)
                    {
                        var row = sheet.CreateRow(currentRow++);
                        for (int j = 0; j < rowData.Count; j++)
                        {
                            var cell = row.CreateCell(j);
                            cell.SetCellValue(rowData[j]);
                        }

                        var cellManifest = row.CreateCell(rowData.Count);
                        cellManifest.SetCellValue(digit);

                        var cellFecha = row.CreateCell(rowData.Count + 1);
                        cellFecha.SetCellValue(fechaSalida);
                    }

                    using (var outputStream = new FileStream(_filePath, FileMode.Create, FileAccess.Write))
                    {
                        workbook.Write(outputStream);
                    }
                }

                _writeLog.Log("Datos exportados con éxito.");
            }
            catch (Exception ex)
            {
                _writeLog.Log($"Error al escribir en el archivo Excel: {ex.Message}");
            }
        }
    }
}
