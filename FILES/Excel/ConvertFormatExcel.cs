using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Extensions.Configuration;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using ROBOTPRUEBA_V1.ADUANET.PAITA;
using ROBOTPRUEBA_V1.ADUANET.PISCO;
using ROBOTPRUEBA_V1.CONFIG;
using ROBOTPRUEBA_V1.FILES.LOG;
using ROBOTPRUEBA_V1.FILES.TXT;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ROBOTPRUEBA_V1.FILES.Excel
{
	internal class ConvertFormatExcel
	{
		private readonly IConfiguration _configuration;
		public ConvertFormatExcel()
		{
			var builder = new ConfigurationBuilder().SetBasePath(AppDomain.CurrentDomain.BaseDirectory).AddJsonFile("appsettings.json", optional: true, reloadOnChange: true); ;

			_configuration = builder.Build();
		}
		public async void ConvertXlsToXlsx(string xlsFilePath)
		{
			WriteLog writeLog = new WriteLog();
			ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

			string Converformatdirectory = _configuration["FilePaths:ConvertFormatDirectoryCALLAO"];
			string xlsxFilePath = Path.Combine(Converformatdirectory, "Reporte competencia semanal - Callao - semana " + GlobalSettings.NumSemana + ".xlsx");
			GlobalSettings.ExcelFileManifestSunat = xlsxFilePath;
			DateTime today = DateTime.Today;
			string[] files = Array.Empty<string>();

			try
			{
				files = Directory.GetFiles(Converformatdirectory, "Reporte competencia semanal - Callao - semana " + GlobalSettings.NumSemana + ".xlsx")
				.Where(f => File.GetCreationTime(f).Date == today)
				.OrderByDescending(f => File.GetCreationTime(f))
				.ToArray();
			}
			catch (Exception ex)
			{
				writeLog.Log($"Reporte competencia semanal no encontrado CALLAO.");
			}
			if (files.Length > 0)
			{
				var firstFile = files[0];
				try
				{
					if (File.Exists(firstFile))
					{
						File.Delete(firstFile);
					}
				}
				catch (Exception ex)
				{
					writeLog.Log($"Error al mover el archivo: {ex.Message}");
				}
			}
			using (var fileStream = new FileStream(xlsFilePath, FileMode.Open, FileAccess.Read))
			{
				var hssfWorkbook = new HSSFWorkbook(fileStream);
				var sheet = hssfWorkbook.GetSheetAt(0);

				using (var package = new ExcelPackage())
				{
					var xlSheet = package.Workbook.Worksheets.Add(sheet.SheetName);

					for (int rowIndex = 0; rowIndex <= sheet.LastRowNum; rowIndex++)
					{
						var hssfRow = sheet.GetRow(rowIndex);
						if (hssfRow != null)
						{
							var xlRow = xlSheet.Row(rowIndex + 1);
							xlRow.Height = hssfRow.Height / 20.0;

							for (int colIndex = 0; colIndex < hssfRow.LastCellNum; colIndex++)
							{
								var hssfCell = hssfRow.GetCell(colIndex);
								if (hssfCell != null)
								{
									var xlCell = xlSheet.Cells[rowIndex + 1, colIndex + 1];
									xlCell.Value = hssfCell.ToString();

									var hssfCellStyle = hssfCell.CellStyle;
									var hssfFont = hssfCellStyle.GetFont(hssfWorkbook);

									xlCell.Style.Font.Name = hssfFont.FontName;
									xlCell.Style.Font.Size = (float)hssfFont.FontHeightInPoints;
									xlCell.Style.Font.Bold = hssfFont.IsBold;
									xlCell.Style.Font.Italic = hssfFont.IsItalic;
									xlCell.Style.Font.UnderLine = hssfFont.Underline == FontUnderlineType.Single;

									var fontColor = hssfFont.Color;
									if (fontColor != HSSFColor.Automatic.Index)
									{
										var hssfColor = hssfWorkbook.GetCustomPalette().GetColor(fontColor);
										if (hssfColor != null)
										{
											var color = System.Drawing.Color.FromArgb(hssfColor.RGB[0], hssfColor.RGB[1], hssfColor.RGB[2]);
											xlCell.Style.Font.Color.SetColor(color);
										}
									}

									var hssfFillColor = hssfCellStyle.FillForegroundColorColor as HSSFColor;
									if (hssfFillColor != null)
									{
										var rgb = hssfFillColor.RGB;
										if (rgb != null)
										{
											var color = System.Drawing.Color.FromArgb(rgb[0], rgb[1], rgb[2]);
											xlCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
											xlCell.Style.Fill.BackgroundColor.SetColor(color);
										}
									}

									xlCell.Style.HorizontalAlignment = (ExcelHorizontalAlignment)hssfCellStyle.Alignment;
									xlCell.Style.VerticalAlignment = (ExcelVerticalAlignment)hssfCellStyle.VerticalAlignment;

									if (hssfCellStyle.BorderTop != BorderStyle.None)
									{
										xlCell.Style.Border.Top.Style = (ExcelBorderStyle)hssfCellStyle.BorderTop;
										xlCell.Style.Border.Top.Color.SetColor(System.Drawing.Color.FromArgb(hssfCellStyle.TopBorderColor));
									}

									if (hssfCellStyle.BorderBottom != BorderStyle.None)
									{
										xlCell.Style.Border.Bottom.Style = (ExcelBorderStyle)hssfCellStyle.BorderBottom;
										xlCell.Style.Border.Bottom.Color.SetColor(System.Drawing.Color.FromArgb(hssfCellStyle.BottomBorderColor));
									}

									if (hssfCellStyle.BorderLeft != BorderStyle.None)
									{
										xlCell.Style.Border.Left.Style = (ExcelBorderStyle)hssfCellStyle.BorderLeft;
										xlCell.Style.Border.Left.Color.SetColor(System.Drawing.Color.FromArgb(hssfCellStyle.LeftBorderColor));
									}

									if (hssfCellStyle.BorderRight != BorderStyle.None)
									{
										xlCell.Style.Border.Right.Style = (ExcelBorderStyle)hssfCellStyle.BorderRight;
										xlCell.Style.Border.Right.Color.SetColor(System.Drawing.Color.FromArgb(hssfCellStyle.RightBorderColor));
									}

									xlCell.Style.WrapText = true;


									bool isMerged = false;
									for (int i = 0; i < sheet.NumMergedRegions; i++)
									{
										var mergedRegion = sheet.GetMergedRegion(i);
										if (mergedRegion.IsInRange(hssfCell.RowIndex, hssfCell.ColumnIndex))
										{
											isMerged = true;
											break;
										}
									}

									if (!isMerged)
									{
										xlCell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
									}
								}
							}
						}
					}

					CellRangeAddress firstMergedRegion = null;
					CellRangeAddress lastMergedRegion = null;
					for (int i = 0; i < sheet.NumMergedRegions; i++)
					{
						var mergedRegion = sheet.GetMergedRegion(i);
						xlSheet.Cells[
							mergedRegion.FirstRow + 1, mergedRegion.FirstColumn + 1,
							mergedRegion.LastRow + 1, mergedRegion.LastColumn + 1
						].Merge = true;

						if (firstMergedRegion == null)
						{
							firstMergedRegion = mergedRegion;
						}
						lastMergedRegion = mergedRegion;
					}

					if (firstMergedRegion != null)
					{
						var firstMergedCell = xlSheet.Cells[firstMergedRegion.FirstRow + 1, firstMergedRegion.FirstColumn + 1];
						firstMergedCell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
					}

					if (lastMergedRegion != null)
					{
						var lastMergedCell = xlSheet.Cells[lastMergedRegion.FirstRow + 1, lastMergedRegion.FirstColumn + 1];
						lastMergedCell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
					}

					for (int i = 0; i < sheet.NumMergedRegions; i++)
					{
						var mergedRegion = sheet.GetMergedRegion(i);
						if (mergedRegion != firstMergedRegion && mergedRegion != lastMergedRegion)
						{
							var mergedCell = xlSheet.Cells[mergedRegion.FirstRow + 1, mergedRegion.FirstColumn + 1];
							mergedCell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
						}
					}

					for (int colIndex = 0; colIndex < sheet.GetRow(0).LastCellNum; colIndex++)
					{
						xlSheet.Column(colIndex + 1).Width = sheet.GetColumnWidth(colIndex) / 256.0;
					}

					package.SaveAs(new FileInfo(xlsxFilePath));
				}
			}
		}

		public async void ConvertXlsToXlsx_PAITA(string xlsFilePath)
		{
			WriteLog writeLog = new WriteLog();
			ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

			string Converformatdirectory = _configuration["FilePaths:ConvertFormatDirectoryPAITA"];
			string xlsxFilePath = Path.Combine(Converformatdirectory, "Reporte competencia semanal Paita semana " + GlobalSettings.NumSemana + ".xlsx");
			GlobalSettings.ExcelFileManifestSunat = xlsxFilePath;
			DateTime today = DateTime.Today;
			string[] files = Array.Empty<string>();

			try
			{
				files = Directory.GetFiles(Converformatdirectory, "Reporte competencia semanal - Callao - semana " + GlobalSettings.NumSemana + ".xlsx")
				.Where(f => File.GetCreationTime(f).Date == today)
				.OrderByDescending(f => File.GetCreationTime(f))
				.ToArray();
			}
			catch (Exception ex)
			{
				writeLog.Log($"Reporte competencia semanal no encontrado CALLAO.");
			}
			if (files.Length > 0)
			{
				var firstFile = files[0];
				try
				{
					if (File.Exists(firstFile))
					{
						File.Delete(firstFile);
					}
				}
				catch (Exception ex)
				{
					writeLog.Log($"Error al mover el archivo: {ex.Message}");
				}
			}
			using (var fileStream = new FileStream(xlsFilePath, FileMode.Open, FileAccess.Read))
			{
				var hssfWorkbook = new HSSFWorkbook(fileStream);
				var sheet = hssfWorkbook.GetSheetAt(0);

				using (var package = new ExcelPackage())
				{
					var xlSheet = package.Workbook.Worksheets.Add(sheet.SheetName);

					for (int rowIndex = 0; rowIndex <= sheet.LastRowNum; rowIndex++)
					{
						var hssfRow = sheet.GetRow(rowIndex);
						if (hssfRow != null)
						{
							var xlRow = xlSheet.Row(rowIndex + 1);
							xlRow.Height = hssfRow.Height / 20.0;

							for (int colIndex = 0; colIndex < hssfRow.LastCellNum; colIndex++)
							{
								var hssfCell = hssfRow.GetCell(colIndex);
								if (hssfCell != null)
								{
									var xlCell = xlSheet.Cells[rowIndex + 1, colIndex + 1];
									xlCell.Value = hssfCell.ToString();

									var hssfCellStyle = hssfCell.CellStyle;
									var hssfFont = hssfCellStyle.GetFont(hssfWorkbook);

									xlCell.Style.Font.Name = hssfFont.FontName;
									xlCell.Style.Font.Size = (float)hssfFont.FontHeightInPoints;
									xlCell.Style.Font.Bold = hssfFont.IsBold;
									xlCell.Style.Font.Italic = hssfFont.IsItalic;
									xlCell.Style.Font.UnderLine = hssfFont.Underline == FontUnderlineType.Single;

									var fontColor = hssfFont.Color;
									if (fontColor != HSSFColor.Automatic.Index)
									{
										var hssfColor = hssfWorkbook.GetCustomPalette().GetColor(fontColor);
										if (hssfColor != null)
										{
											var color = System.Drawing.Color.FromArgb(hssfColor.RGB[0], hssfColor.RGB[1], hssfColor.RGB[2]);
											xlCell.Style.Font.Color.SetColor(color);
										}
									}

									var hssfFillColor = hssfCellStyle.FillForegroundColorColor as HSSFColor;
									if (hssfFillColor != null)
									{
										var rgb = hssfFillColor.RGB;
										if (rgb != null)
										{
											var color = System.Drawing.Color.FromArgb(rgb[0], rgb[1], rgb[2]);
											xlCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
											xlCell.Style.Fill.BackgroundColor.SetColor(color);
										}
									}

									xlCell.Style.HorizontalAlignment = (ExcelHorizontalAlignment)hssfCellStyle.Alignment;
									xlCell.Style.VerticalAlignment = (ExcelVerticalAlignment)hssfCellStyle.VerticalAlignment;

									if (hssfCellStyle.BorderTop != BorderStyle.None)
									{
										xlCell.Style.Border.Top.Style = (ExcelBorderStyle)hssfCellStyle.BorderTop;
										xlCell.Style.Border.Top.Color.SetColor(System.Drawing.Color.FromArgb(hssfCellStyle.TopBorderColor));
									}

									if (hssfCellStyle.BorderBottom != BorderStyle.None)
									{
										xlCell.Style.Border.Bottom.Style = (ExcelBorderStyle)hssfCellStyle.BorderBottom;
										xlCell.Style.Border.Bottom.Color.SetColor(System.Drawing.Color.FromArgb(hssfCellStyle.BottomBorderColor));
									}

									if (hssfCellStyle.BorderLeft != BorderStyle.None)
									{
										xlCell.Style.Border.Left.Style = (ExcelBorderStyle)hssfCellStyle.BorderLeft;
										xlCell.Style.Border.Left.Color.SetColor(System.Drawing.Color.FromArgb(hssfCellStyle.LeftBorderColor));
									}

									if (hssfCellStyle.BorderRight != BorderStyle.None)
									{
										xlCell.Style.Border.Right.Style = (ExcelBorderStyle)hssfCellStyle.BorderRight;
										xlCell.Style.Border.Right.Color.SetColor(System.Drawing.Color.FromArgb(hssfCellStyle.RightBorderColor));
									}

									xlCell.Style.WrapText = true;


									bool isMerged = false;
									for (int i = 0; i < sheet.NumMergedRegions; i++)
									{
										var mergedRegion = sheet.GetMergedRegion(i);
										if (mergedRegion.IsInRange(hssfCell.RowIndex, hssfCell.ColumnIndex))
										{
											isMerged = true;
											break;
										}
									}

									if (!isMerged)
									{
										xlCell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
									}
								}
							}
						}
					}

					CellRangeAddress firstMergedRegion = null;
					CellRangeAddress lastMergedRegion = null;
					for (int i = 0; i < sheet.NumMergedRegions; i++)
					{
						var mergedRegion = sheet.GetMergedRegion(i);
						xlSheet.Cells[
							mergedRegion.FirstRow + 1, mergedRegion.FirstColumn + 1,
							mergedRegion.LastRow + 1, mergedRegion.LastColumn + 1
						].Merge = true;

						if (firstMergedRegion == null)
						{
							firstMergedRegion = mergedRegion;
						}
						lastMergedRegion = mergedRegion;
					}

					if (firstMergedRegion != null)
					{
						var firstMergedCell = xlSheet.Cells[firstMergedRegion.FirstRow + 1, firstMergedRegion.FirstColumn + 1];
						firstMergedCell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
					}

					if (lastMergedRegion != null)
					{
						var lastMergedCell = xlSheet.Cells[lastMergedRegion.FirstRow + 1, lastMergedRegion.FirstColumn + 1];
						lastMergedCell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
					}

					for (int i = 0; i < sheet.NumMergedRegions; i++)
					{
						var mergedRegion = sheet.GetMergedRegion(i);
						if (mergedRegion != firstMergedRegion && mergedRegion != lastMergedRegion)
						{
							var mergedCell = xlSheet.Cells[mergedRegion.FirstRow + 1, mergedRegion.FirstColumn + 1];
							mergedCell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
						}
					}

					for (int colIndex = 0; colIndex < sheet.GetRow(0).LastCellNum; colIndex++)
					{
						xlSheet.Column(colIndex + 1).Width = sheet.GetColumnWidth(colIndex) / 256.0;
					}

					package.SaveAs(new FileInfo(xlsxFilePath));
				}
			}
		}
		public async void ConvertXlsToXlsx_PISCO(string xlsFilePath)
		{
			WriteLog writeLog = new WriteLog();
			ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

			string Converformatdirectory = _configuration["FilePaths:ConvertFormatDirectoryPISCO"];
			string xlsxFilePath = Path.Combine(Converformatdirectory, "Reporte competencia semanal Pisco semana " + GlobalSettings.NumSemana + ".xlsx");
			GlobalSettings.ExcelFileManifestSunat = xlsxFilePath;
			DateTime today = DateTime.Today;
			string[] files = Array.Empty<string>();

			try
			{
				files = Directory.GetFiles(Converformatdirectory, "Reporte competencia semanal - Callao - semana " + GlobalSettings.NumSemana + ".xlsx")
				.Where(f => File.GetCreationTime(f).Date == today)
				.OrderByDescending(f => File.GetCreationTime(f))
				.ToArray();
			}
			catch (Exception ex)
			{
				writeLog.Log($"Reporte competencia semanal no encontrado CALLAO.");
			}
			if (files.Length > 0)
			{
				var firstFile = files[0];
				try
				{
					if (File.Exists(firstFile))
					{
						File.Delete(firstFile);
					}
				}
				catch (Exception ex)
				{
					writeLog.Log($"Error al mover el archivo: {ex.Message}");
				}
			}
			using (var fileStream = new FileStream(xlsFilePath, FileMode.Open, FileAccess.Read))
			{
				var hssfWorkbook = new HSSFWorkbook(fileStream);
				var sheet = hssfWorkbook.GetSheetAt(0);

				using (var package = new ExcelPackage())
				{
					var xlSheet = package.Workbook.Worksheets.Add(sheet.SheetName);

					for (int rowIndex = 0; rowIndex <= sheet.LastRowNum; rowIndex++)
					{
						var hssfRow = sheet.GetRow(rowIndex);
						if (hssfRow != null)
						{
							var xlRow = xlSheet.Row(rowIndex + 1);
							xlRow.Height = hssfRow.Height / 20.0;

							for (int colIndex = 0; colIndex < hssfRow.LastCellNum; colIndex++)
							{
								var hssfCell = hssfRow.GetCell(colIndex);
								if (hssfCell != null)
								{
									var xlCell = xlSheet.Cells[rowIndex + 1, colIndex + 1];
									xlCell.Value = hssfCell.ToString();

									var hssfCellStyle = hssfCell.CellStyle;
									var hssfFont = hssfCellStyle.GetFont(hssfWorkbook);

									xlCell.Style.Font.Name = hssfFont.FontName;
									xlCell.Style.Font.Size = (float)hssfFont.FontHeightInPoints;
									xlCell.Style.Font.Bold = hssfFont.IsBold;
									xlCell.Style.Font.Italic = hssfFont.IsItalic;
									xlCell.Style.Font.UnderLine = hssfFont.Underline == FontUnderlineType.Single;

									var fontColor = hssfFont.Color;
									if (fontColor != HSSFColor.Automatic.Index)
									{
										var hssfColor = hssfWorkbook.GetCustomPalette().GetColor(fontColor);
										if (hssfColor != null)
										{
											var color = System.Drawing.Color.FromArgb(hssfColor.RGB[0], hssfColor.RGB[1], hssfColor.RGB[2]);
											xlCell.Style.Font.Color.SetColor(color);
										}
									}

									var hssfFillColor = hssfCellStyle.FillForegroundColorColor as HSSFColor;
									if (hssfFillColor != null)
									{
										var rgb = hssfFillColor.RGB;
										if (rgb != null)
										{
											var color = System.Drawing.Color.FromArgb(rgb[0], rgb[1], rgb[2]);
											xlCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
											xlCell.Style.Fill.BackgroundColor.SetColor(color);
										}
									}

									xlCell.Style.HorizontalAlignment = (ExcelHorizontalAlignment)hssfCellStyle.Alignment;
									xlCell.Style.VerticalAlignment = (ExcelVerticalAlignment)hssfCellStyle.VerticalAlignment;

									if (hssfCellStyle.BorderTop != BorderStyle.None)
									{
										xlCell.Style.Border.Top.Style = (ExcelBorderStyle)hssfCellStyle.BorderTop;
										xlCell.Style.Border.Top.Color.SetColor(System.Drawing.Color.FromArgb(hssfCellStyle.TopBorderColor));
									}

									if (hssfCellStyle.BorderBottom != BorderStyle.None)
									{
										xlCell.Style.Border.Bottom.Style = (ExcelBorderStyle)hssfCellStyle.BorderBottom;
										xlCell.Style.Border.Bottom.Color.SetColor(System.Drawing.Color.FromArgb(hssfCellStyle.BottomBorderColor));
									}

									if (hssfCellStyle.BorderLeft != BorderStyle.None)
									{
										xlCell.Style.Border.Left.Style = (ExcelBorderStyle)hssfCellStyle.BorderLeft;
										xlCell.Style.Border.Left.Color.SetColor(System.Drawing.Color.FromArgb(hssfCellStyle.LeftBorderColor));
									}

									if (hssfCellStyle.BorderRight != BorderStyle.None)
									{
										xlCell.Style.Border.Right.Style = (ExcelBorderStyle)hssfCellStyle.BorderRight;
										xlCell.Style.Border.Right.Color.SetColor(System.Drawing.Color.FromArgb(hssfCellStyle.RightBorderColor));
									}

									xlCell.Style.WrapText = true;


									bool isMerged = false;
									for (int i = 0; i < sheet.NumMergedRegions; i++)
									{
										var mergedRegion = sheet.GetMergedRegion(i);
										if (mergedRegion.IsInRange(hssfCell.RowIndex, hssfCell.ColumnIndex))
										{
											isMerged = true;
											break;
										}
									}

									if (!isMerged)
									{
										xlCell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
									}
								}
							}
						}
					}

					CellRangeAddress firstMergedRegion = null;
					CellRangeAddress lastMergedRegion = null;
					for (int i = 0; i < sheet.NumMergedRegions; i++)
					{
						var mergedRegion = sheet.GetMergedRegion(i);
						xlSheet.Cells[
							mergedRegion.FirstRow + 1, mergedRegion.FirstColumn + 1,
							mergedRegion.LastRow + 1, mergedRegion.LastColumn + 1
						].Merge = true;

						if (firstMergedRegion == null)
						{
							firstMergedRegion = mergedRegion;
						}
						lastMergedRegion = mergedRegion;
					}

					if (firstMergedRegion != null)
					{
						var firstMergedCell = xlSheet.Cells[firstMergedRegion.FirstRow + 1, firstMergedRegion.FirstColumn + 1];
						firstMergedCell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
					}

					if (lastMergedRegion != null)
					{
						var lastMergedCell = xlSheet.Cells[lastMergedRegion.FirstRow + 1, lastMergedRegion.FirstColumn + 1];
						lastMergedCell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
					}

					for (int i = 0; i < sheet.NumMergedRegions; i++)
					{
						var mergedRegion = sheet.GetMergedRegion(i);
						if (mergedRegion != firstMergedRegion && mergedRegion != lastMergedRegion)
						{
							var mergedCell = xlSheet.Cells[mergedRegion.FirstRow + 1, mergedRegion.FirstColumn + 1];
							mergedCell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
						}
					}

					for (int colIndex = 0; colIndex < sheet.GetRow(0).LastCellNum; colIndex++)
					{
						xlSheet.Column(colIndex + 1).Width = sheet.GetColumnWidth(colIndex) / 256.0;
					}

					package.SaveAs(new FileInfo(xlsxFilePath));
				}
			}
		}

		public async Task ConvertConsulManifest(string date, string digit)
		{
			WriteLog writeLog = new WriteLog();
			ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

			string defaultDownloadDirectory = _configuration["FilePaths:DefaultDownloadDirectory"];
			string downloadDirectory = _configuration["FilePaths:DownloadDirectoryPAITA"];
			string filePattern = "ConsulManifSalida*.xls";

			DateTime today = DateTime.Today;

			var files = Directory.GetFiles(defaultDownloadDirectory, filePattern)
				.Where(f => File.GetCreationTime(f).Date == today)
				.OrderByDescending(f => File.GetCreationTime(f))
				.ToArray();
			if (files.Length > 0)
			{
				var firstFile = files[0];
				var newFileName = Path.Combine(downloadDirectory, "CONSULTMANIFSALIDA_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".xlsx");

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
				using (var fileStream = new FileStream(newFileName, FileMode.Open, FileAccess.Read))
				{
					var hssfWorkbook = new HSSFWorkbook(fileStream);
					var xssfWorkbook = new XSSFWorkbook();

					for (int i = 0; i < hssfWorkbook.NumberOfSheets; i++)
					{
						var hssfSheet = hssfWorkbook.GetSheetAt(i);
						var xssfSheet = xssfWorkbook.CreateSheet(hssfSheet.SheetName);

						foreach (var region in hssfSheet.MergedRegions)
						{
							xssfSheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(
								region.FirstRow, region.LastRow, region.FirstColumn, region.LastColumn));
						}

						for (int rowIndex = 0; rowIndex <= hssfSheet.LastRowNum; rowIndex++)
						{
							var sourceRow = hssfSheet.GetRow(rowIndex);
							var destinationRow = xssfSheet.CreateRow(rowIndex);
							if (sourceRow != null)
							{
								for (int colIndex = 0; colIndex < sourceRow.LastCellNum; colIndex++)
								{
									var sourceCell = sourceRow.GetCell(colIndex);
									var destinationCell = destinationRow.CreateCell(colIndex);
									if (sourceCell != null)
									{
										destinationCell.SetCellValue(sourceCell.ToString());



										if (sourceCell.Hyperlink != null)
										{
											var hyperlink = new XSSFHyperlink(sourceCell.Hyperlink.Type)
											{
												Address = sourceCell.Hyperlink.Address,
												Label = sourceCell.Hyperlink.Label
											};
											destinationCell.Hyperlink = hyperlink;
										}
									}
								}
							}
						}
					}

					using (var xlsxFileStream = new FileStream(newFileName, FileMode.Create, FileAccess.Write))
					{
						xssfWorkbook.Write(xlsxFileStream);
					}
				}
				using (var package = new ExcelPackage(new FileInfo(newFileName)))
				{
					var worksheet = package.Workbook.Worksheets[0];
					int puertoRowIndex = -1;
					int puertoColumnIndex = -1;

					for (int row = 1; row <= worksheet.Dimension.End.Row; row++)
					{
						for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
						{
							if (worksheet.Cells[row, col].Text.Equals("Puerto", StringComparison.OrdinalIgnoreCase))
							{
								puertoRowIndex = row;
								puertoColumnIndex = col;
								break;
							}
						}
						if (puertoRowIndex != -1) break;
					}


					if (puertoRowIndex != -1 && puertoColumnIndex != -1)
					{
						int lastRow = worksheet.Dimension.End.Row;

						var rangeToCopy = worksheet.Cells[puertoRowIndex, 2, lastRow, worksheet.Dimension.End.Column];

						using (var destinationPackage = new ExcelPackage(new FileInfo(GlobalSettings.ExcelFileManifestSunat)))
						{
							var destinationWorksheet = destinationPackage.Workbook.Worksheets.FirstOrDefault(ws => ws.Name == "aduanet") ?? destinationPackage.Workbook.Worksheets.Add("aduanet");
							int destinationStartRow = 1;
							if (destinationWorksheet.Dimension == null)
							{
								rangeToCopy.Copy(destinationWorksheet.Cells[1, 1]);
							}
							else
							{
								destinationStartRow = destinationWorksheet.Dimension.End.Row;
								if (destinationWorksheet.Cells[destinationStartRow, 1].Value != null)
								{
									destinationStartRow += 1;
								}
								var dataRangeToCopy = worksheet.Cells[puertoRowIndex + 1, 2, lastRow, worksheet.Dimension.End.Column];
								dataRangeToCopy.Copy(destinationWorksheet.Cells[destinationStartRow, 1]);
							}
							if (!NavigateConsultaPaita.columnsAdded)
							{
								int newColumnStart = destinationWorksheet.Dimension.End.Column + 1;
								destinationWorksheet.Cells[1, newColumnStart].Value = "N° Manifiesto";
								destinationWorksheet.Cells[1, newColumnStart + 1].Value = "Puerto de Origen";
								destinationWorksheet.Cells[1, newColumnStart + 2].Value = "Fecha";
								destinationWorksheet.Cells[1, newColumnStart + 3].Value = "Producto";
								destinationWorksheet.Cells[1, newColumnStart + 4].Value = "Tamaño Contenedor";
								destinationWorksheet.Cells[1, newColumnStart + 5].Value = "N° de Contenedor";
								NavigateConsultaPaita.columnsAdded = true;
							}
							int manifestColumnIndex = -1;
							int puertoOrigenColumnIndex = -1;
							int fechaColumnIndex = -1;
							for (int col = 1; col <= destinationWorksheet.Dimension.End.Column; col++)
							{
								if (destinationWorksheet.Cells[1, col].Text.Equals("N° Manifiesto", StringComparison.OrdinalIgnoreCase))
								{
									manifestColumnIndex = col;
								}
								else if (destinationWorksheet.Cells[1, col].Text.Equals("Puerto de Origen", StringComparison.OrdinalIgnoreCase))
								{
									puertoOrigenColumnIndex = col;
								}
								else if (destinationWorksheet.Cells[1, col].Text.Equals("Fecha", StringComparison.OrdinalIgnoreCase))
								{
									fechaColumnIndex = col;
								}
							}
							int lastDataRow = destinationWorksheet.Dimension.End.Row;
							for (int row = 2; row <= lastDataRow; row++)
							{
								if (destinationWorksheet.Cells[row, manifestColumnIndex].Value == null &&
									destinationWorksheet.Cells[row, puertoOrigenColumnIndex].Value == null &&
									destinationWorksheet.Cells[row, fechaColumnIndex].Value == null)
								{
									lastDataRow = row - 1;
									break;
								}
							}
							for (int row = lastDataRow + 1; row <= lastDataRow + (lastRow - puertoRowIndex); row++)
							{
								destinationWorksheet.Cells[row, manifestColumnIndex].Value = digit;
								destinationWorksheet.Cells[row, puertoOrigenColumnIndex].Value = "PAITA";
								destinationWorksheet.Cells[row, fechaColumnIndex].Value = date;
							}

							destinationPackage.Save();
						}
					}
					package.Save();
				}
			}
		}
		public async Task ConvertConsulManifestPisco(string date, string digit)
		{
			WriteLog writeLog = new WriteLog();
			ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

			string defaultDownloadDirectory = _configuration["FilePaths:DefaultDownloadDirectory"];
			string downloadDirectory = _configuration["FilePaths:DownloadDirectoryPISCO"];
			string filePattern = "ConsulManifSalida*.xls";

			DateTime today = DateTime.Today;

			var files = Directory.GetFiles(defaultDownloadDirectory, filePattern)
				.Where(f => File.GetCreationTime(f).Date == today)
				.OrderByDescending(f => File.GetCreationTime(f))
				.ToArray();
			if (files.Length > 0)
			{
				var firstFile = files[0];
				var newFileName = Path.Combine(downloadDirectory, "CONSULTMANIFSALIDA_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".xlsx");

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
				using (var fileStream = new FileStream(newFileName, FileMode.Open, FileAccess.Read))
				{
					var hssfWorkbook = new HSSFWorkbook(fileStream);
					var xssfWorkbook = new XSSFWorkbook();

					for (int i = 0; i < hssfWorkbook.NumberOfSheets; i++)
					{
						var hssfSheet = hssfWorkbook.GetSheetAt(i);
						var xssfSheet = xssfWorkbook.CreateSheet(hssfSheet.SheetName);
						foreach (var region in hssfSheet.MergedRegions)
						{
							xssfSheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(
								region.FirstRow, region.LastRow, region.FirstColumn, region.LastColumn));
						}

						for (int rowIndex = 0; rowIndex <= hssfSheet.LastRowNum; rowIndex++)
						{
							var sourceRow = hssfSheet.GetRow(rowIndex);
							var destinationRow = xssfSheet.CreateRow(rowIndex);
							if (sourceRow != null)
							{
								for (int colIndex = 0; colIndex < sourceRow.LastCellNum; colIndex++)
								{
									var sourceCell = sourceRow.GetCell(colIndex);
									var destinationCell = destinationRow.CreateCell(colIndex);
									if (sourceCell != null)
									{
										destinationCell.SetCellValue(sourceCell.ToString());



										if (sourceCell.Hyperlink != null)
										{
											var hyperlink = new XSSFHyperlink(sourceCell.Hyperlink.Type)
											{
												Address = sourceCell.Hyperlink.Address,
												Label = sourceCell.Hyperlink.Label
											};
											destinationCell.Hyperlink = hyperlink;
										}
									}
								}
							}
						}
					}

					using (var xlsxFileStream = new FileStream(newFileName, FileMode.Create, FileAccess.Write))
					{
						xssfWorkbook.Write(xlsxFileStream);
					}
				}
				using (var package = new ExcelPackage(new FileInfo(newFileName)))
				{
					var worksheet = package.Workbook.Worksheets[0];
					int puertoRowIndex = -1;
					int puertoColumnIndex = -1;

					for (int row = 1; row <= worksheet.Dimension.End.Row; row++)
					{
						for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
						{
							if (worksheet.Cells[row, col].Text.Equals("Puerto", StringComparison.OrdinalIgnoreCase))
							{
								puertoRowIndex = row;
								puertoColumnIndex = col;
								break;
							}
						}
						if (puertoRowIndex != -1) break;
					}


					if (puertoRowIndex != -1 && puertoColumnIndex != -1)
					{
						int lastRow = worksheet.Dimension.End.Row;

						var rangeToCopy = worksheet.Cells[puertoRowIndex, 2, lastRow, worksheet.Dimension.End.Column];

						using (var destinationPackage = new ExcelPackage(new FileInfo(GlobalSettings.ExcelFileManifestSunat)))
						{
							var destinationWorksheet = destinationPackage.Workbook.Worksheets.FirstOrDefault(ws => ws.Name == "aduanet") ?? destinationPackage.Workbook.Worksheets.Add("aduanet");
							int destinationStartRow = 1;
							if (destinationWorksheet.Dimension == null)
							{
								rangeToCopy.Copy(destinationWorksheet.Cells[1, 1]);
							}
							else
							{
								destinationStartRow = destinationWorksheet.Dimension.End.Row;
								if (destinationWorksheet.Cells[destinationStartRow, 1].Value != null)
								{
									destinationStartRow += 1;
								}
								var dataRangeToCopy = worksheet.Cells[puertoRowIndex + 1, 2, lastRow, worksheet.Dimension.End.Column];
								dataRangeToCopy.Copy(destinationWorksheet.Cells[destinationStartRow, 1]);
							}
							if (!NavigateConsultPisco.columnsAdded)
							{
								int newColumnStart = destinationWorksheet.Dimension.End.Column + 1;
								destinationWorksheet.Cells[1, newColumnStart].Value = "N° Manifiesto";
								destinationWorksheet.Cells[1, newColumnStart + 1].Value = "Puerto de Origen";
								destinationWorksheet.Cells[1, newColumnStart + 2].Value = "Fecha";
								destinationWorksheet.Cells[1, newColumnStart + 3].Value = "Producto";
								destinationWorksheet.Cells[1, newColumnStart + 4].Value = "Tamaño Contenedor";
								destinationWorksheet.Cells[1, newColumnStart + 5].Value = "N° de Contenedor";
								NavigateConsultPisco.columnsAdded = true;
							}
							int manifestColumnIndex = -1;
							int puertoOrigenColumnIndex = -1;
							int fechaColumnIndex = -1;
							for (int col = 1; col <= destinationWorksheet.Dimension.End.Column; col++)
							{
								if (destinationWorksheet.Cells[1, col].Text.Equals("N° Manifiesto", StringComparison.OrdinalIgnoreCase))
								{
									manifestColumnIndex = col;
								}
								else if (destinationWorksheet.Cells[1, col].Text.Equals("Puerto de Origen", StringComparison.OrdinalIgnoreCase))
								{
									puertoOrigenColumnIndex = col;
								}
								else if (destinationWorksheet.Cells[1, col].Text.Equals("Fecha", StringComparison.OrdinalIgnoreCase))
								{
									fechaColumnIndex = col;
								}
							}
							int lastDataRow = destinationWorksheet.Dimension.End.Row;
							for (int row = 2; row <= lastDataRow; row++)
							{
								if (destinationWorksheet.Cells[row, manifestColumnIndex].Value == null &&
									destinationWorksheet.Cells[row, puertoOrigenColumnIndex].Value == null &&
									destinationWorksheet.Cells[row, fechaColumnIndex].Value == null)
								{
									lastDataRow = row - 1;
									break;
								}
							}
							for (int row = lastDataRow + 1; row <= lastDataRow + (lastRow - puertoRowIndex); row++)
							{
								destinationWorksheet.Cells[row, manifestColumnIndex].Value = digit;
								destinationWorksheet.Cells[row, puertoOrigenColumnIndex].Value = "PISCO";
								destinationWorksheet.Cells[row, fechaColumnIndex].Value = date;
							}

							destinationPackage.Save();
						}
					}
					package.Save();
				}
			}
		}
	}
}
