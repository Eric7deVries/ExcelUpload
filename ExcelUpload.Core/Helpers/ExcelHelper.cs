using OfficeOpenXml;

namespace ExcelUpload.Core.Helpers;

public static class ExcelHelper
{
	public static Task<byte[]> ToExcelByteArray(List<dynamic> data, string sheetName)
	{
		if (data == null || data.Count == 0)
			throw new ArgumentException("No data available for export.");

		ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

		using var package = new ExcelPackage();
		var worksheet = package.Workbook.Worksheets.Add(sheetName);

		var columnHeaders = ((IDictionary<string, object>)data[0]).Keys.ToList();

		for (int col = 1; col <= columnHeaders.Count; col++)
		{
			worksheet.Cells[1, col].Value = columnHeaders[col - 1];
		}

		for (int row = 0; row < data.Count; row++)
		{
			var rowData = data[row];
			var rowDict = (IDictionary<string, object>)rowData;

			for (int col = 0; col < columnHeaders.Count; col++)
			{
				worksheet.Cells[row + 2, col + 1].Value = rowDict[columnHeaders[col]];
			}
		}

		return Task.FromResult(package.GetAsByteArray());
	}
}
