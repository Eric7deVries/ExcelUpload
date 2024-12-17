using ExcelUpload.Core.Providers;
using OfficeOpenXml;
using System.Dynamic;

namespace ExcelUpload.Core.Managers;

public interface IExcelManager
{
	public Task<List<dynamic>> ToList(ExcelPackage package);
}

public class ExcelManager : IExcelManager
{
	private readonly IExcelInformationProvider _excelInformationProvider;

	public ExcelManager(IExcelInformationProvider excelInformationProvider)
	{
		_excelInformationProvider = excelInformationProvider;
	}

	public Task<List<dynamic>> ToList(ExcelPackage package)
	{
		ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

		var worksheet = package.Workbook.Worksheets[0];
		int rowCount = worksheet.Dimension.Rows;
		int colCount = worksheet.Dimension.Columns;

		var columnHeaders = new List<string>();
		for (int col = 1; col <= colCount; col++)
		{
			var headerValue = worksheet.Cells[1, col].Text.Trim();
			columnHeaders.Add(headerValue);
		}

		var idHeaders = columnHeaders.Where(x => x.ToUpper().Equals("ID"));
		if (idHeaders.Count() != 1)
		{
			throw new ArgumentException("Couldnt find mandatory ID column.");
		}

		var existingData = _excelInformationProvider.Data.Sheet ?? [];
		
		for (int row = 2; row <= rowCount; row++)
		{
			dynamic rowData = new ExpandoObject();
			var rowDict = (IDictionary<string, object>)rowData;

			int idColumnIndex = columnHeaders.IndexOf("ID");
			if (idColumnIndex == -1)
			{
				continue; 
			}

			var idValue = worksheet.Cells[row, idColumnIndex + 1].Text.Trim();

			bool allFieldsEmpty = true;

			for (int col = 1; col <= colCount; col++)
			{
				var columnName = columnHeaders[col - 1]; 
				var cellValue = worksheet.Cells[row, col].Text.Trim();

				
				if (!string.IsNullOrEmpty(cellValue))
				{
					allFieldsEmpty = false;
				}
				rowDict[columnName] = string.IsNullOrEmpty(cellValue) ? null : cellValue;
			}

			if (allFieldsEmpty)
			{
				var existingRow = existingData.FirstOrDefault(r => r.ID == idValue);
				if (existingRow != null)
				{
					existingData.Remove(existingRow);
				}
			}
			else
			{
				var existingRow = existingData.FirstOrDefault(r => r.ID == idValue);
				if (existingRow != null)
				{
					var existingRowDict = (IDictionary<string, object>)existingRow;
					foreach (var prop in rowDict)
					{
						existingRowDict[prop.Key] = prop.Value;
					}
				}
				else
				{
					existingData.Add(rowData);
				}
			}
		}

		return Task.FromResult(existingData);
	}
}
