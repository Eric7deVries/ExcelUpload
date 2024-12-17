using ExcelUploadTests.Models;
using Microsoft.AspNetCore.Http;
using OfficeOpenXml;

namespace ExcelUpload.Test.Helpers;

public static class TestHelper
{
	public static IFormFile ConvertTestDataToExcel(TestData data)
	{
		ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

		if (data == null || data.People == null || !data.People.Any())
			throw new ArgumentException("No data available to convert.");

		// Generate an Excel package using EPPlus
		using var package = new ExcelPackage();
		var worksheet = package.Workbook.Worksheets.Add("People");

		// Add headers to the Excel file
		var headers = typeof(TestData.Person).GetProperties().Select(p => p.Name).ToList();
		for (int col = 0; col < headers.Count; col++)
		{
			worksheet.Cells[1, col + 1].Value = headers[col];
		}

		// Add data rows
		for (int row = 0; row < data.People.Count; row++)
		{
			var person = data.People[row];
			var values = headers.Select(header => typeof(TestData.Person).GetProperty(header)?.GetValue(person, null));
			int col = 0;
			foreach (var value in values)
			{
				worksheet.Cells[row + 2, col + 1].Value = value?.ToString() ?? string.Empty;
				col++;
			}
		}

		// Save the Excel package to a MemoryStream
		var stream = new MemoryStream();
		package.SaveAs(stream);
		stream.Position = 0; // Reset stream position for reading

		// Create the IFormFile using the MemoryStream
		var file = new FormFile(stream, 0, stream.Length, "file", "TestData.xlsx")
		{
			Headers = new HeaderDictionary(),
			ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
		};

		return file;
	}

	public static List<TestData.Person> ReadExcelToList(byte[] excelBytes)
	{
		ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

		var people = new List<TestData.Person>();

		using (var stream = new MemoryStream(excelBytes))
		using (var package = new ExcelPackage(stream))
		{
			var worksheet = package.Workbook.Worksheets[0];

			if (worksheet == null)
				throw new InvalidDataException("No worksheet found in the Excel file.");

			var headers = new Dictionary<int, string>();
			for (int col = 1; col <= worksheet.Dimension.Columns; col++)
			{
				headers[col] = worksheet.Cells[1, col].Text.Trim();
			}

			for (int row = 2; row <= worksheet.Dimension.Rows; row++)
			{
				var person = new TestData.Person();

				foreach (var header in headers)
				{
					var colIndex = header.Key;
					var columnName = header.Value;

					var cellValue = worksheet.Cells[row, colIndex].Text;

					// Map the cell values to the Person object properties
					switch (columnName)
					{
						case nameof(TestData.Person.ID):
							person.ID = long.TryParse(cellValue, out var id) ? id : 0;
							break;
						case nameof(TestData.Person.Name):
							person.Name = string.IsNullOrWhiteSpace(cellValue) ? null : cellValue;
							break;
						case nameof(TestData.Person.Age):
							person.Age = int.TryParse(cellValue, out var age) ? age : 0;
							break;
					}
				}

				people.Add(person);
			}
		}

		return people;
	}
}
