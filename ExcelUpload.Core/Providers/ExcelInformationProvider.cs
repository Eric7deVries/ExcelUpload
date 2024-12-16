using ExcelUpload.Core.Helpers;
using System.Dynamic;
using System.Reflection;

namespace ExcelUpload.Core.Providers;

public interface IExcelInformationProvider
{
	public Task<byte[]> GetSampleExcelFile();
	public (List<dynamic> Sheet, string SheetName) Data { get; set; }
}

public class ExcelInformationProvider : IExcelInformationProvider
{
	public (List<dynamic> Sheet, string SheetName) Data { get; set; }

	private byte[] SampleWorkbookByteArray;

	public async Task<byte[]> GetSampleExcelFile()
	{
		if (SampleWorkbookByteArray == null || SampleWorkbookByteArray.Length == 0)
		{
			await CreateSampleExcelFile();
		}

		return SampleWorkbookByteArray;
	}

	private class Person
	{
		public long ID { get; set; }
		public string Name { get; set; }
		public int Age { get; set; }
		public DateOnly DateOfBirth { get; set; }
	}

	private async Task CreateSampleExcelFile()
	{
		var people = new List<Person>() 
		{
			new() { ID = 1,	Name = "David", Age = 20, DateOfBirth = new DateOnly(2001, 11, 1) },
			new() {	ID = 2, Name = "Jan", Age = 34, DateOfBirth = new DateOnly(1990, 3, 23) },
			new() { ID = 3, Name = "Pieter", Age = 65, DateOfBirth = new DateOnly(1956, 6, 11) }
		};

		List<dynamic> sampleData = [];
		foreach (var person in people)
		{
			dynamic rowData = new ExpandoObject();
			var rowDict = (IDictionary<string, object>)rowData;

			var properties = person.GetType().GetProperties();
			foreach (var prop in properties)
			{
				rowDict[prop.Name] = prop.GetValue(person);
			}

			sampleData.Add(rowData);
		}

		SampleWorkbookByteArray = await ExcelHelper.ToExcelByteArray(sampleData, "People");
	}
}
