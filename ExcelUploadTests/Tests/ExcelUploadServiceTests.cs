using ExcelUploadTests.Factories;
using Moq;
using ExcelUpload.Test.Helpers;
using Microsoft.AspNetCore.Http;
using ExcelUpload.Test.Tests;
using ExcelUploadTests.Models;

namespace ExcelUpload.Services.Tests;

[TestClass()]
public class ExcelUploadServiceTests : TestBase
{
	private readonly Mock<IExcelUploadService> _excelUploadService;

	public ExcelUploadServiceTests()
	{
		_excelUploadService = new Mock<IExcelUploadService>();
	}

	[TestMethod()]
	public async Task UploadExcelTestAsync()
	{
		var data = TestFactory.CreateTestData();
		var file = TestHelper.ConvertTestDataToExcel(data);
		_excelUploadService.Setup(service => service.UploadExcel(It.IsAny<IFormFile>()))
			.Returns(Task.CompletedTask);

		await _excelUploadService.Object.UploadExcel(file);

		_excelUploadService.Verify(service => service.UploadExcel(It.IsAny<IFormFile>()), Times.Once);
	}

	[TestMethod()]
	public async Task DownloadSampleExcel()
	{
		var fakeSampleExcel = new byte[] { 1, 2, 3, 4 }; // Simulate a valid byte array
		_excelUploadService.Setup(service => service.DownloadSampleExcel())
			.ReturnsAsync(fakeSampleExcel);

		var result = await _excelUploadService.Object.DownloadSampleExcel();

		Assert.AreEqual(fakeSampleExcel, result);
		_excelUploadService.Verify(service => service.DownloadSampleExcel(), Times.Once);
	}

	[TestMethod()]
	public async Task DownloadExcel()
	{
		var fakeExcelFile = new byte[] { 5, 6, 7, 8 }; // Simulate a valid byte array
		_excelUploadService.Setup(service => service.DownloadExcel())
			.ReturnsAsync(fakeExcelFile);

		var result = await _excelUploadService.Object.DownloadExcel();

		Assert.AreEqual(fakeExcelFile, result);
		_excelUploadService.Verify(service => service.DownloadExcel(), Times.Once);
	}

	[TestMethod()]
	public async Task TestHappyFlow()
	{
		var data1 = TestFactory.CreateTestData();

		var data2 = (TestData)data1.Clone();
		data2.TestSubject1.Age += 1;

		var file = TestHelper.ConvertTestDataToExcel(data1);

		await Services.ExcelUploadService.UploadExcel(file);

		var downloadedFile = await Services.ExcelUploadService.DownloadExcel();
		var returnedPeople = TestHelper.ReadExcelToList(downloadedFile);
		Assert.IsTrue(data1.People.All(returnedPeople.Contains));

		file = TestHelper.ConvertTestDataToExcel(data2);
		await Services.ExcelUploadService.UploadExcel(file);
		downloadedFile = await Services.ExcelUploadService.DownloadExcel();
		returnedPeople = TestHelper.ReadExcelToList(downloadedFile);

		Assert.IsTrue(returnedPeople.Single(x => x.ID == data1.TestSubject1.ID).Age == (data1.TestSubject1.Age + 1));
	}
}