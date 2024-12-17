using ExcelUpload.Services;
using Microsoft.AspNetCore.Mvc;
using Swashbuckle.AspNetCore.Annotations;

namespace ExcelUpload.Controllers;

[ApiController]
[Route("[controller]")]
public class ExcelUploadController : ControllerBase
{
	private readonly IExcelUploadService _excelUploadService;

	public ExcelUploadController(IExcelUploadService excelUploadService)
	{
		_excelUploadService = excelUploadService;
	}

	/// <summary>
	/// Upload an Excel file to modify the objects we store.
	/// </summary>
	/// <remarks>
	/// Here an explaination how you can do CRUD functionaly on these objects.
	/// 
	/// Create: Add a new record to the Excel and upload. Duplicates will be ignored.
	/// 
	/// Read: See DownloadExcel.
	/// 
	/// Update: Change the record but keep the mandetory ID the same. This will update the record.
	/// 
	/// Delete: Send in the ID, but keep all fields empty will delete the record on our side.
	/// 
	/// Want to extend the object with an extra field?
	/// Add the field and make sure to also do this with existing items.
	/// 
	/// Note: ID is mandatory! (must be type of long)
	/// 
	/// For an example layout, call DownloadSampleExcel.
	/// </remarks>
	[HttpPost]
	[Route("[Action]")]
	public async Task<IActionResult> UploadExcel(IFormFile file)
	{
		//Create new thread so main thread is not occupied.
		await Task.Run(() => _excelUploadService.UploadExcel(file));
		return Ok();
	}

	///<summary>
	/// Downloads the current objects we store in Excel format.
	/// </summary>
	[HttpPost]
	[Route("[Action]")]
	public async Task<ActionResult<IFormFile>> DownloadExcel()
	{
		var sampleFile = await Task.Run(_excelUploadService.DownloadExcel);

		return File(sampleFile, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Excel.xlsx");
	}

	///<summary>
	/// Downloads sample objects in excel.
	///</summary>
	/// 
	/// <remarks>
	/// These objects are based on a Person object with the following attributes:
	/// ID
	/// Age
	/// Name
	/// DateOfBirth
	/// </remarks>
	[HttpPost]
	[Route("[Action]")]
	public async Task<ActionResult<IFormFile>> DownloadSampleExcel() 
	{
		var sampleFile = await Task.Run(_excelUploadService.DownloadSampleExcel);

		return File(sampleFile, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "SampleExcel.xlsx");
	}
}
