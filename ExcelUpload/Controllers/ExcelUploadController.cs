using ExcelUpload.Services;
using Microsoft.AspNetCore.Mvc;

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

	[HttpPost]
	[Route("[Action]")]
	public async Task<IActionResult> UploadExcel(IFormFile file)
	{
		//Create new thread so main thread is not occupied.
		await Task.Run(() => _excelUploadService.UploadExcel(file));
		return Ok();
	}

	[HttpPost]
	[Route("[Action]")]
	public async Task<ActionResult<IFormFile>> DownloadExcel()
	{
		var sampleFile = await Task.Run(_excelUploadService.DownloadExcel);

		return File(sampleFile, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Excel.xlsx");
	}

	[HttpPost]
	[Route("[Action]")]
	public async Task<ActionResult<IFormFile>> DownloadSampleExcel() 
	{
		var sampleFile = await Task.Run(_excelUploadService.DownloadSampleExcel);

		return File(sampleFile, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "SampleExcel.xlsx");
	}
}
