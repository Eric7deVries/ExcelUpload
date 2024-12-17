using ExcelUpload.Core.Helpers;
using ExcelUpload.Core.Managers;
using ExcelUpload.Core.Providers;
using OfficeOpenXml;

namespace ExcelUpload.Services;

public interface IExcelUploadService
{
	public Task UploadExcel(IFormFile file);
	public Task<byte[]> DownloadSampleExcel();
	public Task<byte[]> DownloadExcel();
}

public class ExcelUploadService : IExcelUploadService
{
	private readonly IExcelInformationProvider _excelInformationProvider;
	private readonly IExcelManager _excelManager;

	public ExcelUploadService(IExcelInformationProvider excelInformationProvider, IExcelManager excelManager)
	{
		_excelInformationProvider = excelInformationProvider;
		_excelManager = excelManager;
	}

	public async Task UploadExcel(IFormFile file)
	{
		using var stream = new MemoryStream();
		file.CopyTo(stream);

		using var package = new ExcelPackage(stream);
		var @object = await _excelManager.ToList(package);
		_excelInformationProvider.Data = (@object, package.Workbook.Worksheets[0].Name);
	}

	public async Task<byte[]> DownloadExcel()
	{
		if(_excelInformationProvider.Data.Sheet == null)
		{
			throw new Exception("No data found to download");
		}

		return await ExcelHelper.ToExcelByteArray(_excelInformationProvider.Data.Sheet, _excelInformationProvider.Data.SheetName);		
	}

	public async Task<byte[]> DownloadSampleExcel()
	{
		return await _excelInformationProvider.GetSampleExcelFile();
	}
}
