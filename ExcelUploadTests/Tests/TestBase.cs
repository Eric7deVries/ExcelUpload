using ExcelUpload.Core.Managers;
using ExcelUpload.Core.Providers;
using ExcelUpload.Services;
using Microsoft.Extensions.DependencyInjection;
using Moq;

namespace ExcelUpload.Test.Tests;

public class TestBase
{
	private readonly ServiceProvider _serviceProvider;
	private readonly Mock<IExcelInformationProvider> _mockExcelInformationProvider;
	private readonly Mock<IExcelManager> _mockExcelManager;

	public TestBase()
	{
		var services = new ServiceCollection();

		services.AddSingleton<IExcelInformationProvider, ExcelInformationProvider>();

		services.AddScoped<IExcelUploadService, ExcelUploadService>();
		services.AddScoped<IExcelManager, ExcelManager>();

		_serviceProvider = services.BuildServiceProvider();

		Services.SetServiceProvider(_serviceProvider);
	}

	public static class Services
	{
		private static ServiceProvider _serviceProvider;
		public static void SetServiceProvider(ServiceProvider serviceProvider)
		{
			_serviceProvider = serviceProvider;
		}

		public static IExcelUploadService ExcelUploadService => _serviceProvider.GetRequiredService<IExcelUploadService>();
	}
}
