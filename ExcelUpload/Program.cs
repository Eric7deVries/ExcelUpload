
using ExcelUpload.Core.Managers;
using ExcelUpload.Core.Providers;
using ExcelUpload.Services;

namespace ExcelUpload
{
    public class Program
    {
        public static void Main(string[] args)
        {
            var builder = WebApplication.CreateBuilder(args);

			builder.Services.AddSingleton<IExcelInformationProvider, ExcelInformationProvider>();

            builder.Services.AddScoped<IExcelUploadService, ExcelUploadService>();
            builder.Services.AddScoped<IExcelManager, ExcelManager>();

			// Add services to the container.

			builder.Services.AddControllers();
            // Learn more about configuring OpenAPI at https://aka.ms/aspnet/openapi
            builder.Services.AddOpenApi();

			builder.Services.AddEndpointsApiExplorer();
			builder.Services.AddSwaggerGen();

			var app = builder.Build();

            // Configure the HTTP request pipeline.
            if (app.Environment.IsDevelopment())
            {
                app.MapOpenApi();
				app.UseSwagger();
				app.UseSwaggerUI(options =>
				{
					options.SwaggerEndpoint("/swagger/v1/swagger.json", "ExcelUploadPractice");
					options.RoutePrefix = string.Empty;
				});
			}

            app.UseAuthorization();


            app.MapControllers();

            app.Run();
        }
    }
}