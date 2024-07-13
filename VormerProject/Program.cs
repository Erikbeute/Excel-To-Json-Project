using OfficeOpenXml;
using VormerProject.Services;
public class Program 
{
    public static void Main(string[] args) 
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;  //non commercial licentie voor ePPlus voor excel handling. 
        CreateHostBuilder(args).Build().Run(); 
    }

    public static IHostBuilder CreateHostBuilder(string[] args) => 
        Host.CreateDefaultBuilder(args) 
        .ConfigureWebHostDefaults(webBuilder => 
        {
            webBuilder.ConfigureServices(services => 
            {
                services.AddControllers(); 
                services.AddScoped<IExcelService, ExcelService>(); 
                services.AddSwaggerGen();
            });

            webBuilder.Configure(app => 
            {
                var env = app.ApplicationServices.GetRequiredService<IWebHostEnvironment>(); 

                if (env.IsDevelopment()) 
                {
                    app.UseDeveloperExceptionPage();
                    app.UseSwagger(); 
                    app.UseSwaggerUI(c => c.SwaggerEndpoint("/swagger/v1/swagger.json", "ExcelUploadApi v1")); 
                }

                app.UseRouting(); 
                app.UseAuthorization(); 

                app.UseEndpoints(endpoints => 
                {
                    endpoints.MapControllers(); 
                });
            });
        });
}

