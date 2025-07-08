using ExcelReportGenerator.Core.Interfaces;
using ExcelReportGenerator.Application.Interfaces;
using ExcelReportGenerator.Application.Services;
using ExcelReportGenerator.Infrastructure.ClosedXml;
using Microsoft.Extensions.DependencyInjection;

namespace ExcelReportGenerator.Extensions
{
    public static class ServiceCollectionExtensions
    {
        public static IServiceCollection AddExcelReportGenerator(this IServiceCollection services)
        {
            // Register the core Excel generator
            services.AddScoped<IExcelGenerator, ClosedXmlExcelGenerator>();

            // Register the report orchestration service
            services.AddScoped<IReportGenerationService, ReportGenerationService>();

            return services;
        }
    }
}

