using ExcelReportGenerator.Application.Interfaces;
using ExcelReportGenerator.Application.Services;
using ExcelReportGenerator.Core.Interfaces;
using ExcelReportGenerator.Infrastructure.ClosedXml;
using Microsoft.Extensions.DependencyInjection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReportGenerator.DependencyInjection
{
    public static class ServiceCollectionExtensions
    {
        public static IServiceCollection AddExcelReportGenerator(this IServiceCollection services)
        {
            // Register the core Excel generator
            services.AddScoped<IExcelGenerator, ClosedXmlExcelGenerator>();
            services.AddScoped<IMultiSheetExcelGenerator, ClosedXmlExcelGenerator>();

            // Register the report orchestration service
            services.AddScoped<IReportGenerationService, ReportGenerationService>();

            return services;
        }
    }
}
