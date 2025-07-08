using ExcelReportGenerator.Application.Interfaces;
using ExcelReportGenerator.Core.Interfaces;
using ExcelReportGenerator.Core.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReportGenerator.Application.Services
{
    public class ReportGenerationService : IReportGenerationService
    {
        private readonly IExcelGenerator _generator;

        public ReportGenerationService(IExcelGenerator generator)
        {
            _generator = generator;
        }

        public byte[] Generate(DataTable table, ReportOptions options)
        {
            return _generator.GenerateFromDataTable(table, options);
        }

        public byte[] Generate<T>(List<T> list, ReportOptions options)
        {
            return _generator.GenerateFromList(list, options);
        }

        public byte[] GenerateFromJson(string json, ReportOptions options)
        {
            return _generator.GenerateFromJson(json, options);
        }
    }
}
