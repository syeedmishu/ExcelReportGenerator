using ExcelReportGenerator.Application.Interfaces;
using ExcelReportGenerator.Core.Interfaces;
using ExcelReportGenerator.Core.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReportGenerator.Application.Services
{
    public class ReportGenerationService : IReportGenerationService
    {
        private readonly IExcelGenerator _generator;
        private readonly IMultiSheetExcelGenerator _multiSheetGenerator;

        public ReportGenerationService(IExcelGenerator generator
            , IMultiSheetExcelGenerator multiSheetGenerator)
        {
            _generator = generator;
            _multiSheetGenerator = multiSheetGenerator;
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

        public byte[] GenerateMultipleSheets(List<ReportSheetRequest> sheets)
        {
            return _multiSheetGenerator.GenerateMultipleSheets(sheets);
        }

    }
}
