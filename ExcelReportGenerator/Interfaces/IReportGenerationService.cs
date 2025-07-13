using ExcelReportGenerator.Core.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReportGenerator.Application.Interfaces
{
    public interface IReportGenerationService
    {
        byte[] Generate(DataTable table, ReportOptions options);
        byte[] Generate<T>(List<T> list, ReportOptions options);
        byte[] GenerateFromJson(string json, ReportOptions options);
        byte[] GenerateMultipleSheets(List<ReportSheetRequest> sheets);
    }
}
