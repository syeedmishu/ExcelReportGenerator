using ExcelReportGenerator.Core.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReportGenerator.Core.Interfaces
{
    public interface IExcelGenerator
    {
        byte[] GenerateFromDataTable(DataTable table, ReportOptions options);
        byte[] GenerateFromList<T>(List<T> list, ReportOptions options);
        byte[] GenerateFromJson(string json, ReportOptions options);

    }
}
