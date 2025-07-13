using ExcelReportGenerator.Core.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReportGenerator.Core.Interfaces
{
    public interface IMultiSheetExcelGenerator
    {
        byte[] GenerateMultipleSheets(List<ReportSheetRequest> sheets);
    }
}
