using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReportGenerator.Core.Models
{
    public class ReportSheetRequest
    {
        public string SheetName { get; set; }
        public DataTable Table { get; set; }
        public ReportOptions Options { get; set; } = new();
    }
}
