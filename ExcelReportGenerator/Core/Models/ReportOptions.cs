using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReportGenerator.Core.Models
{
    public class ReportOptions
    {
        public string SheetName { get; set; } = "Report";
        public int? ConditionalColumnIndex { get; set; }
        public double? ConditionalThreshold { get; set; }
        public bool FreezeTopRow { get; set; } = true;
        public bool AutoFilter { get; set; } = true;
        public XLColor? HeaderBackgroundColor { get; set; } = XLColor.LightBlue;
    }
}
