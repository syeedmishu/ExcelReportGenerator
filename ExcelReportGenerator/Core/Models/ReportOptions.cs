using ClosedXML.Excel;
using ExcelReportGenerator.Core.Enum;
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
        // Header
        public string Title { get; set; }
        public string[] SubtitleLines { get; set; }

        // Image
        public string ImagePath { get; set; }
        public int ImageWidth { get; set; } = 150;
        public int ImageHeight { get; set; } = 80;
        public ImagePosition ImagePosition { get; set; } = ImagePosition.TopLeft;
        public int CustomImageRow { get; set; } = 1;
        public int CustomImageColumn { get; set; } = 1;

        // Footer
        public string FooterText { get; set; }

        // Table Styling
        public int? ConditionalColumnIndex { get; set; }
        public double? ConditionalThreshold { get; set; }
        public bool FreezeTopRow { get; set; } = true;
        public bool AutoFilter { get; set; } = true;
        public XLColor? HeaderBackgroundColor { get; set; } = XLColor.LightBlue;
    }
}
