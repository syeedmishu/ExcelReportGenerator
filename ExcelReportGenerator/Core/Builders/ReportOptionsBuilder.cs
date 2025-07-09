using ClosedXML.Excel;
using ExcelReportGenerator.Core.Enum;
using ExcelReportGenerator.Core.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReportGenerator.Core.Builders
{
    public class ReportOptionsBuilder
    {
        private readonly ReportOptions _options = new();

        public static ReportOptionsBuilder Create() => new();

        public ReportOptionsBuilder WithSheetName(string sheetName)
        {
            _options.SheetName = sheetName;
            return this;
        }

        public ReportOptionsBuilder WithTitle(string title)
        {
            _options.Title = title;
            return this;
        }

        public ReportOptionsBuilder WithSubtitle(params string[] lines)
        {
            _options.SubtitleLines = lines;
            return this;
        }

        public ReportOptionsBuilder WithImage(string imagePath, ImagePosition position = ImagePosition.TopLeft, int width = 150, int height = 80)
        {
            _options.ImagePath = imagePath;
            _options.ImagePosition = position;
            _options.ImageWidth = width;
            _options.ImageHeight = height;
            return this;
        }

        public ReportOptionsBuilder WithCustomImagePosition(int row, int column)
        {
            _options.ImagePosition = ImagePosition.Custom;
            _options.CustomImageRow = row;
            _options.CustomImageColumn = column;
            return this;
        }

        public ReportOptionsBuilder WithWatermarkText(string text)
        {
            _options.ImagePosition = ImagePosition.Watermark;
            return this;
        }

        public ReportOptionsBuilder WithFooter(string footerText)
        {
            _options.FooterText = footerText;
            return this;
        }

        public ReportOptionsBuilder WithConditionalHighlight(int columnIndex, double threshold)
        {
            _options.ConditionalColumnIndex = columnIndex;
            _options.ConditionalThreshold = threshold;
            return this;
        }

        public ReportOptionsBuilder EnableFreezeTopRow()
        {
            _options.FreezeTopRow = true;
            return this;
        }

        public ReportOptionsBuilder DisableFreezeTopRow()
        {
            _options.FreezeTopRow = false;
            return this;
        }

        public ReportOptionsBuilder EnableAutoFilter()
        {
            _options.AutoFilter = true;
            return this;
        }

        public ReportOptionsBuilder DisableAutoFilter()
        {
            _options.AutoFilter = false;
            return this;
        }

        public ReportOptionsBuilder WithHeaderBackground(XLColor color)
        {
            _options.HeaderBackgroundColor = color;
            return this;
        }

        public ReportOptions Build() => _options;
    }
}