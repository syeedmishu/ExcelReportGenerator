using ClosedXML.Excel;
using ExcelReportGenerator.Core.Enum;
using ExcelReportGenerator.Core.Interfaces;
using ExcelReportGenerator.Core.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

namespace ExcelReportGenerator.Infrastructure.ClosedXml
{
    public class ClosedXmlExcelGenerator : IExcelGenerator, IMultiSheetExcelGenerator
    {
        public byte[] GenerateFromDataTable(DataTable dt, ReportOptions options)
        {
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add(options.SheetName ?? "Report");

            RenderTableWithOptions(ws, dt, options);

            using var ms = new MemoryStream();
            wb.SaveAs(ms);
            return ms.ToArray();
        }
        public byte[] GenerateFromList<T>(List<T> list, ReportOptions options)
        {
            var dt = ToDataTable(list);
            return GenerateFromDataTable(dt, options);
        }

        public byte[] GenerateFromJson(string json, ReportOptions options)
        {
            var dt = JsonSerializer.Deserialize<DataTable>(json);
            return GenerateFromDataTable(dt, options);
        }

        public byte[] GenerateMultipleSheets(List<ReportSheetRequest> sheets)
        {
            using var workbook = new XLWorkbook();

            foreach (var sheet in sheets)
            {
                var ws = workbook.Worksheets.Add(sheet.SheetName);
                RenderTableWithOptions(ws, sheet.Table, sheet.Options);
            }

            using var stream = new MemoryStream();
            workbook.SaveAs(stream);
            return stream.ToArray();
        }

        private void RenderTableWithOptions(IXLWorksheet ws, DataTable dt, ReportOptions options)
        {
            int totalColumns = dt.Columns.Count;
            int currentRow = 1;

            // Add Image
            if (!string.IsNullOrWhiteSpace(options.ImagePath) && File.Exists(options.ImagePath))
            {
                using var imageStream = File.OpenRead(options.ImagePath);
                var picture = ws.AddPicture(imageStream)
                                 .WithSize(options.ImageWidth, options.ImageHeight);

                int col = 1, row = 1;

                if (options.ImagePosition == ImagePosition.Custom)
                {
                    row = options.CustomImageRow;
                    col = options.CustomImageColumn;
                }
                else
                {
                    row = 1;
                    switch (options.ImagePosition)
                    {
                        case ImagePosition.TopLeft: col = 1; break;
                        case ImagePosition.TopCenter: col = (totalColumns / 2); break;
                        case ImagePosition.TopRight: col = totalColumns - 1; break;
                        case ImagePosition.Watermark:
                            col = (totalColumns / 2);
                            row = dt.Rows.Count / 2;
                            break;
                    }
                }

                if (options.ImagePosition != ImagePosition.Watermark)
                    picture.MoveTo(ws.Cell(row, col));

                currentRow += 4;
            }

            // Watermark text
            if (options.ImagePosition == ImagePosition.Watermark)
            {
                var cell = ws.Cell(dt.Rows.Count / 2, totalColumns / 2);
                cell.Style.Font.FontSize = 40;
                cell.Style.Font.FontColor = XLColor.Silver;
                cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            }

            // Title
            if (!string.IsNullOrEmpty(options.Title))
            {
                var titleCell = ws.Cell(currentRow, 1);
                titleCell.Value = options.Title;
                titleCell.Style.Font.Bold = true;
                titleCell.Style.Font.FontSize = 16;
                titleCell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                ws.Range(currentRow, 1, currentRow, totalColumns).Merge();
                currentRow++;
            }

            // Subtitle
            if (options.SubtitleLines != null)
            {
                foreach (var line in options.SubtitleLines)
                {
                    var subCell = ws.Cell(currentRow, 1);
                    subCell.Value = line;
                    subCell.Style.Font.Italic = true;
                    subCell.Style.Font.FontSize = 11;
                    subCell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    ws.Range(currentRow, 1, currentRow, totalColumns).Merge();
                    currentRow++;
                }
            }

            // Insert table
            var table = ws.Cell(currentRow, 1).InsertTable(dt);
            var header = table.HeadersRow();
            header.Style.Font.Bold = true;

            if (options.HeaderBackgroundColor != null)
                header.Style.Fill.BackgroundColor = options.HeaderBackgroundColor;

            if (options.AutoFilter)
                table.ShowAutoFilter = true;

            ws.Columns().AdjustToContents();

            if (options.FreezeTopRow)
                ws.SheetView.FreezeRows(currentRow);

            // Conditional formatting
            if (options.ConditionalColumnIndex.HasValue && options.ConditionalThreshold.HasValue)
            {
                var range = ws.Range(currentRow + 1, 1, currentRow + dt.Rows.Count, totalColumns);
                foreach (var row in range.Rows())
                {
                    var cell = row.Cell(options.ConditionalColumnIndex.Value);
                    if (double.TryParse(cell.GetString(), out var val) && val > options.ConditionalThreshold.Value)
                    {
                        row.Style.Fill.BackgroundColor = XLColor.LightPink;
                    }
                }
            }

            // Footer
            if (!string.IsNullOrWhiteSpace(options.FooterText))
            {
                int footerRow = currentRow + dt.Rows.Count + 2;
                var footerCell = ws.Cell(footerRow, 1);
                footerCell.Value = options.FooterText;
                footerCell.Style.Font.Italic = true;
                footerCell.Style.Font.FontSize = 10;
                footerCell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                ws.Range(footerRow, 1, footerRow, totalColumns).Merge();
            }
        }



        private void StyleHeader(IXLTable table, ReportOptions options)
        {
            var header = table.HeadersRow();
            header.Style.Font.Bold = true;
            if (options.HeaderBackgroundColor != null)
                header.Style.Fill.BackgroundColor = options.HeaderBackgroundColor;
        }

        private void ApplyConditionalFormatting(IXLWorksheet ws, int startRow, int rowCount, int colCount, ReportOptions options)
        {
            if (options.ConditionalColumnIndex.HasValue && options.ConditionalThreshold.HasValue)
            {
                var range = ws.Range(startRow + 1, 1, startRow + rowCount, colCount);
                foreach (var row in range.Rows())
                {
                    var cell = row.Cell(options.ConditionalColumnIndex.Value);
                    if (double.TryParse(cell.GetString(), out var val) && val > options.ConditionalThreshold.Value)
                    {
                        row.Style.Fill.BackgroundColor = XLColor.LightPink;
                    }
                }
            }
        }

        private DataTable ToDataTable<T>(List<T> list)
        {
            var dt = new DataTable(typeof(T).Name);
            var props = typeof(T).GetProperties();
            foreach (var prop in props)
            {
                dt.Columns.Add(prop.Name, Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType);
            }
            foreach (var item in list)
            {
                var row = dt.NewRow();
                foreach (var prop in props)
                {
                    row[prop.Name] = prop.GetValue(item) ?? DBNull.Value;
                }
                dt.Rows.Add(row);
            }
            return dt;
        }


    }
}
