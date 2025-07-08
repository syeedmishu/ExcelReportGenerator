using ClosedXML.Excel;
using ExcelReportGenerator.Core.Interfaces;
using ExcelReportGenerator.Core.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

namespace ExcelReportGenerator.Infrastructure.ClosedXml
{
    public class ClosedXmlExcelGenerator : IExcelGenerator
    {
        public byte[] GenerateFromDataTable(DataTable dt, ReportOptions options)
        {
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add(options.SheetName ?? "Report");

            var table = ws.Cell(1, 1).InsertTable(dt);

            StyleHeader(table, options);

            if (options.AutoFilter)
                table.ShowAutoFilter = true;

            ws.Columns().AdjustToContents();

            if (options.FreezeTopRow)
                ws.SheetView.FreezeRows(1);

            ApplyConditionalFormatting(ws, dt.Rows.Count, dt.Columns.Count, options);

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

        private void StyleHeader(IXLTable table, ReportOptions options)
        {
            var header = table.HeadersRow();
            header.Style.Font.Bold = true;
            if (options.HeaderBackgroundColor != null)
                header.Style.Fill.BackgroundColor = options.HeaderBackgroundColor;
        }

        private void ApplyConditionalFormatting(IXLWorksheet ws, int rowCount, int colCount, ReportOptions options)
        {
            if (options.ConditionalColumnIndex.HasValue && options.ConditionalThreshold.HasValue)
            {
                var dataRange = ws.Range(2, 1, rowCount + 1, colCount);
                foreach (var row in dataRange.Rows())
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
