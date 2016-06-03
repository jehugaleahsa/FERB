using System;
using System.Collections;
using System.Collections.Specialized;
using System.IO;
using OfficeOpenXml;

namespace FERB
{
    public class WorkbookBuilder : IWorkbookBuilder
    {
        public const string ExcelMIMEType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
        public const string ExcelFileExtension = ".xlsx";

        private readonly OrderedDictionary builders;

        public WorkbookBuilder()
        {
            this.builders = new OrderedDictionary();
        }

        public IWorksheetBuilder AddWorksheet(string name)
        {
            if (String.IsNullOrWhiteSpace(name))
            {
                throw new ArgumentException("The worksheet name cannot be blank.", "name");
            }
            var builder = new WorksheetBuilder();
            builders[name] = builder;
            return builder;
        }

        public void SaveTo(Stream stream, Stream template = null)
        {
            using (ExcelPackage package = new ExcelPackage())
            {
                if (template != null)
                {
                    package.Load(template);
                }
                package.DoAdjustDrawings = false;
                var workbook = package.Workbook;
                foreach (DictionaryEntry entry in builders)
                {
                    var worksheetName = (string)entry.Key;
                    var builder = (WorksheetBuilder)entry.Value;
                    ExcelWorksheet worksheet = getWorksheet(workbook, worksheetName);
                    builder.Save(worksheet);
                }
                package.SaveAs(stream);
            }
        }

        private static ExcelWorksheet getWorksheet(ExcelWorkbook workbook, string worksheetName)
        {
            ExcelWorksheet worksheet = workbook.Worksheets[worksheetName];
            if (worksheet == null)
            {
                worksheet = workbook.Worksheets.Add(worksheetName);
            }
            return worksheet;
        }
    }
}
