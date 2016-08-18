using System;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml;

namespace FERB
{
    internal class TableHeaderRow<TModel> : IWorksheetContent
    {
        private readonly int startingCell;
        private readonly IEnumerable<ColumnDefinition<TModel>> definitions;

        public TableHeaderRow(int startingCell, IEnumerable<ColumnDefinition<TModel>> definitions)
        {
            this.startingCell = startingCell;
            this.definitions = definitions;
        }

        public Action<ICellStyle> StyleApplier { get; set; }

        public int Save(ExcelWorksheet worksheet, int rowOffset)
        {
            var visibleColmns = definitions.Where(d => d.Configuration.IsVisible).ToArray();

            // Apply style to entire header
            if (StyleApplier != null)
            {
                int lastCell = startingCell + visibleColmns.Length;
                ExcelRange headerRow = worksheet.Cells[rowOffset, startingCell + 1, rowOffset, lastCell];
                StyleApplier(new CellStyle(headerRow.Style));
            }

            int currentCell = startingCell + 1;
            // Add header row to the worksheet and apply style
            foreach (ColumnDefinition<TModel> definition in visibleColmns)
            {
                var cell = worksheet.Cells[rowOffset, currentCell];
                cell.Value = definition.HeaderName;
                definition.Configuration.ApplyHeaderStyle(cell);
                ++currentCell;
            }

            return rowOffset;
        }
    }
}
