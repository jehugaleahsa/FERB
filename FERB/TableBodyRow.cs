using System;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml;

namespace FERB
{
    internal class TableBodyRow<TModel> : IWorksheetContent
    {
        private readonly TModel model;
        private readonly int startingCell;
        private readonly IEnumerable<ColumnDefinition<TModel>> columnDefinitions;

        public TableBodyRow(TModel model, int startingCell, IEnumerable<ColumnDefinition<TModel>> columnDefinitions)
        {
            this.model = model;
            this.startingCell = startingCell;
            this.columnDefinitions = columnDefinitions;
        }

        public Action<ICellStyle> StyleApplier { get; set; }

        public int Save(ExcelWorksheet worksheet, int rowOffset)
        {
            var visibleColumns = columnDefinitions.Where(d => d.Configuration.IsVisible).ToArray();
            applyRowStyle(worksheet, rowOffset);

            int currentCell = startingCell + 1;
            foreach (var definition in visibleColumns)
            {
                var cell = worksheet.Cells[rowOffset, currentCell];
                saveColumn(definition, cell);
                ++currentCell;
            }
            return rowOffset;
        }

        private void saveColumn(ColumnDefinition<TModel> definition, ExcelRange cell)
        {
            ColumnConfiguration<TModel> configuration = definition.Configuration;

            // Apply cell-specific styles
            if (configuration != null)
            {
                configuration.ApplyStyle(cell);
            }

            // Set value
            var value = definition.Accessor(model);
            if (configuration == null || configuration.Format == null || value == null)
            {
                cell.Value = value;
            }
            else
            {
                cell.Value = String.Format("{0:" + configuration.Format + "}", value);
            }
        }

        private void applyRowStyle(ExcelWorksheet worksheet, int currentRow)
        {
            if (StyleApplier == null)
            {
                return;
            }
            int visibleColumnCount = columnDefinitions.Where(d => d.Configuration.IsVisible).Count();
            int lastCell = startingCell + visibleColumnCount;  // Subtract since the end cell is inclusive
            ExcelRange range = worksheet.Cells[currentRow, startingCell + 1, currentRow, lastCell];
            StyleApplier(new CellStyle(range.Style));
        }
    }
}
