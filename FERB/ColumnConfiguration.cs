using System;
using OfficeOpenXml;

namespace FERB
{
    internal class ColumnConfiguration<TModel> : IColumnConfiguration<TModel>
    {
        private Action<ICellStyle> headerStyleApplier;
        private Action<ICellStyle> cellStyleApplier;

        public ColumnConfiguration()
        {
            IsVisible = true;
        }

        public string Format { get; set; }

        public bool IsVisible { get; set; }

        public void WithHeaderStyleApplied(Action<ICellStyle> applier)
        {
            this.headerStyleApplier = applier;
        }

        public void WithStyleApplied(Action<ICellStyle> applier)
        {
            this.cellStyleApplier = applier;
        }

        public void ApplyHeaderStyle(ExcelRange cell)
        {
            if (headerStyleApplier != null)
            {
                headerStyleApplier(new CellStyle(cell.Style));
            }
        }

        public void ApplyStyle(ExcelRange cell)
        {
            if (cellStyleApplier != null)
            {
                cellStyleApplier(new CellStyle(cell.Style));
            }
        }
    }
}
