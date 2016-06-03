using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace FERB
{
    internal class TitleBarBuilder : ITitleBarBuilder, IWorksheetContent
    {
        private IWorksheetContentBuilder predecessor;
        private int rowNumber;
        private int cellNumber;
        private string title;
        private Action<ICellStyle> applier;
        private int cellCount;

        public TitleBarBuilder()
        {
            this.rowNumber = 1;
            this.cellNumber = 1;
            this.cellCount = 1;
        }

        IWorksheetContentBuilder IWorksheetContentBuilder.Predecessor
        {
            get { return predecessor; }
        }

        public ITitleBarBuilder StartingAfter(IWorksheetContentBuilder builder)
        {
            this.predecessor = builder;
            return this;
        }

        public ITitleBarBuilder StartingAt(int rowNumber, int cellNumber)
        {
            if (rowNumber < 1)
            {
                throw new ArgumentOutOfRangeException("rowNumber", "The row number cannot be less than one.");
            }
            if (cellNumber < 1)
            {
                throw new ArgumentOutOfRangeException("cellNumber", "The cell number cannot be less than one.");
            }
            this.rowNumber = rowNumber;
            this.cellNumber = cellNumber;
            return this;
        }

        public ITitleBarBuilder WithText(string title)
        {
            this.title = title;
            return this;
        }

        public ITitleBarBuilder WithStyleApplied(Action<ICellStyle> applier)
        {
            this.applier = applier;
            return this;
        }

        public ITitleBarBuilder WithColumnSpan(int count)
        {
            if (count < 1)
            {
                throw new ArgumentOutOfRangeException("count", "The count cannot be less than one.");
            }
            this.cellCount = count;
            return this;
        }

        int IWorksheetContent.Save(ExcelWorksheet worksheet, int rowOffset)
        {
            int currentRow = rowOffset + rowNumber;
            int lastCell = cellNumber + cellCount - 1;
            var range = worksheet.Cells[currentRow, cellNumber, currentRow, lastCell];
            range.Merge = true;
            range.Value = title;
            if (applier != null)
            {
                ICellStyle style = new CellStyle(range.Style);
                applier(style);
            }
            return currentRow;
        }
    }
}
