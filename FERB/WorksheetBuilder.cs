using System.Collections.Generic;
using OfficeOpenXml;

namespace FERB
{
    internal class WorksheetBuilder : IWorksheetBuilder
    {
        private readonly List<IWorksheetContent> builders;
        private bool isAuto;

        public WorksheetBuilder()
        {
            this.builders = new List<IWorksheetContent>();
            this.isAuto = true;
        }

        public ITableBuilder<TModel> AddTable<TModel>()
        {
            var builder = new TableBuilder<TModel>();
            builders.Add(builder);
            return builder;
        }

        public ITitleBarBuilder AddTitleBar()
        {
            var builder = new TitleBarBuilder();
            builders.Add(builder);
            return builder;
        }

        public IImageBuilder AddImage(string name)
        {
            var builder = new ImageBuilder(name);
            builders.Add(builder);
            return builder;
        }

        public void AutoFitToContents(bool isAuto = true)
        {
            this.isAuto = isAuto;
        }

        internal void Save(ExcelWorksheet worksheet)
        {
            Dictionary<IWorksheetContentBuilder, int> lastRowLookup = new Dictionary<IWorksheetContentBuilder, int>();
            foreach (IWorksheetContent content in builders)
            {
                IWorksheetContentBuilder builder = (IWorksheetContentBuilder)content;
                int rowOffset = 1;
                IWorksheetContentBuilder predecessor = builder.Predecessor;
                if (predecessor != null && lastRowLookup.ContainsKey(predecessor))
                {
                    rowOffset = lastRowLookup[predecessor];
                }
                int currentRow = content.Save(worksheet, rowOffset);
                lastRowLookup[builder] = currentRow;
            }
            if (isAuto)
            {
                worksheet.Cells.AutoFitColumns();
            }
        }
    }
}
