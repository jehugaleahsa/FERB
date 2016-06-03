using System;
using System.Collections.Generic;
using OfficeOpenXml.Style;

namespace FERB
{
    public interface ITableBuilder<TModel> : IWorksheetContentBuilder
    {
        ITableBuilder<TModel> StartingAfter(IWorksheetContentBuilder builder);

        ITableBuilder<TModel> StartingAt(int rowNumber, int cellNumber);

        ITableBuilder<TModel> WithTitle(string text, Action<ICellStyle> applier = null);

        ITableBuilder<TModel> WithHeaderStyleApplied(Action<ICellStyle> applier);

        ITableBuilder<TModel> WithStyleApplied(Action<ICellStyle> applier);

        ITableBuilder<TModel> WithColumn(string headerName, Func<TModel, object> propertyAccessor, Action<IColumnConfiguration<TModel>> configurator = null);

        IOrderedTableBuilder<TModel> OrderedBy(string headerName, bool isDescending = false);

        ITableBuilder<TModel> WithData(IEnumerable<TModel> models);

        ITableBuilder<TModel> WithNestedTable<TItem>(Func<TModel, IEnumerable<TItem>> propertyAccessor, Action<INestedTableBuilder<TItem>> configurator);
    }
}
