using System;
using System.Collections.Generic;
using OfficeOpenXml.Style;

namespace FERB
{
    public interface IOrderedTableBuilder<TModel> : IWorksheetContentBuilder
    {
        IOrderedTableBuilder<TModel> StartingAfter(IWorksheetContentBuilder builder);

        IOrderedTableBuilder<TModel> ThenBy(string headerName, bool isDescending = false);

        IOrderedTableBuilder<TModel> StartingAt(int rowNumber, int cellNumber);

        IOrderedTableBuilder<TModel> WithTitle(string title, Action<ICellStyle> applier = null);

        IOrderedTableBuilder<TModel> WithHeaderStyleApplied(Action<ICellStyle> applier);

        IOrderedTableBuilder<TModel> WithStyleApplied(Action<ICellStyle> applier);

        IOrderedTableBuilder<TModel> WithColumn(string headerName, Func<TModel, object> propertyAccessor, Action<IColumnConfiguration<TModel>> configurator = null);

        IOrderedTableBuilder<TModel> WithData(IEnumerable<TModel> models);

        IOrderedTableBuilder<TModel> WithNestedTable<TItem>(Func<TModel, IEnumerable<TItem>> propertyAccessor, Action<INestedTableBuilder<TItem>> configurator);
    }
}
