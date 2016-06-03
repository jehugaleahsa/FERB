using System;

namespace FERB
{
    public interface INestedTableBuilder<TModel>
    {
        INestedTableBuilder<TModel> WithTitle(string text, Action<ICellStyle> applier = null);

        INestedTableBuilder<TModel> WithHeaderStyleApplied(Action<ICellStyle> applier);

        INestedTableBuilder<TModel> WithStyleApplied(Action<ICellStyle> applier);

        INestedTableBuilder<TModel> WithColumn(string headerName, Func<TModel, object> propertyAccessor, Action<IColumnConfiguration<TModel>> configurator = null);
    }
}
