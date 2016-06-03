using System;

namespace FERB
{
    public interface IColumnConfiguration<TModel>
    {
        string Format { get; set; }

        bool IsVisible { get; set; }

        void WithHeaderStyleApplied(Action<ICellStyle> applier);

        void WithStyleApplied(Action<ICellStyle> applier);
    }
}
