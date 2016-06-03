using System;

namespace FERB
{
    public interface IWorksheetBuilder
    {
        ITableBuilder<TModel> AddTable<TModel>();

        ITitleBarBuilder AddTitleBar();

        IImageBuilder AddImage(string name);

        void AutoFitToContents(bool isAuto = true);
    }
}
