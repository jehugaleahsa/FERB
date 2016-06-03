using System;
using OfficeOpenXml.Style;

namespace FERB
{
    public interface ITitleBarBuilder : IWorksheetContentBuilder
    {
        ITitleBarBuilder StartingAfter(IWorksheetContentBuilder builder);

        ITitleBarBuilder StartingAt(int rowNumber, int cellNumber);

        ITitleBarBuilder WithText(string title);

        ITitleBarBuilder WithStyleApplied(Action<ICellStyle> applier);

        ITitleBarBuilder WithColumnSpan(int count);
    }
}
