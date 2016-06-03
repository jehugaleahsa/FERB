using System;

namespace FERB
{
    public interface IWorksheetContentBuilder
    {
        IWorksheetContentBuilder Predecessor { get; }
    }
}
