using System;
using System.IO;

namespace FERB
{
    public interface IWorkbookBuilder
    {
        IWorksheetBuilder AddWorksheet(string name);

        void SaveTo(Stream stream, Stream template = null);
    }
}
