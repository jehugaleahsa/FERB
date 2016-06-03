using OfficeOpenXml;

namespace FERB
{
    internal interface IWorksheetContent
    {
        int Save(ExcelWorksheet worksheet, int rowOffset);
    }
}
