using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FERB
{
    public static class ExcelUtilities
    {
        public static int GetColumnIndex(string columnName)
        {
            if (String.IsNullOrWhiteSpace(columnName))
            {
                throw new ArgumentException("The column name must not be blank.", "columnName");
            }
            columnName = columnName.ToUpperInvariant();
            if (columnName.Where(c => c < 'A' || c > 'Z').Any())
            {
                throw new ArgumentException("Encountered an invalid letter.", "columnName");
            }
            int total = 0;
            for (int index = 0; index != columnName.Length; ++index)
            {
                char next = columnName[columnName.Length - index - 1];
                int ordinal = next - 'A' + 1;
                int shifted = ordinal * (int)Math.Pow(26, index);
                total += shifted;
            }
            return total;
        }
    }
}
