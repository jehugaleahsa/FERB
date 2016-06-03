using System;
using System.Drawing;

namespace FERB
{
    public interface ICellStyle
    {
        bool Bold { get; set; }

        float FontSize { get; set; }

        bool WrapText { get; set; }

        void SetBackgroundColor(Color color);
    }
}
