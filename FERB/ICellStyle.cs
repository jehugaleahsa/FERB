﻿using System.Drawing;

namespace FERB
{
    public interface ICellStyle
    {
        bool Bold { get; set; }

        bool Italic { get; set; }

        float FontSize { get; set; }

        bool WrapText { get; set; }

        void SetForegroundColor(Color color);

        void SetBackgroundColor(Color color);
    }
}
