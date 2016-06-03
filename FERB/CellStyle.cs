﻿using System;
using System.Drawing;
using OfficeOpenXml.Style;

namespace FERB
{
    internal class CellStyle : ICellStyle
    {
        private readonly ExcelStyle style;

        public CellStyle(ExcelStyle style)
        {
            this.style = style;
        }

        public bool Bold 
        {
            get { return style.Font.Bold; }
            set { style.Font.Bold = value; }
        }

        public float FontSize 
        {
            get { return style.Font.Size; }
            set { style.Font.Size = value; }
        }

        public bool WrapText
        {
            get { return style.WrapText; }
            set { style.WrapText = value; }
        }

        public void SetBackgroundColor(Color color)
        {
            style.Fill.PatternType = ExcelFillStyle.Solid;
            style.Fill.BackgroundColor.SetColor(color);
        }
    }
}