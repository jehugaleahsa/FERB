using System;
using System.Collections.Generic;
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

        public bool Italic
        {
            get { return style.Font.Italic; }
            set { style.Font.Italic = value; }
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

        public HorizontalAlignment HorizontalAlignment
        {
            get { return (HorizontalAlignment)style.HorizontalAlignment; }
            set { style.HorizontalAlignment = (ExcelHorizontalAlignment)value; }
        }

        public VerticalAlignment VerticalAlignment
        {
            get { return (VerticalAlignment)style.VerticalAlignment; }
            set { style.VerticalAlignment = (ExcelVerticalAlignment)value; }
        }

        public string Format
        {
            get { return style.Numberformat.Format; }
            set { style.Numberformat.Format = value; }
        }

        public void SetForegroundColor(Color color)
        {
            style.Font.Color.SetColor(color);
        }

        public void SetBackgroundColor(Color color)
        {
            style.Fill.PatternType = ExcelFillStyle.Solid;
            style.Fill.BackgroundColor.SetColor(color);
        }

        public void SetBorder(Border border = Border.Outline, BorderStyle borderStyle = BorderStyle.Thin, Color? color = null)
        {
            List<ExcelBorderItem> items = new List<ExcelBorderItem>();
            if (border.HasFlag(Border.Bottom))
            {
                items.Add(style.Border.Bottom);                
            }
            if (border.HasFlag(Border.Top))
            {
                items.Add(style.Border.Top);
            }
            if (border.HasFlag(Border.Left))
            {
                items.Add(style.Border.Left);
            }
            if (border.HasFlag(Border.Right))
            {
                items.Add(style.Border.Right);
            }
            foreach (var item in items)
            {
                item.Style = (ExcelBorderStyle)borderStyle;
                item.Color.SetColor(color ?? Color.Black);
            }
        }
    }
}
