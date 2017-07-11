using System.Drawing;

namespace FERB
{
    public interface ICellStyle
    {
        bool Bold { get; set; }

        bool Italic { get; set; }

        float FontSize { get; set; }

        bool WrapText { get; set; }

        HorizontalAlignment HorizontalAlignment { get; set; }

        VerticalAlignment VerticalAlignment { get; set; }

        string Format { get; set; }

        void SetForegroundColor(Color color);

        void SetBackgroundColor(Color color);

        void SetBorder(Border border = Border.Outline, BorderStyle borderStyle = BorderStyle.Thin, Color? color = null);
    }
}
