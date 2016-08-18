using System;
using System.Drawing;
using System.IO;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;

namespace FERB
{
    internal class ImageBuilder : IImageBuilder, IWorksheetContent
    {
        private readonly string name;
        private IWorksheetContentBuilder predecessor;
        private int rowNumber;
        private int cellNumber;
        private Func<Image> imageAccessor;
        private int cellCount;
        private double scaling;

        internal ImageBuilder(string name)
        {
            this.name = name;
            this.rowNumber = 1;
            this.cellCount = 1;
            this.scaling = 1.0d;
        }

        IWorksheetContentBuilder IWorksheetContentBuilder.Predecessor
        {
            get { return predecessor; }
        }

        public IImageBuilder StartingAfter(IWorksheetContentBuilder builder)
        {
            this.predecessor = builder;
            return this;
        }

        public IImageBuilder StartingAt(int rowNumber, int cellNumber)
        {
            if (rowNumber < 1)
            {
                throw new ArgumentOutOfRangeException("rowNumber", "The row number cannot be less than one.");
            }
            if (cellNumber < 0)
            {
                throw new ArgumentOutOfRangeException("cellNumber", "The cell number cannot be negative.");
            }
            this.rowNumber = rowNumber;
            this.cellNumber = cellNumber;
            return this;
        }

        public IImageBuilder WithColumnSpan(int count)
        {
            if (count < 1)
            {
                throw new ArgumentOutOfRangeException("count", "The count cannot be less than one.");
            }
            this.cellCount = count;
            return this;
        }

        public IImageBuilder ScaledTo(int percent)
        {
            if (percent < 1)
            {
                throw new ArgumentOutOfRangeException("percent", "The scale percent cannot be less than one.");
            }
            this.scaling = percent / 100d;
            return this;
        }

        public IImageBuilder WithImage(Stream imageStream)
        {
            if (imageStream == null)
            {
                throw new ArgumentNullException("imageStream");
            }
            this.imageAccessor = () => Image.FromStream(imageStream);
            return this;
        }

        public IImageBuilder WithImage(byte[] imageData)
        {
            if (imageData == null)
            {
                throw new ArgumentNullException("imageData");
            }
            this.imageAccessor = () => Image.FromStream(new MemoryStream(imageData));
            return this;
        }

        public IImageBuilder WithImage(string imagePath)
        {
            if (String.IsNullOrWhiteSpace(imagePath))
            {
                throw new ArgumentException("The image path cannot be blank.", "imagePath");
            }
            this.imageAccessor = () => Image.FromFile(imagePath);
            return this;
        }

        int IWorksheetContent.Save(ExcelWorksheet worksheet, int rowOffset)
        {
            int currentRow = rowOffset + rowNumber;
            if (imageAccessor == null)
            {
                return currentRow;
            }
            Image image = imageAccessor();

            ExcelRange range = worksheet.Cells[currentRow, cellNumber + 1, currentRow, cellNumber + cellCount];
            range.Merge = true;

            Graphics graphics = Graphics.FromImage(image);
            worksheet.Row(currentRow).Height = image.Height * scaling * 72 / graphics.DpiY;

            ExcelPicture picture = worksheet.Drawings.AddPicture(name, image);
            picture.SetPosition(currentRow - 1, 0, cellNumber, 0);
            picture.SetSize((int)(scaling * 100));
            return currentRow;
        }
    }
}
