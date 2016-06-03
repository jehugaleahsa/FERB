using System.IO;

namespace FERB
{
    public interface IImageBuilder : IWorksheetContentBuilder
    {
        IImageBuilder StartingAfter(IWorksheetContentBuilder builder);

        IImageBuilder StartingAt(int rowNumber, int cellNumber);

        IImageBuilder WithColumnSpan(int count);

        IImageBuilder ScaledTo(int percent);

        IImageBuilder WithImage(Stream imageStream);

        IImageBuilder WithImage(string imagePath);

        IImageBuilder WithImage(byte[] imageData);
    }
}
