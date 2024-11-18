using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;

namespace PptxFileSizeAnalyzer;

internal static class Commands
{
    /// <summary>
    /// Analyze the specified PowerPoint (*.pptx) file.
    /// </summary>
    /// <param name="fileName">-f, target file</param>
    public static void AnalyzePptxFile(string fileName)
    {
        if (!File.Exists(fileName))
        {
            Console.WriteLine($"File not found: {fileName}");
            return;
        }

        using var pptx = PresentationDocument.Open(fileName, false);
        var pptxSize = new FileInfo(fileName).Length;
        var pptxSizeMB = pptxSize / 1024.0 / 1024.0;
        Console.WriteLine($"File: {fileName}");
        Console.WriteLine($"Size: {pptxSize} bytes ({pptxSizeMB:F2} MB)");

        if (pptx.PresentationPart?.Presentation?.SlideIdList == null)
        {
            Console.WriteLine("No slides found in the presentation.");
            return;
        }

        int slideNumber = 1;
        long sum = 0;
        foreach (var slideId in pptx.PresentationPart.Presentation.SlideIdList.Elements<SlideId>())
        {
            string? slidePartId = slideId.RelationshipId;
            if (slidePartId == null)
            {
                throw new InvalidOperationException("Slide part ID is null.");
            }

            if (pptx.PresentationPart.GetPartById(slidePartId) is not SlidePart slidePart)
            {
                throw new InvalidOperationException("Slide part is null.");
            }

            long size = GetPartSize(slidePart);
            sum += size;
            string title = GetSlideTitle(slidePart);
            Console.WriteLine($"Slide: {slideNumber}({title}) => {size.ToReadableSize()} (total: {sum.ToReadableSize()})");

            foreach (var imagePart in slidePart.ImageParts)
            {
                long imageSize = GetPartSize(imagePart);
                Console.WriteLine($"\t - Image: {imageSize.ToReadableSize()}");
            }

            foreach (var relationship in slidePart.DataPartReferenceRelationships)
            {
                if (relationship is
                    not
                    {
                        RelationshipType: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/video",
                        DataPart: MediaDataPart mediaDataPart
                    }) continue;

                long videoSize = GetPartSize(mediaDataPart);
                Console.WriteLine($"\t - Video ({mediaDataPart.ContentType}): {videoSize.ToReadableSize()}");
            }

            slideNumber++;
        }
    }

    /// <summary>
    /// Gets the title of the specified slide.
    /// </summary>
    /// <param name="slidePart">The slide part to get the title from.</param>
    /// <returns>The title of the slide, or "Untitled Slide" if no title is found.</returns>
    private static string GetSlideTitle(SlidePart slidePart)
    {
        var titleShape = slidePart.Slide.Descendants<Shape>()
            .FirstOrDefault(s =>
                s.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape?.Type?.Value ==
                PlaceholderValues.Title);
        return titleShape?.TextBody?.InnerText ?? "Untitled Slide";
    }

    /// <summary>
    /// Gets the size of the specified OpenXmlPart.
    /// </summary>
    /// <param name="part">The OpenXmlPart to get the size of.</param>
    /// <returns>The size of the part in bytes.</returns>
    private static long GetPartSize(OpenXmlPart part)
    {
        using Stream stream = part.GetStream();
        return stream.Length;
    }

    /// <summary>
    /// Gets the size of the specified DataPart.
    /// </summary>
    /// <param name="part">The DataPart to get the size of.</param>
    /// <returns>The size of the part in bytes.</returns>
    private static long GetPartSize(DataPart part)
    {
        using Stream stream = part.GetStream();
        return stream.Length;
    }
}