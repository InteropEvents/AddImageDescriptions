
using Azure;
using Azure.AI.Vision.ImageAnalysis;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;

string pptxPath = @"C:\source\delete-this\test.pptx";
string endpoint = Environment.GetEnvironmentVariable("VISION_ENDPOINT") ?? throw new ArgumentNullException("VISION_ENDPOINT");
string key = Environment.GetEnvironmentVariable("VISION_KEY") ?? throw new ArgumentNullException("VISION_KEY");

ImageAnalysisClient client = new ImageAnalysisClient(new Uri(endpoint), new AzureKeyCredential(key));

using (PresentationDocument presentation = PresentationDocument.Open(pptxPath, true))
{
    var foo = presentation.PresentationPart.GetPartsOfType<FontPart>();

    presentation.PresentationPart.DeleteParts<FontPart>(foo);


    //IEnumerable<SlidePart>? slideParts = presentation?.PresentationPart?.SlideParts;

    //if (slideParts is not null)
    //{
    //    foreach (SlidePart slide in slideParts)
    //    {
    //        IEnumerable<Picture> pictures = slide.Slide.Descendants<Picture>();

    //        if (pictures is not null)
    //        {
    //            foreach (Picture picture in pictures)
    //            {
    //                NonVisualDrawingProperties? cNvPr = picture.NonVisualPictureProperties?.NonVisualDrawingProperties;

    //                if (cNvPr is not null)
    //                {
    //                    StringValue? description = cNvPr.Description;

    //                    string? relationshipId = picture.BlipFill?.Blip?.Embed?.Value;

    //                    if (relationshipId is null)
    //                    {
    //                        continue;
    //                    }

    //                    ImagePart imagePart = (ImagePart)slide.GetPartById(relationshipId);

    //                    using (Stream imageStream = imagePart.GetStream())
    //                    {
    //                        BinaryData binaryData = BinaryData.FromStream(imageStream);
    //                        ImageAnalysisResult result = client.Analyze(binaryData, VisualFeatures.Caption);

    //                        Console.WriteLine("Image analysis results:");
    //                        Console.WriteLine(" Caption:");
    //                        Console.WriteLine($"   '{result.Caption.Text}', Confidence {result.Caption.Confidence:F4}");

    //                        cNvPr.Description = result.Caption.Text;
    //                    }

    //                    if (description is not null)
    //                    {
    //                        Console.WriteLine($"Name: {cNvPr.Name} Description: {description}");
    //                    }
    //                    else
    //                    {
    //                        Console.WriteLine($"Name: {cNvPr.Name}, No Description");
    //                    }
    //                }
    //            }
    //        }
    //    }
    //}
}