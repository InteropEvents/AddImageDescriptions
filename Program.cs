
using Azure;
using Azure.AI.Vision.ImageAnalysis;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;

string pptxPath = @"C:\source\demo\demo-images.pptx";
string endpoint = Environment.GetEnvironmentVariable("VISION_ENDPOINT") ?? throw new ArgumentNullException("VISION_ENDPOINT");
string key = Environment.GetEnvironmentVariable("VISION_KEY") ?? throw new ArgumentNullException("VISION_KEY");

ImageAnalysisClient client = new ImageAnalysisClient(new Uri(endpoint), new AzureKeyCredential(key));

using (var presentation = PresentationDocument.Open(pptxPath, true))
{
    var slideParts = presentation?.PresentationPart?.SlideParts;

    if (slideParts is not null)
    {
        foreach (var slide in slideParts)
        {
            IEnumerable<Picture> pictures = slide.Slide.Descendants<Picture>();

            if (pictures is not null)
            {
                foreach (var picture in pictures)
                {
                    var nVPP = picture.NonVisualPictureProperties;
                    var cNvPr = nVPP?.NonVisualDrawingProperties;

                    if (cNvPr is not null)
                    {
                        var desc = cNvPr.Description;

                        if (desc is null)
                        {
                            string? relationshipId = picture.BlipFill?.Blip?.Embed?.Value;

                            if (relationshipId is null)
                            {
                                continue;
                            }

                            ImagePart imagePart = (ImagePart)slide.GetPartById(relationshipId);

                            using (Stream imageStream = imagePart.GetStream())
                            {
                                BinaryData binaryData = BinaryData.FromStream(imageStream);
                                ImageAnalysisResult result = client.Analyze(binaryData, VisualFeatures.Caption);

                                Console.WriteLine("Image analysis results:");
                                Console.WriteLine(" Caption:");
                                Console.WriteLine($"   '{result.Caption.Text}', Confidence {result.Caption.Confidence:F4}");
                            }
                            //using (Stream stream = picture.BlipFill.Blip.get)
                            //{

                            //}

                            //if (desc is not null)
                            //{
                            //    Console.WriteLine($"Name: {cNvPr.Name} Description: {desc}");
                            //}
                            //else
                            //{
                            //    Console.WriteLine($"Name: {cNvPr.Name}, No Description");
                            //}
                        }
                    }
                }
            }
        }
    }
}