using Microsoft.AspNetCore.Mvc;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using System.IO;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace DocExporApp.Controllers
{
    public class DocController : Controller
    {
        private readonly string _templatePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Sample Company Letter - Deceased.docx");

        [HttpGet]
        public IActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public IActionResult Generate(string Title, string Firstname, string Lastname, IFormFile image)
        {
            if (string.IsNullOrEmpty(Title) || string.IsNullOrEmpty(Firstname) || string.IsNullOrEmpty(Lastname))
            {
                ViewBag.Error = "All fields are required";
                return View("Index");
            }

            string outputPath = Path.Combine(Path.GetTempPath(), $"Generated_{Firstname}.docx");

           
            using (var stream = new MemoryStream())
            {
                using (var fileStream = new FileStream(_templatePath, FileMode.Open, FileAccess.Read))
                {
                    fileStream.CopyTo(stream);
                }

                stream.Position = 0;

                using (var doc = WordprocessingDocument.Open(stream, true))
                {
                    var body = doc.MainDocumentPart.Document.Body;

                   
                    foreach (var text in body.Descendants<Text>())
                    {
                        text.Text = text.Text.Replace("{Title}", Title)
                                             .Replace("{Firstname}", Firstname)
                                             .Replace("{Lastname}", Lastname);
                    }

                   
                    if (image != null && body.InnerText.Contains("{Image}"))
                    {
                        var imagePart = doc.MainDocumentPart.AddImagePart(DocumentFormat.OpenXml.Packaging.ImagePartType.Jpeg);
                        using (var imgStream = image.OpenReadStream())
                        {
                            imagePart.FeedData(imgStream);
                        }

                        var imagePartId = doc.MainDocumentPart.GetIdOfPart(imagePart);

                       
                        foreach (var text in body.Descendants<Text>())
                        {
                            if (text.Text.Contains("{Image}"))
                            {
                                text.Text = text.Text.Replace("{Image}", "");

                                var run = text.Parent as Run; 

                                if (run != null)
                                {
                                    run.AppendChild(CreateImageElement(imagePartId)); 
                                }
                                break;
                            }
                        }
                    }

                    doc.MainDocumentPart.Document.Save();
                }

               
                System.IO.File.WriteAllBytes(outputPath, stream.ToArray());
            }

            var fileBytes = System.IO.File.ReadAllBytes(outputPath);
            return File(fileBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", $"Generated_{Firstname}.docx");
        }


        private Drawing CreateImageElement(string relationshipId)
        {
            return new Drawing(
                new DW.Inline(
                    new DW.Extent() { Cx = 990000L, Cy = 792000L }, 
                    new DW.EffectExtent()
                    {
                        LeftEdge = 0L,
                        TopEdge = 0L,
                        RightEdge = 0L,
                        BottomEdge = 0L
                    },
                    new DW.DocProperties()
                    {
                        Id = (UInt32Value)1U,
                        Name = "Inserted Image"
                    },
                    new DW.NonVisualGraphicFrameDrawingProperties(
                        new A.GraphicFrameLocks() { NoChangeAspect = true }),
                    new A.Graphic(
                        new A.GraphicData(
                            new PIC.Picture(
                                new PIC.NonVisualPictureProperties(
                                    new PIC.NonVisualDrawingProperties()
                                    {
                                        Id = (UInt32Value)0U,
                                        Name = "InsertedImage.jpg"
                                    },
                                    new PIC.NonVisualPictureDrawingProperties()),
                                new PIC.BlipFill(
                                    new A.Blip()
                                    {
                                        Embed = relationshipId,
                                        CompressionState = A.BlipCompressionValues.Print
                                    },
                                    new A.Stretch(new A.FillRectangle())),
                                new PIC.ShapeProperties(
                                    new A.Transform2D(
                                        new A.Offset() { X = 0L, Y = 0L },
                                        new A.Extents() { Cx = 990000L, Cy = 792000L }),
                                    new A.PresetGeometry(new A.AdjustValueList())
                                    { Preset = A.ShapeTypeValues.Rectangle }))
                        )
                        { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
                )
                {
                    DistanceFromTop = (UInt32Value)0U,
                    DistanceFromBottom = (UInt32Value)0U,
                    DistanceFromLeft = (UInt32Value)0U,
                    DistanceFromRight = (UInt32Value)0U,
                    EditId = "50D07946"
                });
        }
    }
}
