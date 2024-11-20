using System.Xml.Linq;
using GemBox.Document;
using GemBox.Document.MailMerging;
using GemboxWithoutRenderIssues.Helpers;
using LoadOptions = GemBox.Document.LoadOptions;

ComponentInfo.SetLicense(
    "DN-2024Mar04-BLirqX5djTixX00g0XDBcUsHkI2Rqesyz9QdZjJEjEvEdrmTVKBJV25oc8C3d5vfQq93s8HzX3TCI4dm5IIYzo51aDA==A");

var workingDirectory = Environment.CurrentDirectory;
var projectDirectory = Directory.GetParent(workingDirectory).Parent.Parent.FullName;

var documentModel = DocumentModel.Load(projectDirectory + "/doc.dotx", LoadOptions.DocxDefault);

var data = XDocument.Load(projectDirectory + "/mailmerge.xml");

documentModel.MailMerge.ClearOptions = MailMergeClearOptions.RemoveEmptyParagraphs
                                       | MailMergeClearOptions.RemoveEmptyRanges
                                       | MailMergeClearOptions.RemoveEmptyTableRows
                                       | MailMergeClearOptions.RemoveUnusedFields;
documentModel.MailMerge.FieldMerging += Helpers.OnMailMergeOnFieldMerging;
documentModel.MailMerge.Execute(new XmlMailMergeDataSource(data.Root));


var processingNode = data.Root.Element("Processing");
var watermarkImagePath = processingNode?.Element("Watermark")?.Value;

if (!string.IsNullOrEmpty(watermarkImagePath))
    Helpers.AddWatermark(documentModel, watermarkImagePath);

var outStream = new MemoryStream();

documentModel.Save(outStream, new PdfSaveOptions
{
    ConformanceLevel =
        PdfConformanceLevel
            .PdfA2a // See: https://www.gemboxsoftware.com/document/docs/GemBox.Document.PdfConformanceLevel.html & https://en.wikipedia.org/wiki/PDF/A
});

await using var fileStream = File.OpenWrite(projectDirectory + "/doc.pdf");
outStream.Position = 0;
await outStream.CopyToAsync(fileStream);