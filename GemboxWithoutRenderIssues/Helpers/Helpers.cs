using System.Globalization;
using GemBox.Document;
using GemBox.Document.MailMerging;

namespace GemboxWithoutRenderIssues.Helpers
{
    public class Helpers
    {

        public static string StripSpecialCharacters(Exception e)
        {
            return e.Message.Replace("\r", "").Replace("\n", "").Replace("\t", "");
        }

        public static void AddWatermark(DocumentModel doc, string imagePath)
        {
            var section = doc.Sections.First();

            var watermarkPicture = new Picture(doc, imagePath);

            if (watermarkPicture.Layout != null)
            {
                var size = GetPictureDimensionsInsideRectangle(watermarkPicture.Layout, section.PageSetup.PageWidth,
                    section.PageSetup.PageHeight);

                watermarkPicture.Layout = new FloatingLayout(
                    new HorizontalPosition(HorizontalPositionType.Center, HorizontalPositionAnchor.Page),
                    new VerticalPosition(VerticalPositionType.Center, VerticalPositionAnchor.Page),
                    size)
                {
                    WrappingStyle = TextWrappingStyle.BehindText
                };
            }

            // If default header exists add Picture element to it otherwise add new default header with the Picture element.
            var header = section.HeadersFooters[HeaderFooterType.HeaderDefault];

            var watermarkParagraph = new Paragraph(doc, watermarkPicture) { ParagraphFormat = { LineSpacing = 0 } };

            if (header == null)
                section.HeadersFooters.Add(new HeaderFooter(doc, HeaderFooterType.HeaderDefault, watermarkParagraph));
            else
                header.Blocks.Add(watermarkParagraph);


            // Additionally check if first or even headers exist in case they do add a copy of the Picture element.
            section.HeadersFooters[HeaderFooterType.HeaderFirst]?.Blocks
                .Add(new Paragraph(doc, watermarkPicture.Clone()));
            section.HeadersFooters[HeaderFooterType.HeaderEven]?.Blocks
                .Add(new Paragraph(doc, watermarkPicture.Clone()));
        }

        public static Size GetPictureDimensionsInsideRectangle(Layout layout, double width, double height)
        {
            var scale = Math.Min(width / layout.Size.Width, height / layout.Size.Height);
            var size = new Size(layout.Size.Width * scale, layout.Size.Height * scale);
            return size;
        }

        private static string? GetFieldFormat(Field field, string type)
        {
            var fieldFormat = field.GetInstructionText();
            if (string.IsNullOrEmpty(fieldFormat))
                return null;

            var delimiter = " " + type + " \"";
            var startPos = fieldFormat.IndexOf(delimiter, StringComparison.Ordinal);
            if (startPos < 0)
                return null;

            startPos += delimiter.Length;
            var endPos = fieldFormat.IndexOf("\"", startPos + 1, StringComparison.Ordinal);
            if (endPos <= startPos)
                return null;

            return fieldFormat.Substring(startPos, endPos - startPos);
        }


        public static void OnMailMergeOnFieldMerging(object? sender, FieldMergingEventArgs args)
        {
            var fieldName = args.FieldName ?? "";

            if (args.Value is PictureMailMergeDataSource pictureDataSource)
            {
                try
                {
                    var stream = pictureDataSource.GetStream();
                    if (stream != null)
                    {
                        var parameters = MergeFieldParameters.FromFieldName(fieldName);
                        var width = parameters.GetParameter<double>("w", 10000);
                        var height = parameters.GetParameter<double>("h", 10000);

                        var picture = new Picture(args.Document, stream);

                        if (picture.Layout != null)
                            picture.Layout.Size = GetPictureDimensionsInsideRectangle(picture.Layout, width, height);

                        args.Inline = picture;
                    }
                    else
                        args.Inline = null;

                    return;
                }
                catch (Exception e)
                {
                    args.Inline = new Run(args.Document, "Failed to load image: " + StripSpecialCharacters(e));
                }
            }

            var dateFormat = GetFieldFormat(args.Field, "\\@");
            if (!string.IsNullOrEmpty(dateFormat))
            {
                var dateValue = Convert.ToString(args.Value);
                if (!string.IsNullOrEmpty(dateValue))
                {
                    DateTime dateTime;
                    try
                    {
                        dateTime = GetDateTime(dateValue);
                    }
                    catch (Exception e)
                    {
                        args.Inline = new Run(args.Document,
                            "(PARSE ERROR: " + args.Value + " - " + e.Message + " - " + CultureInfo.CurrentCulture +
                            "/" + CultureInfo.CurrentUICulture + ")");
                        return;
                    }

                    try
                    {
                        //args.Field.InstructionInlines.Clear();
                        args.Inline = new Run(args.Document, dateTime.ToString(dateFormat))
                            { CharacterFormat = args.Field.CharacterFormat.Clone() };
                    }
                    catch (Exception e)
                    {
                        args.Inline = new Run(args.Document,
                            "(INVALID FORMAT: " + dateFormat + " - " + e.Message + "-" + CultureInfo.CurrentCulture +
                            "/" + CultureInfo.CurrentUICulture + ")");
                        return;
                    }
                }
                else
                {
                    args.Inline = new Run(args.Document, "(EMPTY VALUE)");
                }
            }

            var numberFormat = GetFieldFormat(args.Field, "\\#");
            if (!string.IsNullOrEmpty(numberFormat))
            {
                var numberValue = Convert.ToString(args.Value);
                if (!string.IsNullOrEmpty(numberValue))
                {
                    try
                    {
                        //args.Field.InstructionInlines.Clear();
                        args.Inline = new Run(args.Document, decimal.Parse(numberValue).ToString(numberFormat))
                            { CharacterFormat = args.Field.CharacterFormat.Clone() };
                    }
                    catch (Exception)
                    {
                        args.Inline = new Run(args.Document, "(INVALID FORMAT: " + numberFormat + ")");
                    }
                }
                else
                {
                    args.Inline = new Run(args.Document, "(EMPTY VALUE)");
                }
            }

            // if a value is found but it was empty, we still want to render (empty) content. cfr mail bjorn dd vr 19/02/2016 15:57
            if (args.IsValueFound && args.Value is string && string.IsNullOrWhiteSpace(Convert.ToString(args.Value)))
            {
                var omitIfEmpty = false; // Default
                if (fieldName.EndsWith("$omitIfEmpty"))
                    omitIfEmpty = true;
                else if (fieldName.EndsWith("$dontOmitIfEmpty"))
                    omitIfEmpty = false;

                if (!omitIfEmpty)
                    args.Inline = new Run(args.Document, "") { CharacterFormat = args.Field.CharacterFormat.Clone() };
            }

            if (fieldName.EndsWith("$pageBreakIfNotEmpty"))
            {
                var hasValue = args.IsValueFound;
                if (hasValue)
                {
                    if (args.Value is string && string.IsNullOrWhiteSpace(Convert.ToString(args.Value)))
                        hasValue = false;
                }

                if (hasValue)
                    args.Inline = new SpecialCharacter(args.Document, SpecialCharacterType.PageBreak);
            }
        }

        private static DateTime GetDateTime(string str)
        {
            var formats = new[]
            {
                "yyyy-MM-dd",
                "yyyy-MM-dd HH:mm",
                "yyyy-MM-ddTHH:mm",
                "yyyy-MM-ddTHH:mm:ss"
            };

            foreach (var format in formats)
                if (DateTime.TryParseExact(str, format, null, DateTimeStyles.None, out var date))
                    return date;

            return DateTime.Parse(str, new CultureInfo("nl-BE"));
        }
    }
}