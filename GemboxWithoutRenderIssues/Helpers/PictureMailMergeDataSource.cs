﻿using System.Xml.Linq;

namespace GemboxWithoutRenderIssues.Helpers
{
    // THE CODE BELOW HAS BEEN COPIED FROM RADIUS.REPORTING
    public sealed class PictureMailMergeDataSource
    {
        private readonly XElement _element;

        public PictureMailMergeDataSource(XElement element)
        {
            _element = element;
        }

        public Stream? GetStream()
        {
            var type = _element.Attribute("Type")?.Value;
            switch (type) {
                case "ImageRef":
                    var pathElement = _element.Element("Path");

                    var path = pathElement?.Value;
                    if (string.IsNullOrEmpty(path))
                        return null;

                    // HACK: replace O: with \\131.0.1.170\DATA
                    // because the O: drive isn't mounted for the webserver user.
                    // path = path.Replace("O:", @"\\131.0.1.170\DATA");

                    return new FileStream(path, FileMode.Open, FileAccess.Read);
                case "Image":
                    /* handle inlined image in xml */
                    break;
                default:
                    return null;
            }

            return null;
        }
    }
}