#nullable disable // At this time the IMailMergeDataSource interface is not nullable

using System.Diagnostics.CodeAnalysis;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using GemBox.Document.MailMerging;

namespace GemboxWithoutRenderIssues.Helpers
{
    // THE CODE BELOW HAS BEEN COPIED FROM RADIUS.REPORTING
    [SuppressMessage("Globalization", "CA1310: Specify StringComparison for correctness", Justification = "Code copied from Radius.Reporting")]
    [SuppressMessage("Performance", "CA1845: Use span-based 'string.Concat'", Justification = "Code copied from Radius.Reporting")]
    public sealed class XmlMailMergeDataSource : IMailMergeDataSource
    {
        private readonly XmlMailMergeDataSource _parent;
        private readonly bool _debug;
        private readonly IEnumerator<XElement> _enumerator;

        public XmlMailMergeDataSource(XElement element, bool debug = false)
            : this(null, new List<XElement> { element }, "", debug)
        {
        }

        public XmlMailMergeDataSource(XmlMailMergeDataSource parent, List<XElement> elements, string name, bool debug)
        {
            Name = name;
            _enumerator = elements.GetEnumerator();
            _parent = parent;
            _debug = debug;
        }

        public string Name { get; }

        public bool MoveNext()
        {
            var moveNext = _enumerator.MoveNext();

            //if (_debug)
            //{
            //    if (moveNext)
            //        Debug.WriteLine(Name + ": MoveNext() -> next is " + _enumerator.Current.Name);
            //    else
            //        Debug.WriteLine(Name + ": MoveNext() -> no next");
            //}

            return moveNext;
        }

        public bool TryGetValue(string valueName, out object value)
        {
            //if (_debug)
            //    Debug.Write(Name + ": TryGetValue(" + valueName + "): ");
            value = null;

            var element = CurrentElement;
            var strippedValueName = Regex.Replace(valueName, @"\(.*\)", "");

            if (strippedValueName.ToLowerInvariant() == "$whitespace")
            {
                value = " ";
                return true;
            }

            if (strippedValueName.StartsWith("$parentContext."))
            {
                if (_parent != null)
                    return _parent.TryGetValue(valueName.Substring("$parentContext.".Length), out value);

                //if (_debug)
                //    Debug.WriteLine("$parentContext requested without parent!");

                return true;
            }

            // Special handling for tags ending in "Container": RangeStart:Quality.DetailsContainer will return current root (aka not start a new range)
            // when Quality.Details has child elements. Will return nothing if no child elements found.
            if (strippedValueName.EndsWith("Container") || strippedValueName.EndsWith("$contains"))
                return HandleContainer(strippedValueName, out value);

            if (strippedValueName.EndsWith("$notEmpty"))
                return HandleNotEmpty(strippedValueName, out value);

            if (strippedValueName.EndsWith("$asWhitespace"))
                return HandleAsWhitespace(strippedValueName, out value);

            strippedValueName = RemoveModifier(strippedValueName);
            var propElement = GetElement(strippedValueName, element);
            if (propElement == null)
            {
                //if (_debug)
                //    Debug.WriteLine("(not found)");

                return false;
            }

            var elementType = propElement.Attribute("Type");
            if (elementType != null && elementType.Value == "ImageRef")
            {
                value = new PictureMailMergeDataSource(propElement);
                //if (_debug)
                //    Debug.WriteLine("picture: " + propElement.Name.LocalName);
                return true;
            }

            var childElements = GetChildElements(propElement).ToList();
            if (childElements.Any())
            {
                value = new XmlMailMergeDataSource(this, childElements, (string.IsNullOrEmpty(Name) ? "" : Name + ".") + propElement.Name.LocalName, _debug);
                //if (_debug)
                //    Debug.WriteLine("datasource: " + propElement.Name.LocalName);
                return true;
            }

            value = propElement.Value;
            //if (_debug)
            //{
            //    var v = Convert.ToString(value);
            //    Debug.WriteLine("(value: " + (v != null ? v.Substring(0, Math.Min(10, v.Length)) : "null") + ")");
            //}

            return true;
        }

        private XElement CurrentElement => _enumerator.Current;

        private static string RemoveModifier(string fieldName)
        {
            var dollarPos = fieldName.LastIndexOf('$');
            return dollarPos <= 0 ? fieldName : fieldName.Substring(0, dollarPos);
        }

        private static IEnumerable<XElement> GetChildElements(XElement element)
        {
            var singleElementName = element.Name.LocalName;
            var possibleChildElementNames = new List<string> { singleElementName };
            if (!singleElementName.EndsWith("s")) return element.Elements().Where(e => possibleChildElementNames.Contains(e.Name.LocalName));
            var withoutS = singleElementName.Substring(0, singleElementName.Length - 1);
            possibleChildElementNames.Add(withoutS);
            if (withoutS.EndsWith("ie"))
                possibleChildElementNames.Add(withoutS.Substring(0, withoutS.Length - 2) + "y");

            return element.Elements().Where(e => possibleChildElementNames.Contains(e.Name.LocalName));
        }

        private bool HandleContainer(string valueName, out object value)
        {
            var element = _enumerator.Current;
            var containerElementName = valueName.Substring(0, valueName.Length - "Container".Length);
            var containerElement = GetElement(containerElementName, element);
            if (containerElement != null && GetChildElements(containerElement).Any())
            {
                //if (_debug)
                //    Debug.WriteLine("(container check: ok)");

                value = new XmlMailMergeDataSource(this, new List<XElement> { element }, Name, _debug);
                return true;
            }

            //if (_debug)
            //    Debug.WriteLine("(container check: not ok)");

            value = null;
            return false;
        }

        private bool HandleAsWhitespace(string valueName, out object value)
        {
            var element = _enumerator.Current;
            var elementName = valueName.Substring(0, valueName.Length - "$asWhitespace".Length);
            var elementToCheck = GetElement(elementName, element);
            if (!string.IsNullOrEmpty(elementToCheck?.Value))
            {
                value = " ";
                return true;
            }

            value = null;
            return false;
        }

        private bool HandleNotEmpty(string valueName, out object value)
        {
            var element = _enumerator.Current;
            var elementName = valueName.Substring(0, valueName.Length - "$notEmpty".Length);
            var elementToCheck = GetElement(elementName, element);
            if (!string.IsNullOrEmpty(elementToCheck?.Value))
            {
                //if (_debug)
                //    Debug.WriteLine("(notEmpty check: ok)");

                value = new XmlMailMergeDataSource(this, new List<XElement> { element }, Name, _debug);
                return true;
            }

            //if (_debug)
            //    Debug.WriteLine("(notEmpty check: nok)");
            value = null;
            return false;
        }

        /// <summary>Resolves dotted element path.</summary>
        // TODO: XPath would have been better and faster
        private static XElement GetElement(string valueName, XElement element)
        {
            var names = valueName.Split('.');

            foreach (var name in names)
            {
                if (string.IsNullOrEmpty(name))
                    return null;

                element = element?.Element(name);

                if (element == null)
                    return null;
            }

            return element;
        }
    }
}