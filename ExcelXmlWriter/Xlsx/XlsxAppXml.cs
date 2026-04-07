using System.Globalization;
using System.Xml.Linq;

namespace ExcelXmlWriter.Xlsx
{
    /// <summary>
    /// /docProps/app.xml
    /// </summary>
    internal sealed class XlsxAppXml : XlsxPart
    {
        readonly XNamespace propertiesNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";
        readonly XNamespace variantTypesNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";

        internal void SetSheetCount(IReadOnlyList<XlsxWorksheet> sheets)
        {
            var properties = GetRequiredElement(appXml, propertiesNamespace + "Properties");
            var titlesOfParts = GetRequiredElement(properties, propertiesNamespace + "TitlesOfParts");
            var titlesVector = GetRequiredElement(titlesOfParts, variantTypesNamespace + "vector");

            titlesVector.RemoveNodes();
            foreach (var sheet in sheets)
            {
                titlesVector.Add(new XElement(variantTypesNamespace + "lpstr", sheet.Sheetname));
            }

            GetRequiredAttribute(titlesVector, "size").Value = sheets.Count.ToString(CultureInfo.InvariantCulture);

            var headingPairs = GetRequiredElement(properties, propertiesNamespace + "HeadingPairs");
            var headingVector = GetRequiredElement(headingPairs, variantTypesNamespace + "vector");
            var worksheetCountNode = headingVector
                .Elements(variantTypesNamespace + "variant")
                .Select(variant => variant.Element(variantTypesNamespace + "i4"))
                .FirstOrDefault(value => value is not null)
                ?? throw new InvalidOperationException("The app.xml heading pairs structure is missing the worksheet count node.");

            worksheetCountNode.Value = sheets.Count.ToString(CultureInfo.InvariantCulture);
        }

        internal XlsxAppXml()
        {
            appXml.Add(
                new XElement(
                    propertiesNamespace + "Properties",
                    new XAttribute(XNamespace.Xmlns + "vt", variantTypesNamespace.NamespaceName),
                    new XElement(propertiesNamespace + "Application", "Microsoft Excel"),
                    new XElement(propertiesNamespace + "DocSecurity", "0"),
                    new XElement(propertiesNamespace + "ScaleCrop", "false"),
                    new XElement(
                        propertiesNamespace + "HeadingPairs",
                        new XElement(
                            variantTypesNamespace + "vector",
                            new XAttribute("size", 2),
                            new XAttribute("baseType", "variant"),
                            new XElement(
                                variantTypesNamespace + "variant",
                                new XElement(variantTypesNamespace + "lpstr", "Worksheets")),
                            new XElement(
                                variantTypesNamespace + "variant",
                                new XElement(variantTypesNamespace + "i4", 0)))),
                    new XElement(
                        propertiesNamespace + "TitlesOfParts",
                        new XElement(
                            variantTypesNamespace + "vector",
                            new XAttribute("size", 0),
                            new XAttribute("baseType", "lpstr"),
                            new XElement(variantTypesNamespace + "lpstr", "placeholder"))),
                    new XElement(propertiesNamespace + "Company", Environment.UserName),
                    new XElement(propertiesNamespace + "LinksUpToDate", false),
                    new XElement(propertiesNamespace + "SharedDoc", false),
                    new XElement(propertiesNamespace + "HyperlinksChanged", false),
                    new XElement(propertiesNamespace + "AppVersion", "12.0000")));
        }

        static XElement GetRequiredElement(XContainer container, XName name)
        {
            return container.Element(name)
                ?? throw new InvalidOperationException($"Expected XML element '{name}' was not found.");
        }

        static XAttribute GetRequiredAttribute(XElement element, XName name)
        {
            return element.Attribute(name)
                ?? throw new InvalidOperationException($"Expected XML attribute '{name}' was not found.");
        }
    }
}
