using System.Globalization;
using System.Xml.Linq;

namespace ExcelXmlWriter.Xlsx
{
    /// <summary>
    /// /xl/workbook.xml
    /// </summary>
    internal sealed class XlsxWorkbookMetadata : XlsxPart
    {
        readonly XNamespace relationshipNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        readonly XNamespace workbookNamespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

        internal void SetSheetCount(IReadOnlyList<XlsxWorksheet> sheets)
        {
            var workbook = GetRequiredElement(appXml, workbookNamespace + "workbook");
            var sheetsElement = GetRequiredElement(workbook, workbookNamespace + "sheets");

            int order = 1;
            foreach (var sheet in sheets)
            {
                sheetsElement.Add(
                    new XElement(
                        workbookNamespace + "sheet",
                        new XAttribute("name", sheet.Sheetname),
                        new XAttribute("sheetId", order.ToString(CultureInfo.InvariantCulture)),
                        new XAttribute(relationshipNamespace + "id", sheet.RelationshipId)));
                order++;
            }
        }

        internal XlsxWorkbookMetadata()
        {
            appXml.Add(
                new XElement(
                    workbookNamespace + "workbook",
                    new XAttribute(XNamespace.Xmlns + "r", relationshipNamespace.NamespaceName),
                    new XElement(
                        workbookNamespace + "fileVersion",
                        new XAttribute("appName", "xl"),
                        new XAttribute("lastEdited", "4"),
                        new XAttribute("lowestEdited", "4"),
                        new XAttribute("rupBuild", "4506")),
                    new XElement(
                        workbookNamespace + "workbookPr",
                        new XAttribute("defaultThemeVersion", "124226")),
                    new XElement(
                        workbookNamespace + "bookViews",
                        new XElement(
                            workbookNamespace + "workbookView",
                            new XAttribute("xWindow", "360"),
                            new XAttribute("yWindow", "120"),
                            new XAttribute("windowWidth", "22860"),
                            new XAttribute("windowHeight", "11385"),
                            new XAttribute("activeTab", "0"))),
                    new XElement(workbookNamespace + "sheets"),
                    new XElement(
                        workbookNamespace + "calcPr",
                        new XAttribute("calcId", "125725"))));
        }

        static XElement GetRequiredElement(XContainer container, XName name)
        {
            return container.Element(name)
                ?? throw new InvalidOperationException($"Expected XML element '{name}' was not found.");
        }
    }
}
