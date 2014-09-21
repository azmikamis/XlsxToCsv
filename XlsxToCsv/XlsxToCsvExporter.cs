using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace XlsxToCsv
{
    public class XlsxToCsvExporter
    {
        SharedStringItem[] sharedStringItems;
        CellFormat[] cellFormats;

        public void Export(string excelpath, string worksheetName, string destinationDir)
        {
            Directory.CreateDirectory(destinationDir);

            using (SpreadsheetDocument myDoc = SpreadsheetDocument.Open(excelpath, true))
            {
                WorkbookPart workbookPart = myDoc.WorkbookPart;

                SharedStringTablePart sharedStringPart = workbookPart.SharedStringTablePart;
                if (sharedStringPart != null)
                    sharedStringItems = sharedStringPart.SharedStringTable.Elements<SharedStringItem>().ToArray();

                WorkbookStylesPart stylesPart = workbookPart.WorkbookStylesPart;
                cellFormats = workbookPart.WorkbookStylesPart.Stylesheet.CellFormats.Elements<CellFormat>().ToArray();

                Sheets sheets = workbookPart.Workbook.Sheets;

                foreach (OpenXmlElement sheet in sheets)
                {
                    string sheetName = sheet.GetAttributes().Single(attr => attr.LocalName == "name").Value;

                    if (worksheetName != "" && sheetName != worksheetName)
                        continue;

                    string sheetId = sheet.GetAttributes().Single(attr => attr.LocalName == "id").Value;
                    WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheetId);
                    ExportWorksheetPart(worksheetPart, Path.Combine(destinationDir, sheetName + ".csv"));
                }
            }
        }

        public void Export(string excelpath, string destinationDir)
        {
            Export(excelpath, "", destinationDir);
        }

        private void ExportWorksheetPart(WorksheetPart worksheetPart, string filename)
        {
            string sheetStartCellRef = "", sheetEndCellRef = "";
            string sheetStartColumnName = "", sheetEndColumnName = "", cellColumnName = "";
            int sheetStartColumnIndex = 0, sheetEndColumnIndex = 0;
            int currColumnIndex = 0, cellColumnIndex = 0;
            int rowCount = 0;
            bool hasCellValue = false;
            bool isSharedString = false;
            int numberFormatToApply = -1;
            string text = "";
            Dictionary<string, int> headers = new Dictionary<string, int>();
            StringBuilder lineBuilder = new StringBuilder();

            using (StreamWriter writer = new StreamWriter(filename))
            {
                OpenXmlReader reader = OpenXmlReader.Create(worksheetPart);

                while (reader.Read())
                {
                    #region SheetDimension

                    if (reader.LocalName == "dimension")
                    {
                        if (reader.IsStartElement)
                        {
                            var sheetRefAttr = reader.Attributes.SingleOrDefault(attr => attr.LocalName == "ref");
                            string sheetRefValue = sheetRefAttr.Value;
                            sheetStartCellRef = sheetRefValue.Split(':')[0];
                            sheetEndCellRef = sheetRefValue.Split(':')[1];
                            sheetStartColumnName = GetColumnNameFromCellAddress(sheetStartCellRef);
                            sheetEndColumnName = GetColumnNameFromCellAddress(sheetEndCellRef);
                            sheetStartColumnIndex = ColumnIndexFromName(sheetStartColumnName);
                            sheetEndColumnIndex = ColumnIndexFromName(sheetEndColumnName);
                            continue;
                        }
                        else
                        {
                            continue;
                        }
                    }

                    #endregion SheetDimension

                    #region Row

                    if (reader.LocalName == "row")
                    {
                        if (reader.IsStartElement)
                        {
                            continue;
                        }
                        else
                        {
                            // Sometimes blank cells do not create cell elements
                            // Fill missing cell elements at end of row
                            if (currColumnIndex != sheetEndColumnIndex)
                            {
                                while (currColumnIndex < sheetEndColumnIndex)
                                {
                                    // check end of row
                                    if (currColumnIndex == sheetEndColumnIndex - 1)
                                        lineBuilder.Append("\"\"\n");
                                    else
                                        lineBuilder.Append("\"\",");
                                    currColumnIndex++;
                                }
                            }
                            string line = lineBuilder.ToString();
                            writer.Write(line);

                            rowCount++;

                            // reset counters at end of row
                            currColumnIndex = 0;
                            lineBuilder.Length = 0;
                            continue;
                        }
                    }

                    #endregion Row

                    #region Cell

                    if (reader.LocalName == "c")
                    {
                        if (reader.IsStartElement)
                        {
                            currColumnIndex++;
                            var cellRefAttr = reader.Attributes.SingleOrDefault(attr => attr.LocalName == "r");
                            string cellRefValue = cellRefAttr.Value;
                            cellColumnName = GetColumnNameFromCellAddress(cellRefValue);
                            cellColumnIndex = ColumnIndexFromName(cellColumnName);
                            // Sometimes blank cells do not create cell elements
                            // Fill missing cell elements at beginning of next cell
                            if (currColumnIndex != cellColumnIndex)
                            {
                                while (currColumnIndex < cellColumnIndex)
                                {
                                    // check end of row
                                    if (currColumnIndex == sheetEndColumnIndex)
                                        lineBuilder.Append("\"\"\n");
                                    else
                                        lineBuilder.Append("\"\",");
                                    currColumnIndex++;
                                }
                            }

                            var typeAttr = reader.Attributes.SingleOrDefault(attr => attr.LocalName == "t");
                            string typeValue = typeAttr.Value;
                            if (typeValue == "s") isSharedString = true;

                            if (!isSharedString)
                            {
                                var styleAttr = reader.Attributes.SingleOrDefault(attr => attr.LocalName == "s");
                                string styleValue = styleAttr.Value;
                                if (styleValue != null && styleValue != "")
                                    numberFormatToApply =
                                        Convert.ToInt32(cellFormats[Convert.ToInt16(styleValue)].NumberFormatId.Value);
                            }
                            continue;
                        }
                        else
                        {
                            // Sometimes blank cells have cell elements but do not create cell value elements
                            // Fill missing cell value element at end of cell
                            if (!hasCellValue)
                            {
                                // check end of row
                                if (currColumnIndex == sheetEndColumnIndex)
                                    lineBuilder.Append("\"\"\n");
                                else
                                    lineBuilder.Append("\"\",");
                            }
                            // reset cell flags at end of cell
                            hasCellValue = false;
                            isSharedString = false;
                            numberFormatToApply = -1;
                            continue;
                        }
                    }

                    #endregion Cell

                    #region CellValue

                    if (reader.LocalName == "v")
                    {
                        if (reader.IsStartElement)
                        {
                            hasCellValue = true;
                            if (isSharedString)
                                text = sharedStringItems[Convert.ToInt32(reader.GetText())].InnerText;
                            else
                                text = reader.GetText();

                            text = text.Replace("\n", " ");
                            text = text.Replace("\"", "'");

                            if (numberFormatToApply > -1)
                            {
                                int testInt = 0;
                                DateTime testDateTime;

                                // TODO Add new number format handlers when faced
                                switch (numberFormatToApply)
                                {
                                    case 0:
                                        testInt = Convert.ToInt32(text);
                                        text = testInt.ToString("0");
                                        break;
                                    case 22:
                                        testDateTime = DateTime.FromOADate(Convert.ToDouble(text));
                                        text = testDateTime.ToString("d/M/yyyy h:mm");
                                        break;
                                    case 14:
                                        testDateTime = DateTime.FromOADate(Convert.ToDouble(text));
                                        text = testDateTime.ToString("d/M/yyyy");
                                        break;
                                    case 49:
                                        break;
                                    default:

                                        break;
                                }
                            }

                            // For header row, sometimes it is a requirement to have unique names
                            // If it is not, this part  and the headers dictionary can be removed
                            if (rowCount == 0)
                            {
                                if (!headers.ContainsKey(text))
                                    headers.Add(text, 1);
                                else
                                {
                                    headers[text] = headers[text] + 1;
                                    text = text + headers[text];
                                }
                            }
                            // check end of row
                            if (currColumnIndex == sheetEndColumnIndex)
                                lineBuilder.Append("\"" + text + "\"\n");
                            else
                                lineBuilder.Append("\"" + text + "\",");
                            continue;
                        }
                        else
                        {
                            continue;
                        }
                    }

                    #endregion CellValue

                    #region InlineString

                    if (reader.LocalName == "is")
                    {
                        if (reader.IsStartElement)
                        {
                            hasCellValue = true;
                            InlineString inlineString = reader.LoadCurrentElement() as InlineString;
                            text = inlineString.InnerText;

                            text = text.Replace("\n", " ");
                            text = text.Replace("\"", "'");

                            if (rowCount == 0)
                            {
                                if (!headers.ContainsKey(text))
                                    headers.Add(text, 1);
                                else
                                {
                                    headers[text] = headers[text] + 1;
                                    text = text + headers[text];
                                }
                            }
                            // check end of row
                            if (currColumnIndex == sheetEndColumnIndex)
                                lineBuilder.Append("\"" + text + "\"\n");
                            else
                                lineBuilder.Append("\"" + text + "\",");
                            continue;
                        }
                        else
                        {
                            continue;
                        }
                    }

                    #endregion InlineString
                }
                reader.Close();
            }
        }

        private string GetColumnNameFromCellAddress(string cellAddress)
        {
            return string.Join(null, Regex.Split(cellAddress, "[^A-Z]"));
        }

        private int ColumnIndexFromName(string columnName)
        {
            int columnIndex = 0;
            int index = 0;
            int length = columnName.Length - 1;

            while (length >= 0)
            {
                columnIndex = columnIndex +
                                (Convert.ToInt32(columnName[index]) - 64) *
                                Convert.ToInt32(Math.Pow(26, length));
                index++;
                length--;
            }

            return columnIndex;
        }
    }
}