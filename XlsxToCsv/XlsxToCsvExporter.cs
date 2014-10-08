using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Globalization;
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

        int sheetEndColumnIndex = 0;
        int currColumnIndex = 0, cellColumnIndex = 0;
        int rowCount = 0;
        int numberFormatToApply = -1;

        bool hasCellValue = false;
        bool isSharedString = false;

        string textDelimiter = "\"";
        string columnDelimiter = ",";

        Dictionary<string, int> headers = new Dictionary<string, int>();
        List<int> columnsToSkip = new List<int>() { 35, 36, 37, 38, 39, 40, 41 };
        StringBuilder lineBuilder = new StringBuilder();

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
            currColumnIndex = 0;
            rowCount = 0;
            headers.Clear();
            columnsToSkip.Clear();
            
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
                            HandleSheetDimensionStartElement(reader);
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
                            HandleRowEndElement(reader, writer);
                            continue;
                        }
                    }

                    #endregion Row

                    #region Cell

                    if (reader.LocalName == "c")
                    {
                        if (reader.IsStartElement)
                        {
                            HandleCellStartElement(reader);                            
                            continue;
                        }
                        else
                        {
                            HandleCellEndElement(reader);
                            continue;
                        }
                    }

                    #endregion Cell

                    #region CellValue

                    if (reader.LocalName == "v")
                    {
                        if (reader.IsStartElement)
                        {
                            HandleCellValueStartElement(reader);
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
                            HandleInlineStringStartElement(reader);
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

        private void HandleSheetDimensionStartElement(OpenXmlReader reader)
        {
            var sheetRefAttr = reader.Attributes.SingleOrDefault(attr => attr.LocalName == "ref");
            string sheetRefValue = sheetRefAttr.Value;
            string sheetStartCellRef = sheetRefValue.Split(':')[0];
            string sheetEndCellRef = sheetRefValue.Split(':')[1];
            string sheetStartColumnName = GetColumnNameFromCellAddress(sheetStartCellRef);
            string sheetEndColumnName = GetColumnNameFromCellAddress(sheetEndCellRef);
            sheetEndColumnIndex = ColumnIndexFromName(sheetEndColumnName);
        }

        private void HandleRowEndElement(OpenXmlReader reader, StreamWriter writer)
        {
            // Assume accessing next cell
            currColumnIndex++;
            // Sometimes blank cells do not create cell elements
            // Fill missing cell elements at end of row
            if (currColumnIndex != sheetEndColumnIndex)
            {
                while (currColumnIndex <= sheetEndColumnIndex)
                {
                    if (rowCount == 0)
                    {
                        columnsToSkip.Add(currColumnIndex);
                    }
                    
                    if (!columnsToSkip.Contains(currColumnIndex))
                    {
                        lineBuilder.Append(string.Format("{0}{0}", textDelimiter));

                        // check end of row
                        if (currColumnIndex != sheetEndColumnIndex - 1)
                            lineBuilder.Append(columnDelimiter);
                    }

                    currColumnIndex++;
                }
            }
            else
            {
                if (rowCount == 0)
                {
                    columnsToSkip.Add(currColumnIndex);
                }
            }
            // if we exclude columns from behind, there will be trailing commas
            if (lineBuilder.ToString().EndsWith(columnDelimiter))
                lineBuilder.Length--;
            lineBuilder.Append("\n");
            string line = lineBuilder.ToString();
            writer.Write(line);

            rowCount++;

            // reset counters at end of row
            currColumnIndex = 0;
            lineBuilder.Length = 0;
        }

        private void HandleCellValueStartElement(OpenXmlReader reader)
        {
            string text = "";
            hasCellValue = true;
            
            if (isSharedString)
                text = sharedStringItems[Convert.ToInt32(reader.GetText())].InnerText;
            else
                text = reader.GetText();

            text = text.Replace("\n", " ");
            text = text.Replace("\"", "'");

            if (numberFormatToApply > -1)
            {
                DateTime testDateTime;

                // TODO Add new number format handlers when faced
                switch (numberFormatToApply)
                {
                    case 14:
                        testDateTime = DateTime.FromOADate(Convert.ToDouble(text));
                        text = testDateTime.ToString(CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern);
                        break;
                    case 15:
                        testDateTime = DateTime.FromOADate(Convert.ToDouble(text));
                        text = testDateTime.ToString("d-MMM-yy");
                        break;
                    case 16:
                        testDateTime = DateTime.FromOADate(Convert.ToDouble(text));
                        text = testDateTime.ToString("d-MMM");
                        break;
                    case 17:
                        testDateTime = DateTime.FromOADate(Convert.ToDouble(text));
                        text = testDateTime.ToString("MMM-yy");
                        break;
                    case 18:
                        testDateTime = DateTime.FromOADate(Convert.ToDouble(text));
                        text = testDateTime.ToString("h:mm tt");
                        break;
                    case 19:
                        testDateTime = DateTime.FromOADate(Convert.ToDouble(text));
                        text = testDateTime.ToString("h:mm:ss tt");
                        break;
                    case 20:
                        testDateTime = DateTime.FromOADate(Convert.ToDouble(text));
                        text = testDateTime.ToString("h:mm");
                        break;
                    case 21:
                        testDateTime = DateTime.FromOADate(Convert.ToDouble(text));
                        text = testDateTime.ToString("h:mm:ss");
                        break;
                    case 22:
                        testDateTime = DateTime.FromOADate(Convert.ToDouble(text));
                        text = testDateTime.ToString(string.Format("{0} h:mm", CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern));
                        break;
                    default:
                        break;
                }
            }

            // For header row, sometimes it is a requirement to have unique names
            // If it is not, this part and the headers dictionary can be removed
            if (rowCount == 0)
            {
                // if header text is empty, do not write this column
                if (text == "")
                    columnsToSkip.Add(currColumnIndex);

                if (!headers.ContainsKey(text))
                    headers.Add(text, 1);
                else
                {
                    headers[text] = headers[text] + 1;
                    text = text + headers[text];
                }
            }

            if (!columnsToSkip.Contains(currColumnIndex))
            {
                lineBuilder.Append(string.Format("{0}{1}{0}", textDelimiter, text));

                // check end of row
                if (currColumnIndex != sheetEndColumnIndex)
                    lineBuilder.Append(columnDelimiter);
            }
        }

        private void HandleCellStartElement(OpenXmlReader reader)
        {
            currColumnIndex++;
            var cellRefAttr = reader.Attributes.SingleOrDefault(attr => attr.LocalName == "r");
            string cellRefValue = cellRefAttr.Value;
            string cellColumnName = GetColumnNameFromCellAddress(cellRefValue);
            cellColumnIndex = ColumnIndexFromName(cellColumnName);
            // Sometimes blank cells do not create cell elements
            // Fill missing cell elements at beginning of next cell
            if (currColumnIndex != cellColumnIndex)
            {
                while (currColumnIndex < cellColumnIndex)
                {
                    if (!columnsToSkip.Contains(currColumnIndex))
                    {
                        lineBuilder.Append(string.Format("{0}{0}", textDelimiter));

                        // check end of row
                        if (currColumnIndex != sheetEndColumnIndex)
                            lineBuilder.Append(columnDelimiter);
                    }

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
        }

        private void HandleCellEndElement(OpenXmlReader reader)
        {
            // Sometimes blank cells have cell elements but do not create cell value elements
            // Fill missing cell value element at end of cell
            if (!hasCellValue)
            {
                if (!columnsToSkip.Contains(currColumnIndex))
                {
                    lineBuilder.Append(string.Format("{0}{0}", textDelimiter));
                    // check end of row
                    if (currColumnIndex != sheetEndColumnIndex)
                        lineBuilder.Append(columnDelimiter);
                }
            }
            // reset cell flags at end of cell
            hasCellValue = false;
            isSharedString = false;
            numberFormatToApply = -1;
        }

        private void HandleInlineStringStartElement(OpenXmlReader reader)
        {
            hasCellValue = true;
            InlineString inlineString = reader.LoadCurrentElement() as InlineString;
            string text = inlineString.InnerText;

            text = text.Replace("\n", " ");
            text = text.Replace("\"", "'");

            if (rowCount == 0)
            {
                // if header text is empty, do not write this column
                if (text == "")
                    columnsToSkip.Add(currColumnIndex);

                if (!headers.ContainsKey(text))
                    headers.Add(text, 1);
                else
                {
                    headers[text] = headers[text] + 1;
                    text = text + headers[text];
                }
            }

            if (!columnsToSkip.Contains(currColumnIndex))
            {
                lineBuilder.Append(string.Format("{0}{1}{0}", textDelimiter, text));

                // check end of row
                if (currColumnIndex != sheetEndColumnIndex)
                    lineBuilder.Append(columnDelimiter);
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
