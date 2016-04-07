using ClosedXML.Excel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using A = DocumentFormat.OpenXml.Drawing;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using A14 = DocumentFormat.OpenXml.Office2010.Drawing;
using System;
using System.Collections.Generic;
using System.Diagnostics.Contracts;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using xlsxfmt.format;
using YamlDotNet.Serialization;
using YamlDotNet.Serialization.NamingConventions;
using DocumentFormat.OpenXml.Validation;

namespace xlsxfmt
{
    public class XlsxFormatter
    {
        private string _sourceXlsx;
        private string _formatYaml;
        private string _outputXlsx;
        private Dictionary<string, string> _options = new Dictionary<string, string>();
        private Dictionary<int, string> _aggregateFunctions = new Dictionary<int, string>();
        private string _delimiter = "~";
        private string _calcModeInternal = "internal";
        private string _calcModeFormula = "formula";
        private YamlFile _yaml;

        public XlsxFormatter(string[] args)
        {
            ParseArguments(args);
        }

        private void ParseArguments(string[] args)
        {
            _sourceXlsx = args[0];
            _formatYaml = args[1];

            if (args.Length > 2)
            {
                for (int i = 2; i < args.Length; i++)
                {
                    if (args[i].StartsWith("--") == false)
                    {
                        _outputXlsx = args[i];
                        break;
                    }
                    else
                    {
                        foreach (Match m in Regex.Matches(args[i], @"(\w+(?:-\w+)+)=(\""?)(\w+)(\""?)"))
                            _options.Add(m.Groups[1].Value, m.Groups[2].Value);
                    }
                }
            }
        }

        private List<String> GetLogoSheets()
        {
            List<String> sheets = new List<string>();
            foreach (var sheet in _yaml.Sheet)
            {
                if (!String.IsNullOrEmpty(sheet.IncludeLogo) && sheet.IncludeLogo.Equals("true"))
                    sheets.Add(sheet.Name);
            }
            return sheets;
        }

        /// <summary>
        /// Inserts the image at the specified location 
        /// </summary>
        /// <param name="sheet1">The WorksheetPart where image to be inserted</param>
        /// <param name="startRowIndex">The starting Row Index</param>
        /// <param name="startColumnIndex">The starting column index</param>
        /// <param name="endRowIndex">The ending row index</param>
        /// <param name="endColumnIndex">The ending column index</param>
        /// <param name="imageStream">Stream which contains the image data</param>
        private static void InsertImage(WorksheetPart sheet1, int startRowIndex, int startColumnIndex, int endRowIndex, int endColumnIndex, Stream imageStream)
        {
            //Inserting a drawing element in worksheet
            //Make sure that the relationship id is same for drawing element in worksheet and its relationship part
            int drawingPartId = GetNextRelationShipID(sheet1);
            Drawing drawing1 = new Drawing() { Id = "rId" + drawingPartId.ToString() };

            //Check whether the WorksheetPart contains VmlDrawingParts (LegacyDrawing element)
            if (sheet1.VmlDrawingParts == null)
            {
                //if there is no VMLDrawing part (LegacyDrawing element) exists, just append the drawing part to the sheet
                //!!!sheet1.Worksheet.Append(drawing1);
                sheet1.Worksheet.InsertBefore(drawing1, sheet1.Worksheet.Last());
            }
            else
            {
                //if VmlDrawingPart (LegacyDrawing element) exists, then find the index of legacy drawing in the sheet and inserts the new drawing element before VMLDrawing part
                int legacyDrawingIndex = GetIndexofLegacyDrawing(sheet1);
                if (legacyDrawingIndex != -1)
                    sheet1.Worksheet.InsertAt<OpenXmlElement>(drawing1, legacyDrawingIndex);
                else
                    //!!1sheet1.Worksheet.Append(drawing1);
                    sheet1.Worksheet.InsertBefore(drawing1, sheet1.Worksheet.Last());
            }
            //Adding the drawings.xml part
            DrawingsPart drawingsPart1 = sheet1.AddNewPart<DrawingsPart>("rId" + drawingPartId.ToString());
            GenerateDrawingsPart1Content(drawingsPart1, startRowIndex, startColumnIndex, endRowIndex, endColumnIndex);
            //Adding the image
            ImagePart imagePart1 = drawingsPart1.AddNewPart<ImagePart>("image/jpeg", "rId1");
            imagePart1.FeedData(imageStream);
        }
        #region Helper methods
        /// <summary>
        /// Get the index of legacy drawing element in the specified WorksheetPart
        /// </summary>
        /// <param name="sheet1">The worksheetPart</param>
        /// <returns>Index of legacy drawing</returns>
        private static int GetIndexofLegacyDrawing(WorksheetPart sheet1)
        {
            for (int i = 0; i < sheet1.Worksheet.ChildElements.Count; i++)
            {
                OpenXmlElement element = sheet1.Worksheet.ChildElements[i];
                if (element is LegacyDrawing)
                    return i;
            }
            return -1;
        }
        /// <summary>
        /// Returns the WorksheetPart for the specified sheet name
        /// </summary>
        /// <param name="workbookpart">The WorkbookPart</param>
        /// <param name="sheetName">The name of the worksheet</param>
        /// <returns>Returns the WorksheetPart for the specified sheet name</returns>
        private static WorksheetPart GetSheetByName(WorkbookPart workbookpart, string sheetName)
        {
            var r = workbookpart.Workbook.Descendants<DocumentFormat.OpenXml.Spreadsheet.Sheet>();
            String id = r.First(s => s.Name == sheetName).Id;
            //foreach (WorksheetPart sheetPart in workbookpart.WorksheetParts)
            //{
            //    string uri = sheetPart.Uri.ToString();
            //    if (uri.EndsWith(sheetName + ".xml"))
            //        return sheetPart;
            //}
            if (!String.IsNullOrEmpty(id))
            {
                return workbookpart.GetPartById(id).CastTo<WorksheetPart>();
            }
            return null;
        }
        /// <summary>
        /// Returns the next relationship id for the specified WorksheetPart
        /// </summary>
        /// <param name="sheet1">The worksheetPart</param>
        /// <returns>Returns the next relationship id </returns>
        private static int GetNextRelationShipID(WorksheetPart sheet1)
        {
            int nextId = 0;
            List<int> ids = new List<int>();
            foreach (IdPartPair part in sheet1.Parts)
            {
                ids.Add(int.Parse(part.RelationshipId.Replace("rId", string.Empty)));
            }
            if (ids.Count > 0)
                nextId = ids.Max() + 1;
            else
                nextId = 1;
            return nextId;
        }

        // Generates content of drawingsPart1.
        private static void GenerateDrawingsPart1Content(DrawingsPart drawingsPart1, int startRowIndex, int startColumnIndex, int endRowIndex, int endColumnIndex)
        {
            Xdr.WorksheetDrawing worksheetDrawing1 = new Xdr.WorksheetDrawing();
            worksheetDrawing1.AddNamespaceDeclaration("xdr", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing");
            worksheetDrawing1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            Xdr.TwoCellAnchor twoCellAnchor1 = new Xdr.TwoCellAnchor() { EditAs = Xdr.EditAsValues.OneCell };

            Xdr.FromMarker fromMarker1 = new Xdr.FromMarker();
            Xdr.ColumnId columnId1 = new Xdr.ColumnId();
            columnId1.Text = startColumnIndex.ToString();
            Xdr.ColumnOffset columnOffset1 = new Xdr.ColumnOffset();
            columnOffset1.Text = "38100";
            Xdr.RowId rowId1 = new Xdr.RowId();
            rowId1.Text = startRowIndex.ToString();
            Xdr.RowOffset rowOffset1 = new Xdr.RowOffset();
            rowOffset1.Text = "0";

            fromMarker1.Append(columnId1);
            fromMarker1.Append(columnOffset1);
            fromMarker1.Append(rowId1);
            fromMarker1.Append(rowOffset1);

            Xdr.ToMarker toMarker1 = new Xdr.ToMarker();
            Xdr.ColumnId columnId2 = new Xdr.ColumnId();
            columnId2.Text = endColumnIndex.ToString();
            Xdr.ColumnOffset columnOffset2 = new Xdr.ColumnOffset();
            columnOffset2.Text = "542925";
            Xdr.RowId rowId2 = new Xdr.RowId();
            rowId2.Text = endRowIndex.ToString();
            Xdr.RowOffset rowOffset2 = new Xdr.RowOffset();
            rowOffset2.Text = "161925";

            toMarker1.Append(columnId2);
            toMarker1.Append(columnOffset2);
            toMarker1.Append(rowId2);
            toMarker1.Append(rowOffset2);

            Xdr.Picture picture1 = new Xdr.Picture();

            Xdr.NonVisualPictureProperties nonVisualPictureProperties1 = new Xdr.NonVisualPictureProperties();
            Xdr.NonVisualDrawingProperties nonVisualDrawingProperties1 = new Xdr.NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Picture 1" };

            Xdr.NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties1 = new Xdr.NonVisualPictureDrawingProperties();
            A.PictureLocks pictureLocks1 = new A.PictureLocks() { NoChangeAspect = true };

            nonVisualPictureDrawingProperties1.Append(pictureLocks1);

            nonVisualPictureProperties1.Append(nonVisualDrawingProperties1);
            nonVisualPictureProperties1.Append(nonVisualPictureDrawingProperties1);

            Xdr.BlipFill blipFill1 = new Xdr.BlipFill();

            A.Blip blip1 = new A.Blip() { Embed = "rId1", CompressionState = A.BlipCompressionValues.Print };
            blip1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

            A.BlipExtensionList blipExtensionList1 = new A.BlipExtensionList();

            A.BlipExtension blipExtension1 = new A.BlipExtension() { Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}" };

            A14.UseLocalDpi useLocalDpi1 = new A14.UseLocalDpi() { Val = false };
            useLocalDpi1.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");

            blipExtension1.Append(useLocalDpi1);

            blipExtensionList1.Append(blipExtension1);

            blip1.Append(blipExtensionList1);

            A.Stretch stretch1 = new A.Stretch();
            A.FillRectangle fillRectangle1 = new A.FillRectangle();

            stretch1.Append(fillRectangle1);

            blipFill1.Append(blip1);
            blipFill1.Append(stretch1);

            Xdr.ShapeProperties shapeProperties1 = new Xdr.ShapeProperties();

            A.Transform2D transform2D1 = new A.Transform2D();
            A.Offset offset1 = new A.Offset() { X = 1257300L, Y = 762000L };
            A.Extents extents1 = new A.Extents() { Cx = 2943225L, Cy = 2257425L };

            transform2D1.Append(offset1);
            transform2D1.Append(extents1);

            A.PresetGeometry presetGeometry1 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList1 = new A.AdjustValueList();

            presetGeometry1.Append(adjustValueList1);

            shapeProperties1.Append(transform2D1);
            shapeProperties1.Append(presetGeometry1);

            picture1.Append(nonVisualPictureProperties1);
            picture1.Append(blipFill1);
            picture1.Append(shapeProperties1);
            Xdr.ClientData clientData1 = new Xdr.ClientData();

            twoCellAnchor1.Append(fromMarker1);
            twoCellAnchor1.Append(toMarker1);
            twoCellAnchor1.Append(picture1);
            twoCellAnchor1.Append(clientData1);

            worksheetDrawing1.Append(twoCellAnchor1);

            drawingsPart1.WorksheetDrawing = worksheetDrawing1;
        }

        private void validateXlsx(String xlsxFile)
        {
            try
            {
                OpenXmlValidator validator = new OpenXmlValidator();
                int count = 0;
                foreach (
                    ValidationErrorInfo error in validator.Validate(
                        SpreadsheetDocument.Open(xlsxFile, true))
                 )
                {
                    count++;
                    Console.WriteLine("Error " + count);
                    Console.WriteLine("Description: " + error.Description);
                    Console.WriteLine("Path: " + error.Path.XPath);
                    Console.WriteLine("Part: " + error.Part.Uri);
                    Console.WriteLine("-------------------------------------------");
                }
                Console.ReadKey();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {

            }

        }
        #endregion Helper methods

        private Stream loadLogo(String logoPath)
        {
            Stream logoStream = null;
            try
            {
                Uri uri = new Uri(logoPath);
                logoStream = new FileStream(uri.AbsolutePath, FileMode.Open);
            }
            catch (Exception ex)
            {

            }
            return logoStream;
        }

        public void Process()
        {
            // Read input parameters
            var formatReader = new StreamReader(_formatYaml);
            var formatDeserializer = new Deserializer(namingConvention: new CamelCaseNamingConvention(), ignoreUnmatched: true);
            _yaml = formatDeserializer.Deserialize<YamlFile>(formatReader);


            // Construct output filename
            ConstructOutputFilename();

            // Read source workbook
            var sourceWorkbook = new XLWorkbook(_sourceXlsx);

            // Create empty output workbook
            var outputWorkbook = new XLWorkbook();

            // Initialize aggregation functions
            InitAggregateFunctions();

            // Get logo sheets
            List<String> logoSheets = GetLogoSheets();

            Stream logoStream = loadLogo(_yaml.Format.LogoPath);

            bool needLogoUsage = (logoSheets.Count > 0) && (logoStream != null);

            // Construct output workbook using source workbook and input params
            Construct(sourceWorkbook, outputWorkbook, needLogoUsage);

            // Save output file
            outputWorkbook.SaveAs(_outputXlsx);


            if (needLogoUsage)
            {
                using (SpreadsheetDocument document = SpreadsheetDocument.Open(_outputXlsx, true))
                {
                    WorkbookPart workbookpart = document.WorkbookPart;
                    //WorksheetPart sheet1 = workbookpart.WorksheetParts.First();
                    foreach (var customSheet in logoSheets)
                    {
                        WorksheetPart sheet1 = GetSheetByName(workbookpart, customSheet);

                        Row row = sheet1.Worksheet.GetFirstChild<SheetData>().Elements<Row>().FirstOrDefault();
                        var numberOfColumns = 0;
                        if (row != null)
                        {
                            var spans = row.Spans != null ? row.Spans.InnerText : "";
                            if (spans != String.Empty)
                            {
                                char[] delimiter = new char[1];
                                delimiter[0] = ':';
                                string[] columns = spans.Split(delimiter);
                                numberOfColumns = int.Parse(columns[1]);
                            }
                        }

                        //insert Image by specifying two range

                        InsertImage(sheet1, 0, 0, 1, numberOfColumns, logoStream);
                        //!!!panes
                        /*var panes = sheet1.Worksheet.Descendants<Pane>();
                        foreach (var item in panes)
                        {
                            item.TopLeftCell = "L9";
                        }*/
                    }
                    document.WorkbookPart.Workbook.Save();
                    // Close the document handle.
                    document.Close();
                }
                //   validateXlsx(_outputXlsx);
            }
        }


        private void ConstructOutputFilename()
        {
            if (string.IsNullOrEmpty(_outputXlsx))
            {
                _outputXlsx = _yaml.Format.OutputFilenameBase;

                var fileName = Path.GetFileName(_outputXlsx);

                if (_options.ContainsKey(@"output-filename-prefix"))
                    fileName = _options[@"output-filename-prefix"] + fileName;

                if (_options.ContainsKey(@"output-filename-postfix"))
                    fileName = fileName + _options[@"output-filename-postfix"];

                var dir = Path.GetPathRoot(_outputXlsx);
                if (!string.IsNullOrEmpty(dir))
                    _outputXlsx = dir + "\\" + fileName + @".xlsx";
                else
                    _outputXlsx = fileName + @".xlsx";
            }
        }

        private void Construct(XLWorkbook wsrc, XLWorkbook wout, bool needLogoUsage)
        {
            // Construct sheets
            foreach (var shtFmt in _yaml.Sheet)
            {
               // if (shtFmt.Name.IndexOf("Location") >= 0)
                //{
                    var source = shtFmt.Name;
                    if (!string.IsNullOrEmpty(shtFmt.Source))
                        source = shtFmt.Source;

                    // Find source sheet in source workbook
                    var ssht = wsrc.Worksheets.Where(x => x.Name == source).FirstOrDefault();
                    ConstructSheet(ssht, wout, shtFmt, needLogoUsage);
               // }
            }
        }

        private void InitAggregateFunctions()
        {
            _aggregateFunctions.Add(1, "AVERAGE");
            _aggregateFunctions.Add(2, "COUNT");
            _aggregateFunctions.Add(3, "COUNTA");
            _aggregateFunctions.Add(4, "MAX");
            _aggregateFunctions.Add(5, "MIN");
            _aggregateFunctions.Add(6, "PRODUCT");
            _aggregateFunctions.Add(7, "STDEV");
            _aggregateFunctions.Add(8, "STDEVP");
            _aggregateFunctions.Add(9, "SUM");
            _aggregateFunctions.Add(10, "VAR");
            _aggregateFunctions.Add(11, "VARP");
        }

        private int GetFreezeRow(String cell)
        {
            string col = cell;
            int startIndex = col.IndexOfAny("0123456789".ToCharArray());
            int row = Int32.Parse(col.Substring(startIndex));
            return row - 1;
        }

        private int GetFreezeCol(String cell)
        {
            string col = cell;
            int startIndex = col.IndexOfAny("0123456789".ToCharArray());
            String columnName = col.Substring(0, startIndex);
            columnName = columnName.ToUpperInvariant();

            int sum = 0;

            for (int i = 0; i < columnName.Length; i++)
            {
                sum *= 26;
                sum += (columnName[i] - 'A' + 1);
            }

            return sum - 1;
        }

        private List<xlsxfmt.format.Column> GetSubtotalCols(xlsxfmt.format.Sheet sheet)
        {
            List<xlsxfmt.format.Column> lc = new List<xlsxfmt.format.Column>();
            for (int i = 0; i < sheet.Column.Count; i++)
            {
                if (sheet.Column[i].Subtotal != null && sheet.Column[i].Subtotal.Group == "true")
                {
                    lc.Add(sheet.Column[i]);
                    lc[lc.Count - 1].Number = i + 1; // Columns should have numbers from 1 to 16384
                }

            }
            return lc;
        }

        private Dictionary<int, int> GetColFuncs(xlsxfmt.format.Sheet sheet)
        {
            Dictionary<int, int> lc = new Dictionary<int, int>();
            for (int i = 0; i < sheet.Column.Count; i++)
            {
                if (sheet.Column[i].Subtotal != null && !String.IsNullOrEmpty(sheet.Column[i].Subtotal.Function))
                {
                    lc[i + 1] = int.Parse(sheet.Column[i].Subtotal.Function);
                }

            }
            return lc;
        }

        private int GetColumnNumber(xlsxfmt.format.Sheet sheet, String columnName)
        {
            int t = sheet.Column.FindIndex(x => x.Name == columnName);
            return t + 1;
        }

        private String GetKeyValue(IXLRow row, List<xlsxfmt.format.Column> lc)
        {
            String result = "";
            foreach (var item in lc)
            {
                if (String.IsNullOrEmpty(result))
                    result = row.Worksheet.Cell(row.RowNumber(), item.Number).Value.ToString();
                else
                    result = String.Concat(result, _delimiter, row.Worksheet.Cell(row.RowNumber(), item.Number).Value.ToString());
            }
            return result;
        }

        private List<int> GetTotalLevels(IXLRow prevRow, IXLRow curRow, List<xlsxfmt.format.Column> lc)
        {
            List<int> levels = new List<int>();
            int minLevel = 0;
            foreach (int colNum in (from a in lc
                                    orderby a.Number
                                    select a.Number
                                    )
                    )
            {
                if (!prevRow.Cell(colNum).Value.ToString().ToUpper().Equals(curRow.Cell(colNum).Value.ToString().ToUpper()))
                {
                    break;
                }
                else
                    minLevel++;
            }
            for (int i = lc.Count; i > minLevel; i--)
            {
                levels.Add(i - 1);
            }
            return levels;
        }

        private List<SortKeyLevel> GetTotalKeys(IXLRow row, List<int> totalLevels, int groupLevels, List<xlsxfmt.format.Column> lc)
        {
            List<SortKeyLevel> keys = new List<SortKeyLevel>();
            String prefix = "";
            if (totalLevels.Count == 0)
                return keys;
            int minLevel = totalLevels.Min();
            for (int i = 0; i <= minLevel; i++)
            {
                if (!String.IsNullOrEmpty(prefix))
                    prefix = String.Concat(prefix, _delimiter);

                prefix = String.Concat(prefix, row.Cell(lc[i].Number).Value);

            }
            foreach (int level in (from l in totalLevels orderby l select l))
            {
                if (minLevel < level)
                {
                    for (int i = minLevel + 1; i <= level; i++)
                    {
                        if (!String.IsNullOrEmpty(prefix))
                            prefix = String.Concat(prefix, _delimiter);

                        prefix = String.Concat(prefix, row.Cell(lc[i].Number).Value);

                    }
                }

                keys.Add(new SortKeyLevel(prefix, level));
                minLevel = level;

            }
            return keys;
        }

        private void WriteToTemp(IXLWorksheet sheet, int level, IXLRange range)
        {
            foreach (var item in range.CellsUsed())
            {
                sheet.Column(level).LastCellUsed().CellBelow().Value = item.Value;
            }
        }

        private Double EvaluateTemp(IXLWorksheet sheet, int level, String func)
        {
            IXLCell c = sheet.Column(level).FirstCellUsed();
            IXLCell c1 = sheet.Column(level).LastCellUsed();
            if (c == null || c1 == null) return 0;
            IXLRange cells = sheet.Range(sheet.Column(level).FirstCellUsed(), sheet.Column(level).LastCellUsed());
            return (Double)sheet.Evaluate(String.Concat(func, "(", cells.RangeAddress, ")"));
        }


        private void SetTotalRowStyle(int totalLevel, IXLRow row, List<xlsxfmt.format.Column> lc, Dictionary<int, int> colFuncs, xlsxfmt.format.Sheet sheet)
        {
            xlsxfmt.format.Column c = lc.ElementAtOrDefault(totalLevel);

            String colorStr = "";
            if (c.Subtotal != null && c.Subtotal.TotalRowBgcolor != null && !string.IsNullOrEmpty(c.Subtotal.TotalRowBgcolor))
                colorStr = c.Subtotal.TotalRowBgcolor;
            if (!String.IsNullOrEmpty(colorStr))
            {
                int headerR = int.Parse(colorStr.Substring(0, 2), System.Globalization.NumberStyles.HexNumber);
                int headerG = int.Parse(colorStr.Substring(2, 2), System.Globalization.NumberStyles.HexNumber);
                int headerB = int.Parse(colorStr.Substring(4, 2), System.Globalization.NumberStyles.HexNumber);
                row.Style.Fill.SetBackgroundColor(XLColor.FromArgb(headerR, headerG, headerB));
            }
            xlsxfmt.format.Font cellFont;
            if (c.Font != null && c.Font.Data != null)
                cellFont = c.Font;
            else
                cellFont = sheet.Font;
            if (cellFont != null && cellFont.Data != null)
            {
                //size
                if (!String.IsNullOrEmpty(cellFont.Data.Size))
                {
                    double fontSz;
                    Double.TryParse(cellFont.Data.Size, out fontSz);
                    row.Cell(totalLevel + 1).Style.Font.SetFontSize(fontSz);
                }
                //style
                String dataStyle;
                if (!string.IsNullOrEmpty(cellFont.Data.Style))
                {
                    dataStyle = cellFont.Data.Style;
                }
                else
                {
                    dataStyle = _yaml.Defaults.Font.Footer.Style;
                }
                if (!String.IsNullOrEmpty(dataStyle))
                {
                    if (dataStyle == "bold")
                    {
                        row.Cell(totalLevel + 1).Style.Font.SetBold();
                    }
                    else if (dataStyle == "italic")
                    {
                        row.Cell(totalLevel + 1).Style.Font.SetItalic();
                    }
                    else if (dataStyle == "underline")
                    {
                        row.Cell(totalLevel + 1).Style.Font.SetUnderline();
                    }
                }
                //conditional-formatting
            }
            foreach (var item in colFuncs)
            {
                row.Cell(item.Key).SetDataType(XLCellValues.Number);
                int numDecimalPlaces;
                xlsxfmt.format.Column gc = sheet.Column[item.Key - 1];
                int.TryParse(gc.DecimalPlaces, out numDecimalPlaces);
                if (numDecimalPlaces == 0)
                    int.TryParse(_yaml.Defaults.Column.DecimalPlaces, out numDecimalPlaces);
                if (numDecimalPlaces == 0) numDecimalPlaces = 2;
                string s = "0.";
                s = s.PadRight(numDecimalPlaces + 2, '0');
                row.Cell(item.Key).Style.NumberFormat.Format = String.Concat("_-[$$-C09]* #,##", s, "_-;\\-[$$-C09]* #,##", s, "_-;_-[$$-C09]* \"-\"??_-;_-@_-");
            }

        }

        private void SetGrandTotalRowStyle(int totalLevel, IXLRow row, List<xlsxfmt.format.Column> lc, Dictionary<int, int> colFuncs, xlsxfmt.format.Sheet sheet)
        {
            xlsxfmt.format.Column c = lc.ElementAtOrDefault(totalLevel);

            String colorStr = sheet.GrandTotalRowBgcolor;
            if (String.IsNullOrEmpty(colorStr))
                colorStr = _yaml.Defaults.Sheet.GrandTotalRowBgcolor;
            if (!String.IsNullOrEmpty(colorStr))
            {
                int headerR = int.Parse(colorStr.Substring(0, 2), System.Globalization.NumberStyles.HexNumber);
                int headerG = int.Parse(colorStr.Substring(2, 2), System.Globalization.NumberStyles.HexNumber);
                int headerB = int.Parse(colorStr.Substring(4, 2), System.Globalization.NumberStyles.HexNumber);
                row.Style.Fill.SetBackgroundColor(XLColor.FromArgb(headerR, headerG, headerB));
            }
            xlsxfmt.format.Font cellFont;
            if (c.Font != null && c.Font.Data != null)
                cellFont = c.Font;
            else
                cellFont = sheet.Font;
            if (cellFont != null && cellFont.Data != null)
            {
                //size
                if (!String.IsNullOrEmpty(cellFont.Data.Size))
                {
                    double fontSz;
                    Double.TryParse(cellFont.Data.Size, out fontSz);
                    row.Cell(totalLevel + 1).Style.Font.SetFontSize(fontSz);
                }
                //style
                String dataStyle;
                if (!string.IsNullOrEmpty(cellFont.Data.Style))
                {
                    dataStyle = cellFont.Data.Style;
                }
                else
                {
                    dataStyle = _yaml.Defaults.Font.Footer.Style;
                }
                if (!String.IsNullOrEmpty(dataStyle))
                {
                    if (dataStyle == "bold")
                    {
                        row.Cell(totalLevel + 1).Style.Font.SetBold();
                    }
                    else if (dataStyle == "italic")
                    {
                        row.Cell(totalLevel + 1).Style.Font.SetItalic();
                    }
                    else if (dataStyle == "underline")
                    {
                        row.Cell(totalLevel + 1).Style.Font.SetUnderline();
                    }
                }
                //conditional-formatting
            }
            foreach (var item in colFuncs)
            {
                row.Cell(item.Key).SetDataType(XLCellValues.Number);
                int numDecimalPlaces;
                xlsxfmt.format.Column gc = sheet.Column[item.Key - 1];
                int.TryParse(gc.DecimalPlaces, out numDecimalPlaces);
                if (numDecimalPlaces == 0)
                    int.TryParse(_yaml.Defaults.Column.DecimalPlaces, out numDecimalPlaces);
                if (numDecimalPlaces == 0) numDecimalPlaces = 2;
                string s = "0.";
                s = s.PadRight(numDecimalPlaces + 2, '0');
                row.Cell(item.Key).Style.NumberFormat.Format = String.Concat("_-[$$-C09]* #,##", s, "_-;\\-[$$-C09]* #,##", s, "_-;_-[$$-C09]* \"-\"??_-;_-@_-");
            }

        }
        private double VariancePopulation(List<Double> source)
        {
            // Excel VAR.P function
            // https://support.office.com/en-us/article/VAR-P-function-73d1285c-108c-4843-ba5d-a51f90656f3a
            double mean = source.Average();
            return source.Sum(key => (key - mean) * (key - mean)) / source.Count;
        }

        private double Variance(List<Double> source)
        {
            // Excel VAR.P function
            // https://support.office.com/en-us/article/VAR-P-function-73d1285c-108c-4843-ba5d-a51f90656f3a
            double mean = source.Average();
            return source.Sum(key => (key - mean) * (key - mean)) / (source.Count - 1);
        }

        private double StandardDeviation(List<Double> source)
        {
            // Excel VAR.P function
            // https://support.office.com/en-us/article/STDEV-function-51fecaaa-231e-4bbb-9230-33650a72c9b0
            Double t = Variance(source);
            if (t != 0)
                return Math.Sqrt(t);
            else
                return 0;
        }
        private double StandardDeviationPopulation(List<Double> source)
        {
            // Excel VAR.P function
            // https://support.office.com/en-us/article/STDEV-function-51fecaaa-231e-4bbb-9230-33650a72c9b0
            Double t = VariancePopulation(source);
            if (t != 0)
                return Math.Sqrt(t);
            else
                return 0;
        }

        private Double GetGroupResult(List<Double> elements, int funcId)
        {
            if (_aggregateFunctions.ContainsKey(funcId))
            {
                String function = _aggregateFunctions[funcId];
                if (function.Equals("SUM"))
                {
                    return elements.Sum();
                }
                else if (function.Equals("MAX"))
                {
                    return elements.Max();
                }
                else if (function.Equals("MIN"))
                {
                    return elements.Min();
                }
                else if (function.Equals("AVERAGE"))
                {
                    return elements.Average();
                }
                else if (function.Equals("COUNT"))
                {
                    return elements.Count;
                }
                else if (function.Equals("VAR"))
                {
                    return Variance(elements);
                }
                else if (function.Equals("VARP"))
                {
                    return VariancePopulation(elements);
                }
                else if (function.Equals("STDEV"))
                {
                    return StandardDeviation(elements);
                }
                else if (function.Equals("STDEVP"))
                {
                    return StandardDeviationPopulation(elements);
                }
                else if (function.Equals("PRODUCT"))
                {
                    return elements.Aggregate(1d, (p, d) => p * d);
                }
            }

            return 0;
        }

        private String GetCalcMode(int columnNumber, xlsxfmt.format.Sheet sheet)
        {
            String calcMode = _calcModeInternal;
            try
            {
                xlsxfmt.format.Column c = sheet.Column[columnNumber - 1];
                if (String.IsNullOrEmpty(c.calculationMode))
                {
                    if (String.IsNullOrEmpty(_yaml.Defaults.Sheet.TotalsCalculationMode))
                    {
                        calcMode = _calcModeInternal;
                    }
                    else
                    {
                        calcMode = _yaml.Defaults.Sheet.TotalsCalculationMode;
                    }
                }
                else
                {
                    calcMode = c.calculationMode;
                }
            }
            catch (Exception ex)
            {
                calcMode = _calcModeInternal;
            }
            return calcMode;
        }

        private void ConstructSheet(IXLWorksheet ssht, XLWorkbook wout, xlsxfmt.format.Sheet shtFmt, bool needLogoUsage)
        {
            IXLWorksheet wsht = wout.AddWorksheet(shtFmt.Name);
            int logoRows = 0;
            int headerRows = 1;
            int startRowNum = 1;
            if (needLogoUsage && !String.IsNullOrEmpty(shtFmt.IncludeLogo) && shtFmt.IncludeLogo.Equals("true"))
            {
                wsht.Row(1).Height = 140;//!!!
                logoRows = 2;
            }
            else
            {
                logoRows = 0;
            }
            startRowNum += logoRows;
            int rowNum = startRowNum;
            int colNum = 1;
            #region Freeze on cell
            String freezeCell;
            if (!string.IsNullOrEmpty(shtFmt.FreezeOnCell))
                freezeCell = shtFmt.FreezeOnCell;
            else
                freezeCell = _yaml.Defaults.Sheet.FreezeOnCell;
            if (!string.IsNullOrEmpty(freezeCell))
                wsht.SheetView.Freeze(GetFreezeRow(freezeCell), GetFreezeCol(freezeCell));
            #endregion
            #region header row bg color
            //header row bg color
            String headerColorStr;
            if (!string.IsNullOrEmpty(shtFmt.HeaderRowBgcolor))
                headerColorStr = shtFmt.HeaderRowBgcolor;
            else
                headerColorStr = _yaml.Defaults.Sheet.HeaderRowBgcolor;
            if (!String.IsNullOrEmpty(headerColorStr))
            {
                int headerR = int.Parse(headerColorStr.Substring(0, 2), System.Globalization.NumberStyles.HexNumber);
                int headerG = int.Parse(headerColorStr.Substring(2, 2), System.Globalization.NumberStyles.HexNumber);
                int headerB = int.Parse(headerColorStr.Substring(4, 2), System.Globalization.NumberStyles.HexNumber);
                wsht.Row(rowNum).Style.Fill.SetBackgroundColor(XLColor.FromArgb(headerR, headerG, headerB));
            }
            #endregion
            //#region Grand total row bg color
            //String grandTotalColorStr;
            //if (!string.IsNullOrEmpty(shtFmt.GrandTotalRowBgcolor))
            //    grandTotalColorStr = shtFmt.GrandTotalRowBgcolor;
            //else
            //    grandTotalColorStr = _yaml.Defaults.Sheet.GrandTotalRowBgcolor;
            //if (!String.IsNullOrEmpty(grandTotalColorStr))
            //{
            //    int headerR = int.Parse(grandTotalColorStr.Substring(0, 2), System.Globalization.NumberStyles.HexNumber);
            //    int headerG = int.Parse(grandTotalColorStr.Substring(2, 2), System.Globalization.NumberStyles.HexNumber);
            //    int headerB = int.Parse(grandTotalColorStr.Substring(4, 2), System.Globalization.NumberStyles.HexNumber);
            //    int grandRowNum = 10;
            //    wsht.Row(grandRowNum).Style.Fill.SetBackgroundColor(XLColor.FromArgb(headerR, headerG, headerB));
            //}
            //#endregion
            #region Hidden
            if (Convert.ToBoolean(shtFmt.Hidden))
                wsht.Hide();
            #endregion
            #region fontname
            String fontName;
            if (shtFmt.Font != null && !String.IsNullOrEmpty(shtFmt.Font.Family))
                fontName = shtFmt.Font.Family;
            else
                fontName = _yaml.Defaults.Font.Family;
            if (!String.IsNullOrEmpty(fontName))
                wsht.Style.Font.SetFontName(fontName);
            #endregion
            #region fontsize
            String fontSize;
            if (shtFmt.Font != null && !String.IsNullOrEmpty(shtFmt.Font.Size))
                fontSize = shtFmt.Font.Size;
            else
                fontSize = _yaml.Defaults.Font.Size;
            if (!String.IsNullOrEmpty(fontSize))
            {
                double fontSz;
                Double.TryParse(fontSize, out fontSz);
                if (fontSz != 0)
                    wsht.Style.Font.SetFontSize(fontSz);
            }
            #endregion
            #region fontStyle
            String fontStyle;
            if (shtFmt.Font != null && !String.IsNullOrEmpty(shtFmt.Font.Style))
            {
                fontStyle = shtFmt.Font.Style;
            }
            else
                fontStyle = _yaml.Defaults.Font.Style;
            if (!String.IsNullOrEmpty(fontStyle))
            {
                if (fontStyle == "bold")
                {
                    wsht.Style.Font.SetBold();
                }
                else if (fontStyle == "italic")
                {
                    wsht.Style.Font.SetItalic();
                }
                else if (fontStyle == "underline")
                {
                    wsht.Style.Font.SetUnderline();
                }
            }
            #endregion
            foreach (var colFmt in shtFmt.Column)
            {
                var source = colFmt.Name;
                if (!string.IsNullOrEmpty(colFmt.Source))
                    source = colFmt.Source;

                // Find corresponding column in source sheet
                var srcColumn = ssht.Columns().Where(x => x.Cell(1).Value.ToString()
                    == source).FirstOrDefault();

                // Set output column name
                wsht.Cell(rowNum, colNum).Value = colFmt.Name;
                #region setheaderstyle
                String headerStyle = "";
                if (shtFmt.Font != null && shtFmt.Font.Header != null)
                {
                    if (!string.IsNullOrEmpty(shtFmt.Font.Header.Style))
                    {
                        headerStyle = shtFmt.Font.Header.Style;
                    }
                    else
                    {
                        headerStyle = _yaml.Defaults.Font.Header.Style;
                    }

                }
                else if (_yaml.Defaults.Font != null && _yaml.Defaults.Font.Header != null)
                {
                    headerStyle = _yaml.Defaults.Font.Header.Style;
                }
                if (!String.IsNullOrEmpty(headerStyle))
                {
                    if (headerStyle == "bold")
                    {
                        wsht.Cell(rowNum, colNum).Style.Font.SetBold();
                    }
                    else if (headerStyle == "italic")
                    {
                        wsht.Cell(rowNum, colNum).Style.Font.SetItalic();
                    }
                    else if (headerStyle == "underline")
                    {
                        wsht.Cell(rowNum, colNum).Style.Font.SetUnderline();
                    }
                }
                #endregion
                rowNum++;
                // Populate output column cells
                var cellCnt = srcColumn.Cells().Count();
                for (int i = 2; i < cellCnt; i++)
                {
                    wsht.Cell(rowNum, colNum).Value = srcColumn.Cell(i).Value;
                    #region setdatacellstyle
                    xlsxfmt.format.Font cellFont;
                    if (colFmt.Font != null && colFmt.Font.Data != null)
                        cellFont = colFmt.Font;
                    else
                        cellFont = shtFmt.Font;
                    if (cellFont != null && cellFont.Data != null)
                    {
                        //size
                        if (!String.IsNullOrEmpty(cellFont.Data.Size))
                        {
                            double fontSz;
                            Double.TryParse(cellFont.Data.Size, out fontSz);
                            wsht.Cell(rowNum, colNum).Style.Font.SetFontSize(fontSz);
                        }
                        //style
                        String dataStyle = "";
                        if (!string.IsNullOrEmpty(cellFont.Data.Style))
                        {
                            dataStyle = cellFont.Data.Style;
                        }
                        else
                        {
                            if (_yaml.Defaults.Font.Data != null && _yaml.Defaults.Font.Data.Style != null)
                                dataStyle = _yaml.Defaults.Font.Data.Style;
                        }
                        if (!String.IsNullOrEmpty(dataStyle))
                        {
                            if (dataStyle == "bold")
                            {
                                wsht.Cell(rowNum, colNum).Style.Font.SetBold();
                            }
                            else if (dataStyle == "italic")
                            {
                                wsht.Cell(rowNum, colNum).Style.Font.SetItalic();
                            }
                            else if (dataStyle == "underline")
                            {
                                wsht.Cell(rowNum, colNum).Style.Font.SetUnderline();
                            }
                        }
                        //conditional-formatting

                    }
                    #endregion
                    rowNum++;
                }
                colNum++;
                rowNum = startRowNum;
            }

            colNum = 1;
            foreach (xlsxfmt.format.Column colFmt in shtFmt.Column)
            {
                #region column format
                if (string.IsNullOrEmpty(colFmt.FormatType))
                    colFmt.FormatType = "GENERAL";
                var rng = wsht.Range(wsht.Column(colNum).FirstCellUsed().CellBelow(), wsht.Column(colNum).CellsUsed().Last());
                if (colFmt.FormatType.ToUpper() == "NUMBER")
                {
                    rng.SetDataType(XLCellValues.Number);
                    int numDecimalPlaces;
                    int.TryParse(colFmt.DecimalPlaces, out numDecimalPlaces);
                    if (numDecimalPlaces == 0)
                        int.TryParse(_yaml.Defaults.Column.DecimalPlaces, out numDecimalPlaces);
                    if (numDecimalPlaces != 0)
                    {
                        string s = "0.";
                        s = s.PadRight(numDecimalPlaces + 2, '0');
                        wsht.Column(colNum).CellsUsed().Style.NumberFormat.Format = s;
                    }
                }
                else if (colFmt.FormatType.ToUpper() == "TEXT")
                {
                    rng.SetDataType(XLCellValues.Text);
                    // wsht.Column(colNum).CellsUsed().SetDataType(XLCellValues.Text);
                }
                else if (colFmt.FormatType.ToUpper() == "DATE")
                {
                    rng.SetDataType(XLCellValues.DateTime);
                    //wsht.Column(colNum).CellsUsed().SetDataType(XLCellValues.DateTime);
                    String dateFormat;
                    if (!string.IsNullOrEmpty(colFmt.DateFormat))
                    {
                        dateFormat = colFmt.DateFormat;
                    }
                    else
                        dateFormat = _yaml.Defaults.Column.DateFormat;
                    if (!string.IsNullOrEmpty(colFmt.DateFormat))
                        wsht.Column(colNum).CellsUsed().Style.DateFormat.Format = dateFormat;
                }
                else if (colFmt.FormatType.ToUpper() == "ACCOUNTING")
                {
                    rng.SetDataType(XLCellValues.Number);
                    int numDecimalPlaces;
                    int.TryParse(colFmt.DecimalPlaces, out numDecimalPlaces);
                    if (numDecimalPlaces == 0)
                        int.TryParse(_yaml.Defaults.Column.DecimalPlaces, out numDecimalPlaces);
                    if (numDecimalPlaces == 0) numDecimalPlaces = 2;
                    string s = "0.";
                    s = s.PadRight(numDecimalPlaces + 2, '0');
                    rng.Style.NumberFormat.Format = String.Concat("_-[$$-C09]* #,##", s, "_-;\\-[$$-C09]* #,##", s, "_-;_-[$$-C09]* \"-\"??_-;_-@_-");
                }
                if (colFmt.conditionalFormatting != null && colFmt.conditionalFormatting.Type == "databar")
                {
                    //gradient-green, gradient-red, gradient-orange, gradient-ltblue, gradient-purple
                    if (colFmt.conditionalFormatting.Style == "gradient-ltblue")
                    {
                        rng.AddConditionalFormat().DataBar(XLColor.Blue).LowestValue().HighestValue();
                    }
                    else if (colFmt.conditionalFormatting.Style == "gradient-green")
                    {
                        rng.AddConditionalFormat().DataBar(XLColor.Green).LowestValue().HighestValue(); ;
                    }
                    else if (colFmt.conditionalFormatting.Style == "gradient-orange")
                    {
                        rng.AddConditionalFormat().DataBar(XLColor.Orange).LowestValue().HighestValue(); ;
                    }
                    else if (colFmt.conditionalFormatting.Style == "gradient-purple")
                    {
                        rng.AddConditionalFormat().DataBar(XLColor.Purple).LowestValue().HighestValue(); ;
                    }
                    else
                        if (colFmt.conditionalFormatting.Style == "gradient-red")
                        {
                            rng.AddConditionalFormat().DataBar(XLColor.Red).LowestValue().HighestValue(); ;
                        }
                }
                #endregion
                colNum++;
            }

            List<xlsxfmt.format.Column> lc = GetSubtotalCols(shtFmt);
            Dictionary<int, int> colFuncs = GetColFuncs(shtFmt);
            String key;
            String prevKey = "";
            IXLRow curRow, prevRow;
            int groupLevel = lc.Count;
            int lastColUsed = wsht.LastColumnUsed().ColumnNumber();
            int lastRowUsed = wsht.LastRowUsed().RowNumber();
            int sortColNumber = lastColUsed + 1;
            var tableRange = wsht.Range(wsht.FirstCellUsed().CellBelow(), wsht.LastColumnUsed().LastCellUsed());
            if (groupLevel > 0)
            {
                int rowCnt = tableRange.RowCount();
                List<SortKeyLevel> totalKeys = new List<SortKeyLevel>();
                List<int> totalLevels = new List<int>();
                prevRow = null;
                curRow = tableRange.FirstRowUsed().WorksheetRow();
                //getting keys and row levels
                //wsht.Column(sortColNumber).SetDataType(XLCellValues.Number);
                for (int i = 0; i <= tableRange.RowCount(); i++)
                {
                    key = GetKeyValue(curRow, lc);
                    wsht.Cell(i + logoRows + headerRows, sortColNumber).Value = 0;
                    if (!String.IsNullOrEmpty(prevKey) && key != prevKey)
                    {
                        var lvls = GetTotalLevels(prevRow, curRow, lc);
                        totalLevels.AddRange(lvls);
                        totalKeys.AddRange(GetTotalKeys(prevRow, lvls, groupLevel, lc));

                    }

                    prevKey = key;
                    prevRow = curRow;
                    curRow = curRow.RowBelow();

                }
                // Adding lines to the end
                IXLRow lastR = wsht.LastRowUsed();
                var sortKeys = totalKeys.GroupBy(s => new { s.grouplevel, s.key }).Select(x => new SortKeyLevel(x.Key.key, x.Key.grouplevel)).ToList();
                for (int i = 0; i < sortKeys.Count(); i++)
                {
                    //lastR.InsertRowsBelow(1);
                    wsht.Cell(lastRowUsed + i + 1, sortColNumber).Value = groupLevel - totalKeys[i].grouplevel;
                    string[] keys = sortKeys[i].key.Split(new string[] { _delimiter }, StringSplitOptions.None);
                    for (int j = 0; j < keys.Length; j++)
                    {
                        if (!String.IsNullOrEmpty(keys[j]))
                            wsht.Cell(lastRowUsed + i + 1, lc[j].Number).Value = keys[j];
                    }
                    lastR = lastR.RowBelow();
                }
            }

            //Sorting

            var tableSortRange = wsht.Range(wsht.FirstCellUsed().CellBelow(), wsht.LastColumnUsed().LastCellUsed());
            if (shtFmt.Sort.Count > 0)
            {
                foreach (Sort s in shtFmt.Sort)
                {
                    tableSortRange.SortColumns.Add(GetColumnNumber(shtFmt, s.Column), (s.Direction == "descending" ? XLSortOrder.Descending : XLSortOrder.Ascending));
                    if (lc.Count > 0 && s.Column.Equals(lc.Last().Name))
                        tableRange.SortColumns.Add(sortColNumber);
                }
                tableSortRange.Sort();
            }



            // At this point we have filled and formatted header data and value cells data
            // Now it's time to add totals and outlines if needed

            if (groupLevel > 0)
            //if (1==2)
            {
                //IXLWorksheet wshtTemp = wout.AddWorksheet("Temp");
                Dictionary<int, List<RowsRange>> groupRanges = new Dictionary<int, List<RowsRange>>();
                var dict = new Dictionary<String, double>();

                wsht.Outline.SummaryVLocation = XLOutlineSummaryVLocation.Bottom;
                int initLastRowNumber = tableRange.FirstRow().WorksheetRow().RowAbove().RowNumber();
                int initFirstRowNumber = tableRange.FirstRow().WorksheetRow().RowNumber();
                //[ColNumber][totalLevel][elementValue]
                Dictionary<int, Dictionary<int, List<Double>>> elements = new Dictionary<int, Dictionary<int, List<Double>>>();
                // Get mapping of group level (the greater - the more detail) and array of array of ranges of cells to calc aggregate func(if several columns to aggregate)
                for (int i = 0; i < colFuncs.Count; i++)
                {
                    int col = colFuncs.Keys.ElementAt(i);
                    elements[col] = new Dictionary<int, List<double>>();

                    for (int j = 0; j < groupLevel + 1; j++)
                    {
                        elements[col][j] = new List<double>();
                    }
                }

                for (int j = 0; j < groupLevel + 1; j++)
                {
                    groupRanges[j] = new List<RowsRange>();
                    groupRanges[j].Add(new RowsRange(initFirstRowNumber, initFirstRowNumber - 1));
                }

                prevRow = null;
                curRow = tableRange.FirstRow().WorksheetRow();

                int lastRowNumber = wsht.LastRowUsed().RowNumber();
                while (curRow.RowNumber() <= lastRowNumber)
                {
                    int curRowNumber = curRow.RowNumber();
                    //key = getKeyValue(curRow, lc);
                    //like 1 col - 9 func, 2 col - 10 func
                    //foreach (var item in colFuncs)
                    //{
                    //if key is changed
                    int totalLevel = groupLevel - Convert.ToInt32(curRow.Cell(sortColNumber).Value);
                    if (totalLevel == groupLevel) // data row
                    {
                        foreach (var coLFunc in colFuncs)
                        {
                            for (int j = 0; j < groupLevel + 1; j++)
                            {
                                String vStr = wsht.Cell(curRowNumber, coLFunc.Key).Value.ToString().Trim();
                                Double val = (String.IsNullOrEmpty(vStr) ? 0 : (Double)wsht.Cell(curRowNumber, coLFunc.Key).Value);
                                elements[coLFunc.Key][j].Add(val);
                            }
                        }
                        for (int i = 0; i < groupLevel + 1; i++)
                        {
                            groupRanges[i].Last().endRow++;
                        }
                    }
                    else //total row
                    {
                        int groupColumn = lc[totalLevel].Number;
                        var prevValue = curRow.Cell(groupColumn).Value;
                        wsht.Row(curRowNumber).Clear();
                        curRow.Cell(groupColumn).Value = prevValue;
                        SetTotalRowStyle(totalLevel, curRow, lc, colFuncs, shtFmt);

                        //writing totals for every function
                        foreach (var colFunc in (from f in colFuncs orderby f.Key select f))
                        {
                            int rL = Convert.ToInt32((double)totalLevel);
                            //Double tempResult = evaluateTemp(wshtTemp, rL + 1, aggregateFunctions[colFunc.Value]);
                            String calcMode = GetCalcMode(colFunc.Key, shtFmt);
                            if (calcMode == _calcModeInternal)
                            {
                                Double tempResult = GetGroupResult(elements[colFunc.Key][rL + 1], colFunc.Value);//.Sum();
                                curRow.Cell(colFunc.Key).Value = tempResult;
                            }
                            else if (calcMode == _calcModeFormula)
                            {
                                String range = "";
                                RowsRange row = groupRanges[totalLevel + 1].Last();
                                range = "R[" + (row.startRow - curRowNumber) + "]C" + ":" + "R[" + (row.endRow - curRowNumber) + "]C";
                                curRow.Cell(colFunc.Key).FormulaR1C1 = "=SUBTOTAL(" + colFunc.Value + "," + range + ")";
                            }
                            elements[colFunc.Key][rL + 1].Clear();
                        }

                        groupRanges[totalLevel + 1].Add(new RowsRange(curRowNumber + 1, curRowNumber));
                        for (int i = 0; i < groupLevel + 1; i++)
                        {
                            if (i < totalLevel + 1)
                                groupRanges[i].Last().endRow++;
                            else if (i > totalLevel + 1)
                            {
                                groupRanges[i].Last().endRow++;
                                groupRanges[i].Last().startRow++;
                            }

                        }
                    }
                    //}
                    //prevKey = key;
                    prevRow = curRow;
                    curRow = curRow.RowBelow();
                }
                foreach (var item in (from rg in groupRanges orderby rg.Key descending select rg))
                {

                    foreach (var r in item.Value)
                    {
                        wsht.Rows(r.startRow, r.endRow).Group();
                        wsht.Rows(r.startRow, r.endRow).Collapse();
                    }
                }
                //constructing grandtotal
                int grandTotalRowNumber = tableSortRange.LastRowUsed().RowNumber() + 1;
                String prefix = "";
                if (_options.ContainsKey(@"grand-total-prefix"))
                    prefix = prefix + _options[@"grand-total-prefix"] + " ";
                wsht.Cell(grandTotalRowNumber, wsht.FirstColumnUsed().ColumnNumber()).Value = prefix + "Grand total";

                SetGrandTotalRowStyle(0, wsht.LastRowUsed(), lc, colFuncs, shtFmt);
                foreach (var colFunc in (from f in colFuncs orderby f.Key select f))
                {
                    String calcMode = GetCalcMode(colFunc.Key, shtFmt);
                    if (calcMode == _calcModeInternal)
                    {
                        Double grandTotal = GetGroupResult(elements[colFunc.Key][0], colFunc.Value);//.Sum();
                        wsht.Cell(grandTotalRowNumber, colFunc.Key).Value = grandTotal;
                    }
                    else if (calcMode == _calcModeFormula)
                    {
                        String range = "";
                        RowsRange row = groupRanges[0].Last();
                        range = "R[" + (row.startRow - grandTotalRowNumber) + "]C" + ":" + "R[" + (row.endRow - grandTotalRowNumber) + "]C";
                        curRow.Cell(colFunc.Key).FormulaR1C1 = "=SUBTOTAL(" + colFunc.Value + "," + range + ")";
                    }
                }
                //Clear column, used for sorting
                wsht.Column(sortColNumber).Clear();
                //wshtTemp.Delete();
            }
            // Adjust columns width
            wsht.Columns().AdjustToContents();
            colNum = 1;
            foreach (xlsxfmt.format.Column colFmt in shtFmt.Column)
            {
                #region width
                double w = 0;
                double.TryParse(colFmt.Width, out w);
                if (w != 0)
                    wsht.Column(colNum).Width = w;
                #endregion
                colNum++;
            }
        }

    }
}
