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
using System.Threading;
using System.Resources;

namespace xlsxfmt
{
    public class XlsxFormatter
    {
        private string _sourceXlsx;
        private string _formatYaml;
        private string _outputXlsxBase;
        private string _outputXlsxBaseName;
        private string _outputXlsxBaseExt;
        private List<String> optionKeys = new List<string>(new string[] { "output-filename-prefix", "output-filename-postfix", "grand-total-prefix", "burst-on-column", "max-thread-amount" });
        private Dictionary<string, string> _options = new Dictionary<string, string>();
        private Dictionary<int, string> _aggregateFunctions = new Dictionary<int, string>();
        private Dictionary<String, int> _moveTotalSheets = new Dictionary<string, int>();
        //private ResourceManager ErrorMessages.ResourceManager;
        private List<String> _logoSheets = new List<string>();
        private string _delimiter = "~";
        private string _calcModeInternal = "internal";
        private string _calcModeFormula = "formula";
        private string _burstOnColumn = "burst-on-column";
        //private Image _logo;
        private int _logoHeight;
        private int _logoWidth;
        private YamlFile _yaml;
        private Stream _logoStream;
        private float _logoDpiY;
        private float _logoHorizontalResolution;
        private float _logoVerticalResolution;
        private int _maxThreadAmount;
        private int _defaultMaxThreadAmount = 3;
        private string _maxThreadNumberOptionName = "max-thread-amount";

        public void ValidateArguments()
        {
            if (String.IsNullOrEmpty(_sourceXlsx) || String.IsNullOrWhiteSpace(_sourceXlsx))
            {
                throw new System.ArgumentException("Source filename should not be empty or null or whitespaces");
            }
            if (String.IsNullOrEmpty(_formatYaml) || String.IsNullOrWhiteSpace(_formatYaml))
            {
                throw new System.ArgumentException("Format filename should not be empty or null or whitespaces");
            }

            foreach (var item in _options)
            {
                if (!optionKeys.Contains(item.Key))
                {
                    throw new System.ArgumentException("Illegal argument option \"" + item.Key + "\" specified. Please, check usage note.");
                }
            }

        }

        public XlsxFormatter(string[] args)
        {
            ParseArguments(args);
            ValidateArguments();
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
                        //Console.WriteLine("Parsing options: " + args[i]);
                        _outputXlsxBase = args[i];
                        int lastIndex = args[i].LastIndexOf('.');
                        _outputXlsxBaseName = args[i].Substring(0, lastIndex);
                        _outputXlsxBaseExt = args[i].Substring(lastIndex);

                        //break;
                    }
                    else
                    {
                        //Console.WriteLine("Parsing options: " + args[i]);
                        MatchCollection col = Regex.Matches(args[i], @"(\w+(?:-\w+)+)=(?:\""?)([0-9a-zA-Z_ ]+)(?:\""?)");
                        if (col.Count == 0)
                        {
                            throw new System.ArgumentException("Illegal argument option \"" + args[i] + "\" specified. Please, check usage note.");
                        }
                        else
                        {
                            foreach (Match m in col)
                            {
                                _options.Add(m.Groups[1].Value, m.Groups[2].Value);
                            }
                        }

                    }
                }
            }
        }

        private void GetLogoSheets()
        {
            foreach (var sheet in _yaml.Sheet)
            {
                if (!String.IsNullOrEmpty(sheet.IncludeLogo) && sheet.IncludeLogo.Equals("true"))
                    _logoSheets.Add(sheet.Name);
            }
        }

        private void GetMoveTotalSheets()
        {
            foreach (var sheet in _yaml.Sheet)
            {
                Dictionary<int, int> colFuncs = GetColFuncs(sheet);
                String freezeCell;
                if (!string.IsNullOrEmpty(sheet.FreezeOnCell))
                    freezeCell = sheet.FreezeOnCell;
                else
                    freezeCell = _yaml.Defaults.Sheet.FreezeOnCell;
                int freezeCol = GetFreezeCol(freezeCell);
                // Mapping (Name of sheet, minimum number of totalling columns
                if (colFuncs.Count > 0)
                {
                    int colNum = colFuncs.Keys.Min();
                    if (colNum > freezeCol)
                        _moveTotalSheets.Add(sheet.Name, colNum);
                }
            }
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
        private void InsertImage(WorksheetPart sheet1, int startRowIndex, int startColumnIndex, int endRowIndex, int endColumnIndex, Stream imageStream)
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
            GenerateDrawingsPart1Content(drawingsPart1);
            //Adding the image
            ImagePart imagePart1 = drawingsPart1.AddNewPart<ImagePart>("image/jpeg", "rId1");
            imageStream.Seek(0, SeekOrigin.Begin);
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
            var sh = r.FirstOrDefault(s => s.Name == sheetName);
            String id = "";
            if (sh != null)
                id = sh.Id;
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
        public void GenerateDrawingsPart1Content(DrawingsPart drawingsPart1)
        {
            Xdr.WorksheetDrawing worksheetDrawing1 = new Xdr.WorksheetDrawing();
            worksheetDrawing1.AddNamespaceDeclaration("xdr", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing");
            worksheetDrawing1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");


            Xdr.NonVisualDrawingProperties nvdp = new Xdr.NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Picture 1" };
            DocumentFormat.OpenXml.Drawing.PictureLocks picLocks = new DocumentFormat.OpenXml.Drawing.PictureLocks();
            picLocks.NoChangeAspect = true;
            picLocks.NoChangeArrowheads = true;
            Xdr.NonVisualPictureDrawingProperties nvpdp = new Xdr.NonVisualPictureDrawingProperties();
            nvpdp.PictureLocks = picLocks;
            Xdr.NonVisualPictureProperties nvpp = new Xdr.NonVisualPictureProperties();
            nvpp.NonVisualDrawingProperties = nvdp;
            nvpp.NonVisualPictureDrawingProperties = nvpdp;

            DocumentFormat.OpenXml.Drawing.Stretch stretch = new DocumentFormat.OpenXml.Drawing.Stretch();
            stretch.FillRectangle = new DocumentFormat.OpenXml.Drawing.FillRectangle();

            Xdr.BlipFill blipFill = new Xdr.BlipFill();
            A.Blip blip = new A.Blip();
            blip.Embed = "rId1";
            blip.CompressionState = A.BlipCompressionValues.Print;
            blipFill.Blip = blip;
            blipFill.SourceRectangle = new DocumentFormat.OpenXml.Drawing.SourceRectangle();
            blipFill.Append(stretch);

            DocumentFormat.OpenXml.Drawing.Transform2D t2d = new DocumentFormat.OpenXml.Drawing.Transform2D();
            DocumentFormat.OpenXml.Drawing.Offset offset = new DocumentFormat.OpenXml.Drawing.Offset();
            offset.X = 0;
            offset.Y = 0;
            t2d.Offset = offset;

            A.Extents extents = new A.Extents();

            //if (width == null)
            extents.Cx = (long)_logoWidth * (long)((float)914400 / _logoHorizontalResolution);
            //else
            //   extents.Cx = width;

            //  if (height == null)
            extents.Cy = (long)_logoHeight * (long)((float)914400 / _logoVerticalResolution);
            // else
            //    extents.Cy = height;

            // bm.Dispose();
            t2d.Extents = extents;
            Xdr.ShapeProperties sp = new Xdr.ShapeProperties();
            sp.BlackWhiteMode = DocumentFormat.OpenXml.Drawing.BlackWhiteModeValues.Auto;
            sp.Transform2D = t2d;
            DocumentFormat.OpenXml.Drawing.PresetGeometry prstGeom = new DocumentFormat.OpenXml.Drawing.PresetGeometry();
            prstGeom.Preset = DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle;
            prstGeom.AdjustValueList = new DocumentFormat.OpenXml.Drawing.AdjustValueList();
            sp.Append(prstGeom);
            sp.Append(new DocumentFormat.OpenXml.Drawing.NoFill());

            DocumentFormat.OpenXml.Drawing.Spreadsheet.Picture picture = new DocumentFormat.OpenXml.Drawing.Spreadsheet.Picture();
            picture.NonVisualPictureProperties = nvpp;
            picture.BlipFill = blipFill;
            picture.ShapeProperties = sp;

            DocumentFormat.OpenXml.Drawing.Spreadsheet.Position pos = new DocumentFormat.OpenXml.Drawing.Spreadsheet.Position();
            pos.X = 0;
            pos.Y = 0;
            DocumentFormat.OpenXml.Drawing.Spreadsheet.Extent ext = new DocumentFormat.OpenXml.Drawing.Spreadsheet.Extent();
            ext.Cx = extents.Cx;
            ext.Cy = extents.Cy;
            DocumentFormat.OpenXml.Drawing.Spreadsheet.AbsoluteAnchor anchor = new DocumentFormat.OpenXml.Drawing.Spreadsheet.AbsoluteAnchor();
            anchor.Position = pos;
            anchor.Extent = ext;
            anchor.Append(picture);
            anchor.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.ClientData());
            worksheetDrawing1.Append(anchor);
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
                Uri uri = new Uri(logoPath, UriKind.Absolute);
                logoStream = new FileStream(uri.AbsolutePath, FileMode.Open, FileAccess.Read, FileShare.Read);
                Image logo = Image.FromStream(logoStream);
                _logoHeight = logo.Height;
                _logoWidth = logo.Width;
                _logoHorizontalResolution = logo.HorizontalResolution;
                _logoVerticalResolution = logo.VerticalResolution;
                Graphics g = Graphics.FromImage(logo);
                _logoDpiY = g.DpiY;
                g.Dispose();
                //_logoStream = Stream.Synchronized(logoStream);
                logoStream.Seek(0, SeekOrigin.Begin);
            }
            catch (Exception ex)
            {

            }
            return logoStream;
        }

        /* private void DeserializeStopValues()
         {
             foreach (var sheet in _yaml.Sheet)
             {
                 foreach (var col in sheet.Column)
                 {
                     if (!String.IsNullOrEmpty(col.stopValues))
                     {
                         col.stopValuesList = Regex.Split(col.stopValues, @"((?:\""?)([0-9a-zA-Z_ ]+)(?:\""?))+");
                     }
                 }
             }
         }*/

        private List<String> GetBurstColumnValues(XLWorkbook sourceWorkbook, String burstColumnName)
        {
            List<String> result = new List<string>();
            foreach (var sheet in sourceWorkbook.Worksheets)
            {
                var srcColumn = sheet.Columns().Where(x => x.Cell(1).Value.ToString()
                   == burstColumnName).FirstOrDefault();
                if (srcColumn == null)
                {
                    throw new System.ArgumentException("Sheet \"" + sheet.Name + "\" does not contain burst column \"" + burstColumnName +
                                                       "\". Please, check source file.");
                }
                else
                {
                    var cellCnt = srcColumn.Cells().Count();
                    for (int i = 2; i <= cellCnt; i++)
                    {
                        String colValue = srcColumn.Cell(i).Value.ToString();
                        if (!result.Contains(colValue))
                        {
                            result.Add(colValue);
                        }
                    }
                }
            }
            return result;
        }

        private void AddImageAndPaneMove(String workbookName, bool needLogoUsage, bool needMoveTotalColumn)
        {
            if (needLogoUsage || needMoveTotalColumn)
            {
                using (SpreadsheetDocument document = SpreadsheetDocument.Open(workbookName, true))
                {
                    WorkbookPart workbookpart = document.WorkbookPart;
                    //WorksheetPart sheet1 = workbookpart.WorksheetParts.First();
                    if (needLogoUsage)
                    {
                        foreach (var customSheet in _logoSheets)
                        {
                            WorksheetPart sheet1 = GetSheetByName(workbookpart, customSheet);

                            if (sheet1 != null)
                            {
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
                                // Stream logoStream = loadLogo(_yaml.Format.LogoPath);
                                //Stream stream = Stream.Synchronized(_logoStream);
                                Uri uri = new Uri(_yaml.Format.LogoPath);
                                Stream stream = new FileStream(uri.AbsolutePath, FileMode.Open, FileAccess.Read, FileShare.Read);
                                InsertImage(sheet1, 0, 0, 1, numberOfColumns, stream);
                            }
                        }
                    }
                    if (needMoveTotalColumn)
                    {
                        foreach (var customSheet in _moveTotalSheets)
                        {
                            WorksheetPart sheet1 = GetSheetByName(workbookpart, customSheet.Key);
                            //!!!panes
                            if (sheet1 != null)
                            {
                                var panes = sheet1.Worksheet.Descendants<Pane>();
                                foreach (var it in panes)
                                {
                                    it.TopLeftCell = GetExcelColumnName(customSheet.Value) + "2";
                                }
                            }
                        }
                    }
                    document.WorkbookPart.Workbook.Save();
                    // Close the document handle.
                    document.Close();
                }
                //   validateXlsx(_outputXlsx);
            }
        }

        public void ConstructWorkBook(XLWorkbook sourceWorkbook, XLWorkbook outputWorkbook, String outputFileName, bool needLogoUsage, bool needMoveTotalColumn, String burstColumnName, String burstColumnValue)
        {
            // try
            //{
            // Construct output workbook using source workbook and input params
            Console.WriteLine("Thread id:" + Thread.CurrentThread.ManagedThreadId + " started processing at " + System.DateTime.Now);

            Construct(sourceWorkbook, outputWorkbook, needLogoUsage, burstColumnName, burstColumnValue);

            Console.WriteLine("Thread id:" + Thread.CurrentThread.ManagedThreadId + " finished processing at " + System.DateTime.Now);

            // Save output file
            if (outputWorkbook.Worksheets.Count > 0)
            {
                Console.WriteLine("Thread id:" + Thread.CurrentThread.ManagedThreadId + " started saving file at " + outputFileName + " at " + System.DateTime.Now);
                outputWorkbook.SaveAs(outputFileName);
                Console.WriteLine("Thread id:" + Thread.CurrentThread.ManagedThreadId + " finished saving file at " + outputFileName + " at " + System.DateTime.Now);
                Console.WriteLine("Thread id:" + Thread.CurrentThread.ManagedThreadId + " started processing image insert and column moving at " + System.DateTime.Now);
                AddImageAndPaneMove(outputFileName, needLogoUsage, needMoveTotalColumn);
                Console.WriteLine("Thread id:" + Thread.CurrentThread.ManagedThreadId + " finished processing image insert and column moving at " + System.DateTime.Now);
            }
            else
            {
                Console.WriteLine(String.Format(ErrorMessages.ResourceManager.GetString("NoSheets"), outputFileName));
            }
            //  }
            // catch (Exception ex)
            //{
            //    Console.WriteLine(ex.Message);
            // }
        }

        public void ValidateFormatFile(XLWorkbook wsrc)
        {
            int numErrors = 0;
            foreach (format.Sheet sheet in _yaml.Sheet)
            {
                var shtSource = sheet.Name;
                if (!string.IsNullOrEmpty(sheet.Source))
                    shtSource = sheet.Source;
                if (String.IsNullOrEmpty(shtSource))
                {
                    Console.WriteLine(String.Format(ErrorMessages.ResourceManager.GetString("NullSheetSource"), sheet.Name));
                    numErrors++;
                }
                else
                {
                    var ssht = wsrc.Worksheets.Where(x => x.Name == shtSource).FirstOrDefault();
                    if (ssht == null)
                    {
                        Console.WriteLine(String.Format(ErrorMessages.ResourceManager.GetString("NullSheetSource"), shtSource));
                        numErrors++;
                    }
                    else
                    {
                        foreach (format.Column col in sheet.Column)
                        {
                            string source = col.Name;
                            if (!string.IsNullOrEmpty(col.Source))
                                source = col.Source;
                            if (String.IsNullOrEmpty(source))
                            {
                                Console.WriteLine(String.Format(ErrorMessages.ResourceManager.GetString("NullColumnSource"), col.Name, sheet.Name));
                                numErrors++;
                            }
                            else
                            {
                                var srcCol = ssht.Columns().Where(x => x.Cell(1).Value.ToString() == source).FirstOrDefault();
                                if (srcCol == null)
                                {
                                    Console.WriteLine(String.Format(ErrorMessages.ResourceManager.GetString("NullColumnSource"), col.Name, sheet.Name));
                                    numErrors++;
                                }
                            }

                        }
                    }
                }
            }
            if (numErrors > 0)
            {
                throw new Exception(ErrorMessages.ResourceManager.GetString("FormatFileError"));
            }
        }

        public void Process()
        {
            // Read input parameters
            var formatReader = new StreamReader(_formatYaml);
            var formatDeserializer = new Deserializer(namingConvention: new CamelCaseNamingConvention(), ignoreUnmatched: true);
            _yaml = formatDeserializer.Deserialize<YamlFile>(formatReader);

            // Read source workbook
            var sourceWorkbook = new XLWorkbook(_sourceXlsx);
            //ErrorMessages.ResourceManager = new ResourceManager("ErrorMessages", Assembly.GetExecutingAssembly());
            ValidateFormatFile(sourceWorkbook);

            bool needBursting = false;
            int numOutputBooks = 1;

            List<String> burstColumnValues = new List<string>();
            String burstColumnName = "";
            if (_options.ContainsKey(_burstOnColumn))
            {
                burstColumnName = _options[_burstOnColumn];
            }
            else
                burstColumnName = _yaml.Format.BurstOnColumn;

            if (_options.ContainsKey(_maxThreadNumberOptionName))
            {
                Int32.TryParse(_options[_maxThreadNumberOptionName], out _maxThreadAmount);
                if (_maxThreadAmount <= 0)
                {
                    _maxThreadAmount = _defaultMaxThreadAmount;
                    Console.WriteLine("Specified max thread number is invalid and was ignored. Max number of threads was set to " + _defaultMaxThreadAmount);
                }
            }
            else
                _maxThreadAmount = _defaultMaxThreadAmount;

            if (!String.IsNullOrEmpty(burstColumnName))
            {
                burstColumnValues = GetBurstColumnValues(sourceWorkbook, burstColumnName);
                needBursting = burstColumnValues.Count > 0;
                if (needBursting)
                    numOutputBooks = burstColumnValues.Count;
            }
            sourceWorkbook.Dispose();

            // Initialize aggregation functions
            InitAggregateFunctions();

            // Get logo sheets
            GetLogoSheets();


            Stream logoStream = loadLogo(_yaml.Format.LogoPath);
            _logoStream = Stream.Synchronized(logoStream);
            bool needLogoUsage = (_logoSheets.Count > 0) && (logoStream != null);

            GetMoveTotalSheets();

            bool needMoveTotalColumn = _moveTotalSheets.Count > 0;

            Dictionary<String, KeyValuePair<String, XLWorkbook>> outputs = new Dictionary<string, KeyValuePair<String, XLWorkbook>>();
            if (needBursting)
            {
                foreach (var item in burstColumnValues)
                {
                    String outputFileName = ConstructOutputFilename(item);
                    outputs[item] = new KeyValuePair<string, XLWorkbook>(outputFileName, new XLWorkbook());
                }
            }
            else
            {
                String outputFileName = ConstructOutputFilename(null);
                outputs[""] = new KeyValuePair<string, XLWorkbook>(outputFileName, new XLWorkbook());
            }
            /*
            using (ManualResetEvent e = new ManualResetEvent(false))
            {
                ThreadPool.SetMaxThreads(_maxThreadAmount, _maxThreadAmount);
                foreach (var item in outputs)
                {
                    ThreadPool.QueueUserWorkItem(new WaitCallback(x =>
                    {
                        //XLWorkbook src = sourceWorkbook;
                        // Read source workbook
                        using (var src = new XLWorkbook(_sourceXlsx))
                        {
                            ConstructWorkBook(src, item.Value.Value, item.Value.Key, needLogoUsage, needMoveTotalColumn, burstColumnName, item.Key);
                            GC.Collect();
                        }
                        if (Interlocked.Decrement(ref numOutputBooks) == 0)
                            e.Set();
                    }
                    )
                    );
                }
                e.WaitOne();
            }
            */
            var parallelOptions = new ParallelOptions
            {
                MaxDegreeOfParallelism = _maxThreadAmount
            };
            Parallel.ForEach(outputs, parallelOptions, (output, index) =>
            {
                using (var src = new XLWorkbook(_sourceXlsx))
                {
                    ConstructWorkBook(src, output.Value.Value, output.Value.Key, needLogoUsage, needMoveTotalColumn, burstColumnName, output.Key);
                    //GC.Collect();
                }
            }
            );
        }


        private String ConstructOutputFilename(String burstColumnValue)
        {
            String result;
            if (string.IsNullOrEmpty(_outputXlsxBase))
            {
                result = _yaml.Format.OutputFilenameBase;

                if (burstColumnValue != null)
                {
                    result = burstColumnValue + result;
                }

                var fileName = Path.GetFileName(result);

                if (_options.ContainsKey(@"output-filename-prefix"))
                {
                    fileName = _options[@"output-filename-prefix"] + fileName;
                }



                if (_options.ContainsKey(@"output-filename-postfix"))
                    fileName = fileName + _options[@"output-filename-postfix"];

                var dir = Path.GetPathRoot(result);
                if (!string.IsNullOrEmpty(dir))
                    result = dir + "\\" + fileName + @".xlsx";
                else
                    result = fileName + @".xlsx";
            }
            else
            {
                result = _outputXlsxBaseName + " " + burstColumnValue + " " + _outputXlsxBaseExt;
            }
            return result;
        }

        private void Construct(XLWorkbook wsrc, XLWorkbook wout, bool needLogoUsage, String burstColumnName, String burstColumnValue)
        {
            // Construct sheets
            foreach (var shtFmt in _yaml.Sheet)
            {
                //if (shtFmt.Name.IndexOf("Unpriced") >= 0)
                //{
                    var source = shtFmt.Name;
                    if (!string.IsNullOrEmpty(shtFmt.Source))
                        source = shtFmt.Source;

                    // Find source sheet in source workbook
                    var ssht = wsrc.Worksheets.Where(x => x.Name == source).FirstOrDefault();
                    if (ssht != null)
                    {
                        Console.WriteLine("Thread id:" + Thread.CurrentThread.ManagedThreadId.ToString() + " started constructing sheet " + source + " at " + System.DateTime.Now);
                        ConstructSheet(ssht, wout, shtFmt, needLogoUsage, burstColumnName, burstColumnValue);
                        Console.WriteLine("Thread id:" + Thread.CurrentThread.ManagedThreadId.ToString() + " finished constructing sheet " + source + " at " + System.DateTime.Now);
                    }
                //}
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
        private string GetExcelColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
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
            if (sheet.Font != null && sheet.Font.Footer != null)
                cellFont = sheet.Font;
            else
                cellFont = _yaml.Defaults.Font;
            if (cellFont != null && cellFont.Footer != null)
            {
                //size
                if (!String.IsNullOrEmpty(cellFont.Size))
                {
                    double fontSz;
                    Double.TryParse(cellFont.Size, out fontSz);
                    row.Style.Font.SetFontSize(fontSz);
                }
                //style
                String dataStyle = cellFont.Footer.Style;
                if (!String.IsNullOrEmpty(dataStyle))
                {
                    if (dataStyle == "bold")
                    {
                        row.Style.Font.SetBold();
                    }
                    else if (dataStyle == "italic")
                    {
                        row.Style.Font.SetItalic();
                    }
                    else if (dataStyle == "underline")
                    {
                        row.Style.Font.SetUnderline();
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
            if (sheet.Font != null && sheet.Font.Footer != null)
                cellFont = sheet.Font;
            else
                cellFont = _yaml.Defaults.Font;
            if (cellFont != null && cellFont.Footer != null)
            {
                //size
                if (!String.IsNullOrEmpty(cellFont.Size))
                {
                    double fontSz;
                    Double.TryParse(cellFont.Size, out fontSz);
                    row.Style.Font.SetFontSize(fontSz);
                }
                //style
                String dataStyle = cellFont.Footer.Style;
                if (!String.IsNullOrEmpty(dataStyle))
                {
                    if (dataStyle == "bold")
                    {
                        row.Style.Font.SetBold();
                    }
                    else if (dataStyle == "italic")
                    {
                        row.Style.Font.SetItalic();
                    }
                    else if (dataStyle == "underline")
                    {
                        row.Style.Font.SetUnderline();
                    }
                }
                //conditional-formatting
            }
            foreach (var item in colFuncs)
            {
                try
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
                catch (Exception ex)
                {

                }
            }

        }
        #region group functions
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
        #endregion
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

        private List<int> GetExcludedRows(IXLWorksheet sourceSheet, xlsxfmt.format.Sheet shtFmt)
        {
            List<int> excludedRows = new List<int>();
            foreach (var col in shtFmt.Column.Where(x => x.stopValues != null && x.stopValues.Count != 0))
            {
                var srcColumn = sourceSheet.Columns().Where(x => x.Cell(1).Value.ToString()
                    == col.Name).FirstOrDefault();
                if (srcColumn != null)
                {
                    var cellCollection = srcColumn.Cells();
                    var cellCnt = cellCollection.Count();
                    foreach (var cell in cellCollection)
                    {
                        int rn = cell.WorksheetRow().RowNumber();
                        if (rn == 1) continue;
                        String v = cell.Value.ToString();
                        if (col.stopValues.Any(x => x.stopValue == v) && !excludedRows.Contains(rn))
                        {
                            excludedRows.Add(rn);
                        }
                    }
                }
            }
            return excludedRows;
        }
        private List<int> GetIncludedRows(IXLWorksheet sourceSheet, xlsxfmt.format.Sheet shtFmt, out bool filteredByRequiredValues)
        {
            List<int> includedRows = new List<int>();
            filteredByRequiredValues = false;
            foreach (var col in shtFmt.Column.Where(x => x.requiredValues != null && x.requiredValues.Count != 0))
            {
                filteredByRequiredValues = true;
                var srcColumn = sourceSheet.Columns().Where(x => x.Cell(1).Value.ToString()
                    == col.Name).FirstOrDefault();
                if (srcColumn != null)
                {
                    //var cellCnt = srcColumn.Cells().Count();
                    //for (int i = 2; i <= cellCnt; i++)
                    var cellCollection = srcColumn.Cells();
                    var cellCnt = cellCollection.Count();
                    foreach (var cell in cellCollection)
                    {
                        int rn = cell.WorksheetRow().RowNumber();
                        if (rn == 1) continue;
                        String v = cell.Value.ToString();
                        if (col.requiredValues.Any(x => x.requiredValue == v) && !includedRows.Contains(rn))
                        {
                            includedRows.Add(rn);
                        }
                    }
                }
            }
            return includedRows;
        }

        private List<int> GetNeededRows(IXLWorksheet sourceSheet, format.Sheet shtFmt, List<int> rowsExcluded, List<int> rowsIncluded,bool includedFiltered, int burstColNumber, String burstColumnValue)
        {
            List<int> rows = new List<int>();
            Dictionary<int, xlsxfmt.sorting.FormatOrder> sortColumns = new Dictionary<int, xlsxfmt.sorting.FormatOrder>();
            int numRowsUsed = sourceSheet.RowsUsed().Count();
            // get column numbers to sort
            if (shtFmt.Sort != null && shtFmt.Sort.Count > 0)
            {
                foreach (Sort col in shtFmt.Sort)
                {
                    xlsxfmt.sorting.FormatOrder fo = new xlsxfmt.sorting.FormatOrder();
                    xlsxfmt.format.Column formatCol = shtFmt.Column.Where(x => x.Name == col.Column).FirstOrDefault();
                    fo.isAscending = (col.Direction != "descending");
                    fo.isDate = false;
                    fo.isNumeric = false;
                    fo.isString = false;
                    if (formatCol != null && formatCol.FormatType != null)
                    {
                        if (formatCol.FormatType == "DATE")
                            fo.isDate = true;
                        else if (formatCol.FormatType == "ACCOUNTING" || formatCol.FormatType == "NUMBER")
                            fo.isNumeric = true;
                        else
                            fo.isString = true;
                    }
                    else
                    {
                        fo.isString = true;
                    }
                    sortColumns.Add(GetSourceColumnNumber(sourceSheet, shtFmt, col.Column), fo);
                }
            }
            // get needed rows
            /*foreach (format.Column colFmt in shtFmt.Column)
            {
                var source = colFmt.Name;
                if (!string.IsNullOrEmpty(colFmt.Source))
                    source = colFmt.Source;
                // Find corresponding column in source sheet
                var cols = sourceSheet.Columns();
                var srcColumn = cols.Where(x => x.Cells().Count() > 0 && x.Cell(1).Value.ToString() == source).FirstOrDefault();
                var cellCnt = srcColumn.Cells().Count();
                for (int i = 2; i <= cellCnt; i++)
                {
                    if (!rows.Contains(i) && !rowsExcluded.Contains(i) && (rowsIncluded.Count() == 0 || rowsIncluded.Contains(i)))
                    {
                        rows.Add(i);
                    }
                }
            }*/
            for (int i = 2; i <= numRowsUsed; i++)
            {
                if (!rowsExcluded.Contains(i) && (!includedFiltered || rowsIncluded.Contains(i)))
                {
                    rows.Add(i);
                }
            }
            // perform Sorting
            if (sortColumns.Count > 0)
            {
                xlsxfmt.sorting.SortRows sortRows = new xlsxfmt.sorting.SortRows(rows.Count);
                for (int j = 0; j < rows.Count; j++)
                {
                    SortRow sr = new SortRow();
                    sr.index = rows.ElementAt(j);
                    sr.cellValues = new List<sorting.CellValue>();
                    for (int i = 0; i < sortColumns.Count; i++)
                    {
                        int colN = sortColumns.Keys.ElementAt(i);
                        xlsxfmt.sorting.FormatOrder fo = sortColumns[colN];
                        Object cell = sourceSheet.Cell(sr.index, colN).Value;
                        if (fo.isString)
                        {
                            sr.cellValues.Add(new sorting.CellValue(cell.ToString()));
                        }
                        else if (fo.isDate)
                        {
                            DateTime dt = new DateTime();
                            DateTime.TryParse(cell.ToString(), out dt);
                            sr.cellValues.Add(new sorting.CellValue(dt));
                        }
                        else if (fo.isNumeric)
                        {
                            Double d;
                            Double.TryParse(cell.ToString(), out d);
                            sr.cellValues.Add(new sorting.CellValue(d));
                        }
                    }
                    sortRows.rows[j] = sr;
                }
                // perform sort
                for (int i = 0; i < sortColumns.Count; i++)
                {
                    sortRows.sort(i, sortColumns.Values.ElementAt(i).isAscending);
                }
                // get result
                int ttlRowCnt = rows.Count;
                int rowCnt = ttlRowCnt;
                if (!String.IsNullOrEmpty(shtFmt.topNRows))
                {
                    int rc = 0;
                    if (Int32.TryParse(shtFmt.topNRows, out rc))
                    {
                        rowCnt = rc;
                    }
                }
                if (rowCnt < ttlRowCnt)
                {
                    rows.Clear();
                    int rowsChecked = 0;
                    int rowsGathered = 0;
                    //while not have topnrows needed
                    while (rowsGathered < rowCnt && rowsChecked < ttlRowCnt)
                    {
                        if (burstColNumber == 0 || (burstColNumber != 0 && sourceSheet.Cell(sortRows.rows[rowsChecked].index, burstColNumber).Value.Equals(burstColumnValue)))
                        {

                            rows.Add(sortRows.rows[rowsChecked].index);
                            rowsGathered++;
                        }
                        rowsChecked++;
                    }
                }

            }
            return rows;
        }

        private int GetSourceColumnNumber(IXLWorksheet sourceSheet, format.Sheet shtFmt, string dstColumnName)
        {
            format.Column c = shtFmt.Column.Where(x => x.Name == dstColumnName).FirstOrDefault();
            if (c == null)
                return -1;
            string sourceColumnName = c.Name;
            if (!String.IsNullOrEmpty(c.Source))
                sourceColumnName = c.Source;
            var col = sourceSheet.Columns().Where(x => x.Cells().Count() > 0 && x.Cell(1).Value.ToString() == sourceColumnName).FirstOrDefault();
            if (col != null)
                return col.ColumnNumber();
            return -1;
        }

        private void ConstructSheet(IXLWorksheet ssht, XLWorkbook wout, xlsxfmt.format.Sheet shtFmt, bool needLogoUsage, String burstColumnName, String burstColumnValue)
        {
            IXLWorksheet wsht;
            if (shtFmt.Source == null)
            {
                if (_yaml.Format.emptySheet != null)
                {
                    if (_yaml.Format.emptySheet.exclude == "true")
                    {
                        return;
                    }
                    if (_yaml.Format.emptySheet.DefaultText != null)
                    {
                        wsht = wout.AddWorksheet(shtFmt.Name);
                        wsht.Cell(2, 1).Value = _yaml.Format.emptySheet.DefaultText;
                        return;
                    }
                }
                return;

            }
            wsht = wout.AddWorksheet(shtFmt.Name);
            int logoRows = 0;
            int headerRows = 1;
            int startRowNum = 1;
            int numDataRows = 0;

            bool logoInserted = false;
            List<xlsxfmt.format.Column> lc = GetSubtotalCols(shtFmt);
            Dictionary<int, int> colFuncs = GetColFuncs(shtFmt);
            bool includedFiltered = false;
            List<int> rowsExcluded = GetExcludedRows(ssht, shtFmt);
            List<int> rowsIncluded = GetIncludedRows(ssht, shtFmt, out includedFiltered);
            // Adding group columns into sort column list
            foreach (var item in lc)
            {
                if (shtFmt.Sort == null)
                    shtFmt.Sort = new List<Sort>();
                if (!shtFmt.Sort.Any(x => x.Column == item.Name))
                {
                    Sort sr = new Sort();
                    sr.Direction = "";
                    sr.Column = item.Name;
                    shtFmt.Sort.Add(sr);
                }
            }
            int burstColNumber = 0;
            if (!String.IsNullOrEmpty(burstColumnName))
            {
                var c = ssht.Columns().Where(x => x.Cell(1).Value.ToString()
                        == burstColumnName).FirstOrDefault();
                if (c != null)
                {
                    burstColNumber = c.ColumnNumber();
                }
            }
            Console.WriteLine("Thread id:" + Thread.CurrentThread.ManagedThreadId.ToString() + " started getting needed rows in sheet " + shtFmt.Name + " at " + System.DateTime.Now);
            List<int> rowsSortedNeeded = GetNeededRows(ssht, shtFmt, rowsExcluded, rowsIncluded, includedFiltered, burstColNumber, burstColumnValue);
            Console.WriteLine("Thread id:" + Thread.CurrentThread.ManagedThreadId.ToString() + " finished getting needed rows in sheet " + shtFmt.Name + " at " + System.DateTime.Now);
            if (needLogoUsage && !String.IsNullOrEmpty(shtFmt.IncludeLogo) && shtFmt.IncludeLogo.Equals("true"))
            {
                // Graphics g = Graphics.FromImage(_logo);
                wsht.Row(1).Height = (int)(_logoHeight * 72 / _logoDpiY) + 1;//!!!
                //g.Dispose();
                logoRows = 1;
                logoInserted = true;
            }
            else
            {
                logoRows = 0;
            }
            int lastRowUsed = headerRows + logoRows;
            startRowNum += logoRows;
            int colNum = 1;
            int rowNum = startRowNum;
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
                if (logoInserted)
                {
                    wsht.Row(rowNum - 1).Style.Fill.SetBackgroundColor(XLColor.FromArgb(headerR, headerG, headerB));
                }
            }
            #endregion
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
            #region popalating with source data
            //for(int j=0; j<shtFmt.Column.Count; j++)
            List<IXLColumn> srcCols = new List<IXLColumn>();
            List<format.Column> formatCols = new List<format.Column>();
            foreach (format.Column colFmt in shtFmt.Column)
            {
                var source = colFmt.Name;
                if (!string.IsNullOrEmpty(colFmt.Source))
                    source = colFmt.Source;
                //if (source == "")
                // Find corresponding column in source sheet
                var cols = ssht.Columns();
                var srcColumn = cols.Where(x => x.Cells().Count() > 0 && x.Cell(1).Value.ToString() == source).FirstOrDefault();
                // try
                // {
                if (srcColumn != null)
                {
                    srcCols.Add(srcColumn);
                    formatCols.Add(colFmt);
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
                    // Populate output column cells
                    if (colFmt.hidden == "true")
                        wsht.Column(colNum).Hide();
                }
                if (colFmt.hidden == "true")
                    wsht.Column(colNum).Hide();
                colNum++;
            }
            colNum = 1;
            rowNum++;
            Console.WriteLine("Thread id:" + Thread.CurrentThread.ManagedThreadId.ToString() + " started populating with needed data sheet " + shtFmt.Name + " at " + System.DateTime.Now);
            bool wroteFlag = false;
            foreach (var row in rowsSortedNeeded)
            {
                for(int i =0;i<srcCols.Count;i++)
                {
                    if ((String.IsNullOrEmpty(burstColumnName) ||
                                (burstColNumber != 0 &&
                                    ssht.Cell(row, burstColNumber).Value.Equals(burstColumnValue)
                                )
                            ) //&& (srcCols[i].Cells().Any(x => x.WorksheetRow().RowNumber() == row))
                         && (srcCols[i].Cell(row).Value != null)
                        )
                    {
                       // rowNum++;
                        //wsht.Cell(rowNum, colNum).Value = srcColumn.Cell(i).Value;
                        wsht.Cell(rowNum, colNum).Value = srcCols[i].Cell(row).Value;
                        wroteFlag = true; // if were writing to cell
                        #region setdatacellstyle
                        xlsxfmt.format.Font cellFont;
                        if (formatCols[i].Font != null && formatCols[i].Font.Data != null)
                            cellFont = formatCols[i].Font;
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
                    }
                    Interlocked.Increment(ref colNum);
                }
                if (wroteFlag)
                    rowNum++;
                colNum = 1;
                wroteFlag = false;
            }
            Console.WriteLine("Thread id:" + Thread.CurrentThread.ManagedThreadId.ToString() + " finished populating with needed data sheet " + shtFmt.Name + " at " + System.DateTime.Now);
            numDataRows = rowNum - headerRows - logoRows - 1;
            #endregion
            colNum = 1;
            foreach (xlsxfmt.format.Column colFmt in shtFmt.Column)
            {
                #region column format
                var source = colFmt.Name;
                if (!string.IsNullOrEmpty(colFmt.Source))
                    source = colFmt.Source;
                var srcColumn = ssht.Columns().Where(x => x.Cells().Count() > 0 && x.Cell(1).Value.ToString() == source).FirstOrDefault();
                if (srcColumn != null)
                {
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
                }
                #endregion
                colNum++;
            }

            //List<xlsxfmt.format.Column> lc = GetSubtotalCols(shtFmt);
            //Dictionary<int, int> colFuncs = GetColFuncs(shtFmt);
            String key;
            String prevKey = "";
            IXLRow curRow, prevRow;
            int groupLevel = lc.Count;
            int lastColUsed = wsht.LastColumnUsed().ColumnNumber();
            //lastRowUsed += numDataRows;
            lastRowUsed = wsht.LastRowUsed().RowNumber();
            int sortColNumber = lastColUsed + 1;
            //var tableRange = wsht.Range(wsht.FirstCellUsed().CellBelow(), wsht.LastColumnUsed().LastCellUsed());
            var lru = wsht.LastRowUsed();
            var tableRange = wsht.Range(wsht.FirstCellUsed().CellBelow(), wsht.Row(lastRowUsed).LastCellUsed());
            if (groupLevel > 0 && numDataRows > 0)
            {
                Console.WriteLine("Thread id:" + Thread.CurrentThread.ManagedThreadId.ToString() + " started preparing for grouping in sheet " + shtFmt.Name + " at " + System.DateTime.Now);
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
                //lastRowUsed = wsht.LastRowUsed().RowNumber();
                var sortKeys = totalKeys.GroupBy(s => new { s.grouplevel, s.key }).Select(x => new SortKeyLevel(x.Key.key, x.Key.grouplevel)).ToList();
                for (int i = 0; i < sortKeys.Count(); i++)
                {
                    Interlocked.Increment(ref lastRowUsed);
                    wsht.Cell(lastRowUsed, sortColNumber).Value = groupLevel - totalKeys[i].grouplevel;
                    string[] keys = sortKeys[i].key.Split(new string[] { _delimiter }, StringSplitOptions.None);
                    for (int j = 0; j < keys.Length; j++)
                    {
                        if (!String.IsNullOrEmpty(keys[j]))
                            wsht.Cell(lastRowUsed, lc[j].Number).Value = keys[j];
                    }

                }
                // Adding group columns into sort column list
                /*foreach (var item in lc)
                {
                    if (shtFmt.Sort == null)
                        shtFmt.Sort = new List<Sort>();
                    if (!shtFmt.Sort.Any(x => x.Column == item.Name))
                    {
                        Sort sr = new Sort();
                        sr.Direction = "";
                        sr.Column = item.Name;
                        shtFmt.Sort.Add(sr);
                    }
                }*/
                Console.WriteLine("Thread id:" + Thread.CurrentThread.ManagedThreadId.ToString() + " finished preparing for grouping in sheet " + shtFmt.Name + " at " + System.DateTime.Now);

            }

            if (shtFmt.Sort != null && shtFmt.Sort.Count > 0 && numDataRows > 0)
            {
                Console.WriteLine("Thread id:" + Thread.CurrentThread.ManagedThreadId.ToString() + " started sorting sheet " + shtFmt.Name + " at " + System.DateTime.Now);
                //Sorting
                var fcu = wsht.FirstCellUsed();
                var lcu = wsht.LastCellUsed();
                var tableSortRange = wsht.Range(wsht.FirstCellUsed().CellBelow(), wsht.Row(lastRowUsed).LastCellUsed());

                foreach (Sort s in shtFmt.Sort)
                {
                    tableSortRange.SortColumns.Add(GetColumnNumber(shtFmt, s.Column), (s.Direction == "descending" ? XLSortOrder.Descending : XLSortOrder.Ascending));
                    if (lc.Count > 0 && s.Column.Equals(lc.Last().Name))
                        tableRange.SortColumns.Add(sortColNumber);
                }
                tableSortRange.Sort();
                Console.WriteLine("Thread id:" + Thread.CurrentThread.ManagedThreadId.ToString() + " finished sorting sheet " + shtFmt.Name + " at " + System.DateTime.Now);
            }

            GC.Collect();

            // At this point we have filled and formatted header data and value cells data
            // Now it's time to add totals and outlines if needed

            #region generating totals
            if (groupLevel > 0 && numDataRows > 0)
            //if (1==2)
            {
                //IXLWorksheet wshtTemp = wout.AddWorksheet("Temp");
                Console.WriteLine("Thread id:" + Thread.CurrentThread.ManagedThreadId.ToString() + " started generating totals in sheet " + shtFmt.Name + " at " + System.DateTime.Now);
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
                    int t;
                    Int32.TryParse(curRow.Cell(sortColNumber).Value.ToString(), out t);
                    int totalLevel = groupLevel - t;
                    if (totalLevel == groupLevel) // data row
                    {
                        foreach (var coLFunc in colFuncs)
                        {
                            for (int j = 0; j < groupLevel + 1; j++)
                            {
                                var rv = wsht.Cell(curRowNumber, coLFunc.Key);
                                if (rv != null)
                                {
                                    try
                                    {
                                        var vs = rv.Value;
                                        String vStr = vs.ToString().Trim();
                                        String v = wsht.Cell(curRowNumber, coLFunc.Key).Value.ToString();
                                        Double d;
                                        Double.TryParse(v, out d);
                                        Double val = (String.IsNullOrEmpty(vStr) ? 0 : d);
                                        elements[coLFunc.Key][j].Add(val);
                                    }
                                    catch (Exception ex)
                                    {
                                        elements[coLFunc.Key][j].Add(0);
                                    }
                                }
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
                        if (item.Key != 0)
                            wsht.Rows(r.startRow, r.endRow).Collapse();
                    }
                }
                //constructing grandtotal
                int grandTotalRowNumber = lastRowUsed + 1;
                String prefix = "";
                if (_options.ContainsKey(@"grand-total-prefix"))
                    prefix = prefix + _options[@"grand-total-prefix"] + " ";
                wsht.Cell(grandTotalRowNumber, wsht.FirstColumnUsed().ColumnNumber()).Value = prefix + "Grand total";

                SetGrandTotalRowStyle(0, wsht.Row(grandTotalRowNumber), lc, colFuncs, shtFmt);
                Interlocked.Increment(ref lastRowUsed);
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
                elements = null;
                Console.WriteLine("Thread id:" + Thread.CurrentThread.ManagedThreadId.ToString() + " finished generating totals in sheet " + shtFmt.Name + " at " + System.DateTime.Now);
            }
            #endregion
            /*
            #region process topnrows
            if (!String.IsNullOrEmpty(shtFmt.topNRows))
            {
                
                int topRows = 0;
                Int32.TryParse(shtFmt.topNRows, out topRows);
                if (topRows > 0 && topRows <= lastRowUsed)
                {
                    Console.WriteLine("Thread " + Thread.CurrentThread.ManagedThreadId + " started processing top-n-rows in sheet " + shtFmt.Name + " at " + System.DateTime.Now);
                    for (int i = topRows + 1; i <= lastRowUsed; i++)
                    {
                        wsht.Row(i).Delete();
                    }
                    Console.WriteLine("Thread " + Thread.CurrentThread.ManagedThreadId + " finished processing top-n-rows in sheet " + shtFmt.Name + " at " + System.DateTime.Now);
                }

            }
            #endregion
            */
            GC.Collect();
            #region empty sheet
            if (wsht.IsEmpty() || numDataRows == 0)
            {
                if (_yaml.Format.emptySheet != null)
                {
                    if (_yaml.Format.emptySheet.DefaultText != null)
                    {
                        wsht.Cell(2, 1).Value = _yaml.Format.emptySheet.DefaultText;
                    }
                    if (_yaml.Format.emptySheet.exclude == "true")
                    {
                        wsht.Delete();
                    }
                }
            }
            #endregion
            // Adjust columns width
            Console.WriteLine("Thread id:" + Thread.CurrentThread.ManagedThreadId.ToString() + " started adjusting columns to content in sheet " + shtFmt.Name + " at " + System.DateTime.Now);
            wsht.Columns().AdjustToContents();
            Console.WriteLine("Thread id:" + Thread.CurrentThread.ManagedThreadId.ToString() + " finished adjusting columns to content in sheet " + shtFmt.Name + " at " + System.DateTime.Now);
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
