using System;
using System.IO;
using System.Text;
using Aspose.Words;
using System.Threading.Tasks;
using System.IO.Compression;
using System.Drawing;

using IBM.Broker.Plugin;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Linq;
using static IntegrationPDFGeneration.CommonDefinitions;

namespace IntegrationPDFGeneration
{
    /// <summary>
    /// ModifyNode Class
    /// </summary>
    public class ModifyNode : NBComputeNode
    {
        /// <summary>
        /// Evaluate Method
        /// </summary>
        /// <param name="inputAssembly"></param>

        private Boolean isError = false;
        private String requestId = "";

        private enum InputFileTypes { RTF, CSV }
        private enum OutputFileTypes { PDF, CSV }

        public override void Evaluate(NBMessageAssembly inputAssembly)
        {
            NBOutputTerminal outTerminal = OutputTerminal("Out");

            NBMessage inputMessage = inputAssembly.Message;

            // Create a new message from a copy of the inboundMessage, ensuring it is disposed of after use
            using (NBMessage outputMessage = new NBMessage(inputMessage))
            {
                isError = false;
                NBMessageAssembly outAssembly = new NBMessageAssembly(inputAssembly, outputMessage);
                NBElement inputRoot = inputMessage.RootElement;
                NBElement outputRoot = outputMessage.RootElement;

                #region UserCode
                try
                {
                    // check if we need to just do a test
                    if (inputRoot["JSON"]["Data"]["headTest"] == null)
                    {
                        requestId = inputRoot["JSON"]["Data"]["requestId"].ValueAsString;

                        if (inputRoot["JSON"]["Data"]["asyncRequest"] != null)
                        {
                            String uploadFileName = "";

                            if (inputRoot["JSON"]["Data"]["templateDetails"]["templates"].FirstChild["isZipUpload"].ValueAsString == "FALSE")
                            {
                                uploadFileName = inputRoot["JSON"]["Data"]["templateDetails"]["templates"].FirstChild["generatedName"].ValueAsString;
                                uploadFileName = ReplaceDateTimeReferences(uploadFileName);
                            }

                            getUploadFileLocation(outputRoot, uploadFileName, inputRoot, requestId);
                        }
                        else
                        {
                            outputRoot["JSON"]["Data"].DeleteAllChildren();
                            applyLicense(outputRoot);

                            if (isError == false)
                            {
                                var pdfStreams = new List<PDFStream>();
                                var nonZipPDFStreams = new List<PDFStream>();

                                foreach (NBElement template in inputRoot["JSON"]["Data"]["templateDetails"]["templates"])
                                {
                                    String patentDataItem = template["parentDataItem"].ValueAsString;
                                    String csvCollatedFile = "";
                                    String isZipUpload = template["isZipUpload"].ValueAsString;

                                    String inputFileExtension = Path.GetExtension(template["templateName"].ValueAsString.ToUpper());
                                    String outputFileExtension = Path.GetExtension(template["generatedName"].ValueAsString.ToUpper());

                                    InputFileTypes inputFileType = (InputFileTypes)Enum.Parse(typeof(InputFileTypes), inputFileExtension.Replace(".", ""), true);
                                    OutputFileTypes outputFileType = (OutputFileTypes)Enum.Parse(typeof(OutputFileTypes), outputFileExtension.Replace(".", ""), true);

                                    String templateContent = "";

                                    if (inputFileType != InputFileTypes.CSV)
                                        templateContent = System.IO.File.ReadAllText(template["templateLocation"].ValueAsString);

                                    NBElement envIn = inputRoot["JSON"]["Data"];

                                    String documentIdentifierField = "";

                                    if (envIn["documentIdentifierField"] != null)
                                    {
                                        documentIdentifierField = envIn["documentIdentifierField"].ValueAsString;
                                    }

                                    Boolean elementFound = true;

                                    foreach (String elementName in patentDataItem.Split('.'))
                                    {
                                        if (envIn[elementName] != null)
                                            envIn = envIn[elementName];
                                        else
                                        {
                                            elementFound = false;
                                            break;
                                        }
                                    }

                                    if (elementFound == true && envIn.Children().Count() > 0)
                                    {
                                        List<NBElement> elementsToProcess = new List<NBElement>();

                                        if (envIn.FirstChild.Name == "Item" && envIn.LastChild.Name == "Item")
                                            elementsToProcess = envIn.ToList();
                                        else
                                            elementsToProcess.Add(envIn);

                                        String outputFileName = template["generatedName"].ValueAsString;


                                        if (outputFileType != OutputFileTypes.CSV)
                                        {
                                            Parallel.ForEach(elementsToProcess, docData =>
                                            {
                                                String documentContent = templateContent;
                                                String templateName = outputFileName;

                                                String identifier = "";

                                                if (documentIdentifierField != "")
                                                {
                                                    identifier = ReplaceAllFields(docData, "[[" + documentIdentifierField + "]]");
                                                    isZipUpload = "FALSE";
                                                }

                                                documentContent = ReplaceAllFields(docData, documentContent);
                                                templateName = ReplaceAllFields(docData, templateName);


                                                //Example of IF field - [[IFdepositAccount!=""|Deposit Account|]]
                                                Regex ifregex = new Regex(string.Format("\\[IF.*?\\]]"));
                                                foreach (Match m in ifregex.Matches(documentContent))
                                                {
                                                    String fieldContent = getFieldContent(docData, m.Value.Replace("[", "").Replace("]", ""));
                                                    documentContent = documentContent.Replace("[" + m.Value, fieldContent);
                                                }

                                                //Example of IF field - [[FORMATrenewalDate|dd-MM-yyyy]]
                                                Regex formatregex = new Regex(string.Format("\\[FORMAT.*?\\]]"));
                                                foreach (Match m in formatregex.Matches(documentContent))
                                                {
                                                    String fieldName = m.Value.Replace("[", "").Replace("]", "");
                                                    if (fieldName.IndexOf('|') > -1)
                                                    {
                                                        String fieldFormat = fieldName.Split('|')[1];
                                                        fieldName = fieldName.Split('|')[0].Replace("FORMAT","");

                                                        String fieldContent = ReplaceAllFields(docData, "[[" + fieldName + "]]");

                                                        if (fieldContent != "")
                                                        {
                                                            switch (fieldFormat.ToLower())
                                                            {
                                                                case "caps":
                                                                    fieldContent = fieldContent.First().ToString().ToUpper() + fieldContent.Substring(1);
                                                                    break;
                                                                case "camelcase":
                                                                    fieldContent = System.Globalization.CultureInfo.CurrentCulture.TextInfo.ToTitleCase(fieldContent);
                                                                    break;
                                                                case "lower":
                                                                    fieldContent = fieldContent.ToLower();
                                                                    break;
                                                                case "upper":
                                                                    fieldContent = fieldContent.ToUpper();
                                                                    break;
                                                                default:
                                                                    try
                                                                    {
                                                                        DateTime parsedDate;
                                                                        String pattern = "yyyy-MM-dd";
                                                                        if (DateTime.TryParseExact(fieldContent, pattern, null, System.Globalization.DateTimeStyles.None, out parsedDate))
                                                                            fieldContent = String.Format("{0:" + fieldFormat + "}", parsedDate);
                                                                        else
                                                                            fieldContent = String.Format("{0:" + fieldFormat + "}", fieldContent);
                                                                    }
                                                                    catch { }
                                                                    break;
                                                            }
                                                        }
                                                        documentContent = documentContent.Replace("[" + m.Value, fieldContent);
                                                    }
                                                }

                                                List<string> tables = new List<string>();

                                                String docTable = documentContent;
                                                while (docTable.IndexOf("[[##TABLE") > 0)
                                                {
                                                    int start = docTable.IndexOf("[[##TABLE");
                                                    int end = docTable.IndexOf("TABLE##]]");

                                                    if (end <= 0) break;

                                                    tables.Add(docTable.Substring(start + 9, end - start - 10));
                                                    docTable = docTable.Substring(end + 1);
                                                }

                                                var rtfBuilder = new RTFBuilder();
                                                List<String> tableNames = new List<string>();

                                                // replace any table definitions
                                                foreach (String table in tables)
                                                {
                                                    // Get Column Definitions
                                                    String groupDataItem = "";
                                                    String tableName = "";
                                                    List<CommonDefinitions.ColumnDefinition> colDefs = new List<CommonDefinitions.ColumnDefinition>();
                                                    List<CommonDefinitions.TotalColumns> totalCols = new List<CommonDefinitions.TotalColumns>();

                                                    foreach (String sColDef in table.Split('|'))
                                                    {
                                                        if (sColDef == table.Split('|')[0])
                                                        {
                                                            tableName = sColDef.Split(':')[0];
                                                            groupDataItem = sColDef.Split(':')[1];
                                                        }
                                                        else
                                                        {
                                                            if (sColDef.Split(':').Count() == 2)
                                                                colDefs.Add(new CommonDefinitions.ColumnDefinition(sColDef.Split(':')[0], sColDef.Split(':')[1]));
                                                            if (sColDef.Split(':').Count() == 3)
                                                            {
                                                                CommonDefinitions.ColumnDefinition colDef = new CommonDefinitions.ColumnDefinition(sColDef.Split(':')[0], sColDef.Split(':')[1], sColDef.Split(':')[2]);
                                                                colDefs.Add(colDef);
                                                                if (colDef.isTotalColumn == true)
                                                                {
                                                                    totalCols.Add(new CommonDefinitions.TotalColumns(tableName, (totalCols.Count + 1), colDef.fieldName));
                                                                }
                                                            }
                                                            if (sColDef.Split(':').Count() == 4)
                                                            {
                                                                CommonDefinitions.ColumnDefinition colDef = new CommonDefinitions.ColumnDefinition(sColDef.Split(':')[0], sColDef.Split(':')[1], sColDef.Split(':')[2], sColDef.Split(':')[3]);
                                                                colDefs.Add(colDef);
                                                                if (colDef.isTotalColumn == true)
                                                                {
                                                                    totalCols.Add(new CommonDefinitions.TotalColumns(tableName, (totalCols.Count + 1), colDef.fieldName));
                                                                }
                                                            }
                                                        }
                                                    }

                                                    tableNames.Add(tableName);

                                                    byte[] tableByteArray = Encoding.ASCII.GetBytes(documentContent);
                                                    MemoryStream tableStream = new MemoryStream(tableByteArray);
                                                    //Document tableDoc = new Document(tableStream);

                                                    Document tableDoc = new Document();
                                                    DocumentBuilder builder = new DocumentBuilder(tableDoc);

                                                    builder.Write("##" + tableName + "START##");
                                                    builder.Writeln();

                                                    Aspose.Words.Tables.Table tbl = builder.StartTable();

                                                    builder.CellFormat.Borders.Top.LineStyle = LineStyle.None;
                                                    builder.CellFormat.Borders.Bottom.LineStyle = LineStyle.Single;
                                                    builder.CellFormat.Borders.Right.LineStyle = LineStyle.None;
                                                    builder.CellFormat.Borders.Left.LineStyle = LineStyle.None;

                                                    builder.CellFormat.TopPadding = 4;
                                                    builder.CellFormat.BottomPadding = 4;
                                                    builder.CellFormat.LeftPadding = 0;

                                                    ParagraphFormat paragraphFormat = builder.ParagraphFormat;
                                                    paragraphFormat.Alignment = ParagraphAlignment.Left;
                                                    paragraphFormat.LeftIndent = 0;

                                                    builder.Font.Name = "Arial";
                                                    builder.Font.Size = 10;
                                                    builder.Font.Bold = true;

                                                    builder.Font.Bold = false;
                                                    Color headerFont = new Color();
                                                    headerFont = Color.FromArgb(39, 55, 96);
                                                    builder.Font.Color = headerFont;

                                                    foreach (CommonDefinitions.ColumnDefinition colDef in colDefs)
                                                    {

                                                        builder.InsertCell();
                                                        if (colDef == colDefs.First())
                                                        {
                                                            builder.StartBookmark(tableName + "START");
                                                            builder.EndBookmark(tableName + "START");
                                                        }
                                                        tbl.LeftIndent = -42;

                                                        tbl.AutoFit(Aspose.Words.Tables.AutoFitBehavior.FixedColumnWidths);
                                                        builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(colDef.columnWidth);

                                                        builder.Write(colDef.columnName);
                                                    }

                                                    builder.EndRow();

                                                    NBElement rowItems = docData;
                                                    foreach (String elementName in groupDataItem.Split('.'))
                                                    {
                                                        rowItems = rowItems[elementName];
                                                    }

                                                    builder.Font.Bold = false;
                                                    Color contentFont = new Color();
                                                    contentFont = Color.FromArgb(128, 128, 128);
                                                    builder.Font.Color = contentFont;

                                                    int rowId = 1;
                                                    //foreach (NBElement row in rowItems)
                                                    //{
                                                    foreach (ColumnDefinition colDef in colDefs)
                                                    {
                                                        builder.InsertCell();
                                                        tbl.AutoFit(Aspose.Words.Tables.AutoFitBehavior.FixedColumnWidths);
                                                        builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(colDef.columnWidth);

                                                        builder.Write("[[" + colDef.fieldName + "]]");

                                                    }
                                                    rowId++;
                                                    builder.EndRow();

                                                    foreach (ColumnDefinition colDef in colDefs)
                                                    {
                                                        builder.InsertCell();
                                                        tbl.AutoFit(Aspose.Words.Tables.AutoFitBehavior.FixedColumnWidths);
                                                        builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(colDef.columnWidth);

                                                        builder.Write("[[" + colDef.fieldName + "Last]]");
                                                    }

                                                    rowId++;
                                                    builder.EndRow();

                                                    builder.EndTable();

                                                    builder.Writeln();
                                                    builder.Write("##" + tableName + "END##");

                                                    MemoryStream rtfTableStream = new MemoryStream();
                                                    tableDoc.Save(rtfTableStream, SaveFormat.Rtf);

                                                    Byte[] rtfByteArray = rtfTableStream.ToArray();
                                                    String rtfDocString = Encoding.ASCII.GetString(rtfByteArray);

                                                    String tableRowString = rtfDocString.Substring(rtfDocString.IndexOf("}\\trowd\\irow1\\") + 1, rtfDocString.IndexOf("}\\trowd\\irow2\\") - rtfDocString.IndexOf("}\\trowd\\irow1\\"));

                                                    int insertIndex = rtfDocString.IndexOf(tableRowString) + tableRowString.Length;
                                                    //String tableContents = "";
                                                    StringBuilder tableContents = new StringBuilder(rtfDocString.Substring(0, insertIndex));

                                                    int rowCount = 2;
                                                    foreach (NBElement row in rowItems)
                                                    {
                                                        String rowString = tableRowString;
                                                        foreach (ColumnDefinition colDef in colDefs)
                                                        {
                                                            String sData = "";
                                                            if (colDef.fieldName != "lineNo")
                                                            {
                                                                if (row[colDef.fieldName] != null)
                                                                {
                                                                    sData = row[colDef.fieldName].ValueAsString;
                                                                    if (sData.ToUpper() == "NULL")
                                                                        sData = "";
                                                                }
                                                                else
                                                                    sData = colDef.defaultValue;
                                                            }
                                                            else
                                                                sData = (rowCount - 1).ToString();

                                                            if (rowCount < rowItems.Count() + 1)
                                                            {
                                                                rowString = rowString.Replace("[[" + colDef.fieldName + "]]", sData);
                                                            }
                                                            else
                                                            {
                                                                rtfDocString = rtfDocString.Replace("[[" + colDef.fieldName + "Last]]", sData);
                                                            }
                                                        }

                                                        if (rowCount < rowItems.Count() + 1)
                                                        {
                                                            rowString = rowString.Replace("irow1", "irow" + rowCount.ToString());
                                                            tableContents.Append(rowString);
                                                        }
                                                        rowCount++;
                                                    }

                                                    tableContents.Replace(tableRowString, "");
                                                    tableContents.Append(rtfDocString.Substring(insertIndex));

                                                    String tableOnly = tableContents.ToString();

                                                    //tableOnly = tableOnly.Substring(tableOnly.IndexOf(@"\trowd", tableOnly.Length - tableOnly.IndexOf(@"{\*\latentstyles")));
                                                    int startIndex = tableOnly.IndexOf("##" + tableName + "START##") + ("##" + tableName + "START##").Length;
                                                    int endIndex = tableOnly.IndexOf("##" + tableName + "END##");

                                                    tableOnly = tableOnly.Substring(startIndex, endIndex - startIndex);

                                                    documentContent = documentContent.Replace("[[##TABLE" + table + "|TABLE##]]", tableOnly);
                                                }


                                                // replace any items not found in the dataset
                                                Regex regex = new Regex(string.Format("\\[.*?\\]]"));
                                                documentContent = regex.Replace(documentContent, "");

                                                templateName = ReplaceDateTimeReferences(templateName);

                                                byte[] byteArray = Encoding.ASCII.GetBytes(documentContent);

                                                MemoryStream stream = new MemoryStream(byteArray);

                                                Document doc = new Document(stream);
                                                stream.Dispose();

                                                //String tempFileName = @"c:\temp\FileTest\" + Guid.NewGuid().ToString();
                                                //File.WriteAllText(tempFileName, documentContent);
                                                //Document doc = new Document(tempFileName);
                                                //File.Delete(tempFileName);

                                                foreach (String tableName in tableNames)
                                                {
                                                    Bookmark bm = doc.Range.Bookmarks[tableName + "START"];
                                                    Aspose.Words.Tables.Table t = (Aspose.Words.Tables.Table)bm.BookmarkStart.GetAncestor(NodeType.Table);

                                                    foreach (Aspose.Words.Tables.Cell c in t.GetChildNodes(NodeType.Cell, true))
                                                    {
                                                        foreach (Run run in c.GetChildNodes(NodeType.Run, true))
                                                        {
                                                            Aspose.Words.Font font = run.Font;
                                                            Color contentFont = new Color();
                                                            contentFont = Color.FromArgb(128, 128, 128);
                                                            font.Color = contentFont;
                                                            font.Name = "Arial";
                                                        }
                                                    }
                                                    foreach (Aspose.Words.Tables.Cell c in t.FirstRow.Cells)
                                                    {
                                                        foreach (Run run in c.GetChildNodes(NodeType.Run, true))
                                                        {
                                                            Aspose.Words.Font font = run.Font;
                                                            Color contentFont = new Color();
                                                            contentFont = Color.FromArgb(39, 55, 96);
                                                            font.Color = contentFont;
                                                            font.Name = "Arial";
                                                            font.Bold = true;
                                                        }
                                                    }
                                                }

                                                MemoryStream pdfStream = new MemoryStream();

                                                doc.Save(pdfStream, SaveFormat.Pdf);

                                                PDFStream pdfs = new PDFStream(pdfStream.ToArray(), templateName, identifier);
                                                pdfStream.Close();

                                                if (isZipUpload == "TRUE")
                                                    pdfStreams.Add(pdfs);
                                                else
                                                    nonZipPDFStreams.Add(pdfs);
                                                //}
                                            });
                                        }

                                        if (outputFileType == OutputFileTypes.CSV)
                                        {
                                            foreach (NBElement docData in elementsToProcess)
                                            {
                                                csvCollatedFile = docData["headerRow"].ValueAsString;
                                                csvCollatedFile += Environment.NewLine;

                                                Boolean isFirstRow = true;

                                                foreach (NBElement row in docData["rows"])
                                                {
                                                    if (isFirstRow == true)
                                                        isFirstRow = false;
                                                    else
                                                        csvCollatedFile += Environment.NewLine;

                                                    csvCollatedFile += row.ValueAsString;
                                                }
                                            }

                                            byte[] byteArray = Encoding.ASCII.GetBytes(csvCollatedFile);
                                            outputFileName = ReplaceDateTimeReferences(outputFileName);
                                            outputFileName = getUploadFileName(outputFileName, inputRoot);

                                            if (isZipUpload == "TRUE")
                                                pdfStreams.Add(new PDFStream(byteArray, outputFileName));
                                            else
                                                nonZipPDFStreams.Add(new PDFStream(byteArray, outputFileName));
                                        }
                                    }
                                }

                                if (pdfStreams.Count() > 0)
                                {
                                    byte[] zipContent = writeContentToZip(pdfStreams, outputRoot);
                                    writeFileToBlobStore(outputRoot, zipContent, inputRoot, requestId);
                                }

                                if (nonZipPDFStreams.Count() > 0)
                                {
                                    writeMultiFilesToBlobStore(outputRoot, nonZipPDFStreams, inputRoot, requestId);
                                    outputRoot["JSON"]["Data"].CreateFirstChild("requestId").SetValue(requestId);
                                }
                            }
                        }
                    }
                    else
                    {
                        outputRoot["JSON"]["Data"].DeleteAllChildren();
                        checkBlobStore(inputRoot, outputRoot);
                    }
                }
                catch (Exception e)
                {
                    handleError("PDF10000", e, outputRoot, "General Error in PDF Generation Process");
                }

                #endregion UserCode
                // Change the following if not propagating message to the 'Out' terminal
                outTerminal.Propagate(outAssembly);
            }
        }

        static string GetRtfUnicodeEscapedString(string s)
        {
            var sb = new StringBuilder();
            foreach (var c in s)
            {
                if (c == '\\' || c == '{' || c == '}')
                    sb.Append(@"\" + c);
                else if (c <= 0x7f)
                    sb.Append(c);
                else
                    sb.Append("\\u" + Convert.ToUInt32(c) + "?");
            }
            return sb.ToString();
        }

        private string ReplaceAllFields(NBElement docData, String templateString, String prefix = "", Boolean isCSV = false)
        {
            foreach (NBElement data in docData)
            {
                String sData = data.ValueAsString;
                if (sData.ToUpper() == "NULL")
                    sData = "";

                if (isCSV == true && sData.Contains(","))
                {
                    sData = "\"" + sData + "\"";
                }

                if (ContainsUnicodeCharacter(sData))
                    sData = GetRtfUnicodeEscapedString(sData);

                templateString = templateString.Replace("[[" + prefix + data.Name.ToString() + "]]", sData);
                templateString = templateString.Replace("[" + prefix + data.Name.ToString() + "]", sData);

                if (data.ElementType != NBParsers.JSON.Array)
                {
                    if (data.Children().Count() > 0)
                    {
                        templateString = ReplaceAllFields(data, templateString, prefix + data.Name.ToString() + ".");
                    }
                }
            }
            return templateString;
        }

        public bool ContainsUnicodeCharacter(string input)
        {

            return Encoding.ASCII.GetByteCount(input) != Encoding.UTF8.GetByteCount(input);

            //const int MaxAnsiCode = 255;
            //return input.Any(c => c > MaxAnsiCode);
            //return System.Text.ASCIIEncoding.GetEncoding(0).GetString(System.Text.ASCIIEncoding.GetEncoding(0).GetBytes(input)) != input;

        }

        private Double getRecursiveTotal(NBElement docData, String fieldName)
        {
            Double total = 0;
            foreach (NBElement child in docData)
            {
                if (child.Children().Count() > 0)
                    total += getRecursiveTotal(child, fieldName);
                else
                    if (child.Name == fieldName)
                    total += Convert.ToDouble(child.ValueAsString);
            }
            return total;
        }

        private void applyLicense(NBElement outputRoot)
        {
            try
            {
                string License = "Aspose.Total.lic";
                var license = new License();
                license.SetLicense(License);
            }
            catch (Exception e)
            {
                handleError("PDF10001", e, outputRoot, "Error applying license");
            }
        }

        private String ReplaceDateTimeReferences(String templateName)
        {
            templateName = templateName.Replace("yyyy", DateTime.Now.Year.ToString())
                                        .Replace("MM", DateTime.Now.Month.ToString().PadLeft(2, '0'))
                                        .Replace("dd", DateTime.Now.Day.ToString().PadLeft(2, '0'))
                                        .Replace("HH", DateTime.Now.Hour.ToString().PadLeft(2, '0'))
                                        .Replace("mm", DateTime.Now.Minute.ToString().PadLeft(2, '0'))
                                        .Replace("ss", DateTime.Now.Second.ToString().PadLeft(2, '0'));

            return templateName;
        }

        private byte[] writeContentToZip(List<PDFStream> pdfStreams, NBElement outputRoot)
        {
            byte[] zipContent = null;
            List<String> fileNames = new List<String>();
            try
            {
                using (var outStream = new MemoryStream())
                {
                    using (var archive = new ZipArchive(outStream, ZipArchiveMode.Create, true))
                    {
                        foreach (PDFStream pdfs in pdfStreams)
                        {
                            if (fileNames.Where(fn => fn == pdfs.fileName).Count() == 0)
                            {
                                fileNames.Add(pdfs.fileName);
                                ZipArchiveEntry zipItem = archive.CreateEntry(pdfs.fileName);
                                using (System.IO.MemoryStream originalFileMemoryStream = new System.IO.MemoryStream(pdfs.stream))
                                {
                                    using (System.IO.Stream entryStream = zipItem.Open())
                                    {
                                        originalFileMemoryStream.CopyTo(entryStream);
                                    }
                                }
                            }
                        }
                    }
                    zipContent = outStream.ToArray();
                }
            }
            catch (Exception e)
            {
                handleError("PDF10003", e, outputRoot, "Error writing content to zip file");
            }
            return zipContent;
        }

        private void handleError(String errorCode, Exception e, NBElement outputRoot, String errorDescription)
        {
            isError = true;
            outputRoot["JSON"]["Data"].CreateLastChild("errorCode").SetValue(errorCode);
            outputRoot["JSON"]["Data"].CreateLastChild("error").SetValue(e.ToString());
            outputRoot["JSON"]["Data"].CreateLastChild("errorDesc").SetValue(errorDescription);
            outputRoot["JSON"]["Data"].CreateLastChild("requestId").SetValue(requestId);
            outputRoot["JSON"]["Data"].CreateLastChild("errorDate").SetValue(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
        }

        private BlobStorageHelper createBlobStoreHelper(NBElement inputRoot)
        {
            NBElement endpointLocation = inputRoot["JSON"]["Data"]["endpointLocation"];

            string connectionString = endpointLocation["endpointConnString"].ValueAsString;
            string containerRef = endpointLocation["containerRef"].ValueAsString;

            var blobHelper = new BlobStorageHelper(connectionString, containerRef);
            return blobHelper;
        }

        private void checkBlobStore(NBElement inputRoot, NBElement outputRoot)
        {
            try
            {
                var blobHelper = createBlobStoreHelper(inputRoot);
                if (blobHelper.checkConnection() == false)
                {
                    Exception e = new Exception("Container does not exist in Blob Store");
                    handleError("PDF10012", e, outputRoot, "Container does not exist in Blob Store");
                }
            }
            catch (Exception e)
            {
                handleError("PDF10002", e, outputRoot, "Error connecting to blob store");
            }
        }
        private void writeMultiFilesToBlobStore(NBElement outputRoot, List<PDFStream> pdfStreams, NBElement inputRoot, String requestId)
        {
            try
            {
                var blobHelper = createBlobStoreHelper(inputRoot);

                List<multiDocument> docUrls = blobHelper.uploadMultipleDocuments(pdfStreams, requestId);

                var docNode = outputRoot["JSON"]["Data"].CreateLastChild(NBParsers.JSON.Array, "documentURLs", "");

                foreach (multiDocument md in docUrls)
                {
                    var refOut = docNode;

                    refOut = refOut.CreateLastChild("Document");
                    //refOut.CreateLastChild("url").SetValue(url);

                    refOut.CreateLastChild("url").SetValue(md.url);
                    refOut.CreateLastChild("identifier").SetValue(md.identifier);
                }
            }
            catch (Exception e)
            {
                handleError("PDF10002", e, outputRoot, "Error writing file(s) to blob store");
            }
        }
        private void getUploadFileLocation(NBElement outputRoot, String templateName, NBElement inputRoot, String requestId)
        {
            try
            {
                var blobHelper = createBlobStoreHelper(inputRoot);
                string url = blobHelper.getFolderLocation(requestId);

                String fileName = getUploadFileName(templateName, inputRoot);
                url = url + fileName;

                outputRoot["JSON"]["Data"].CreateFirstChild("downloadURL").SetValue(url);
                outputRoot["JSON"]["Data"].CreateFirstChild("downloadFileName").SetValue(fileName);
            }
            catch (Exception e)
            {
                handleError("PDF10002", e, outputRoot, "Error connecting to blob store");
            }

        }
        private String getUploadFileName(String templateName, NBElement inputRoot)
        {
            String fileName;

            if (inputRoot["JSON"]["Data"]["downloadFileName"] != null)
            {
                fileName = inputRoot["JSON"]["Data"]["downloadFileName"].ValueAsString;
            }
            else
            {
                if (templateName != "")
                    fileName = templateName;
                else
                    fileName = "IPO-documents-" + DateTime.Now.ToString("yyyy-MM-dd-HHmmss") + ".zip";
            }
            return fileName;
        }

        private void writeFileToBlobStore(NBElement outputRoot, Byte[] fileContent, NBElement inputRoot, String requestId, String templateName = "")
        {
            try
            {
                var blobHelper = createBlobStoreHelper(inputRoot);

                String fileName = getUploadFileName(templateName, inputRoot);

                String url = blobHelper.uploadDocument(fileContent, fileName, requestId);

                outputRoot["JSON"]["Data"].CreateFirstChild("downloadURL").SetValue(url);

            }
            catch (Exception e)
            {
                handleError("PDF10002", e, outputRoot, "Error writing file to blob store");
            }
        }




        private String getFieldContent(NBElement docData, String ifField)
        {
            if (ifField.Split('|').Length != 3)
                return "";

            foreach (NBElement data in docData)
            {
                //[[IF{depositAccount}!=""|Deposit Account:|]]
                String sData = data.ValueAsString;
                if (sData.ToUpper() == "NULL")
                    sData = "";

                string ifStatament = ifField.Substring(0, ifField.IndexOf("|"));
                string values = ifField.Substring(ifField.IndexOf("|"));

                ifStatament = ifStatament.Replace("{" + data.Name + "}", "'" + sData + "'");
                values = values.Replace("{" + data.Name + "}", sData);

                ifField = ifStatament + values;

            }

            Regex regex = new Regex(@"\{([^\}]+)\}");

            string finalIfStatament = ifField.Substring(0, ifField.IndexOf("|"));
            string finalValues = ifField.Substring(ifField.IndexOf("|"));

            finalIfStatament = regex.Replace(finalIfStatament, "''");
            finalValues = regex.Replace(finalValues, "");

            ifField = finalIfStatament + finalValues;

            try
            {
                List<String> ifList = ifField.Split('|').ToList();
                ifList[0] = ifList[0].ToString().Substring(2, ifList[0].Length - 2);

                if (ifList[0].ToString().IndexOf("{") == -1)
                {
                    System.Data.DataTable dt = new System.Data.DataTable();
                    var val = dt.Compute(ifList[0], "");
                    if (Convert.ToBoolean(val))
                        return ifList[1].ToString();
                    else
                        return ifList[2].ToString();
                }
                else
                    return "";
            }
            catch
            {
                return "";
            }
        }
    }
}