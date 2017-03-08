//TODO - replace with usings from the tool. 

//Keep this namespace
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2010.Excel;
using DocumentFormat.OpenXml.Office2013.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace WpfGenerateExcel
{
    //TODO - replace contents with class generated using the 
    //Open XML Productivity Tool.  Yes, the entire class!
    public class GeneratedClass
    {
        public void CreatePackage(string filePath, List<ExcelData> excelData)
        {
            string[] columns = { "Barcode", "Client", "Content_Owner", "IsClientVisible",
                   "Material_Type", "Client_Barcode", "MaterialId","Standard",
                   "Class", //"Subtitle_Language", "Subtitle_Type", 
                   "Title_Name",
                   "Language", "StartOfMessage", "EndOfMessage", "StartOfRecording",
                   "EndOfRecording", "Audio Channels"};
            MemoryStream memoryStream = new MemoryStream();
            //string filename = "ETMetaData.xlsx";
            //  using (var workbook = SpreadsheetDocument.Create(destination, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook))
            using (var workbook = SpreadsheetDocument.Create( filePath, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook, true))
            {
                var workbookPart = workbook.AddWorkbookPart();
                workbook.WorkbookPart.Workbook = new Workbook();
                workbook.WorkbookPart.Workbook.Sheets = new Sheets();


                var sheetPart = workbook.WorkbookPart.AddNewPart<WorksheetPart>();
                var sheetData = new SheetData();
                sheetPart.Worksheet = new Worksheet(sheetData);

                Sheets sheets = workbook.WorkbookPart.Workbook.GetFirstChild<Sheets>();
                string relationshipId = workbook.WorkbookPart.GetIdOfPart(sheetPart);

                //AddStyle(workbookPart, relationshipId);
                WorkbookStylesPart workbookStylesPart = workbook.WorkbookPart.AddNewPart<WorkbookStylesPart>("rIdStyles");
                GenerateWorkbookStylesPart1Content(workbookStylesPart);


                uint sheetId = 1;
                if (sheets.Elements<Sheet>().Count() > 0)
                {
                    sheetId =
                        sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
                }

                Sheet sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = "Element Profiles" };
                sheets.Append(sheet);

                Row headerRow = new Row();

                foreach (var column in columns)
                {
                    Cell cell = new Cell();
                    cell.DataType = CellValues.String;
                    cell.CellValue = new CellValue(column);
                    headerRow.AppendChild(cell);
                }

                sheetData.AppendChild(headerRow);

                foreach (ExcelData data in excelData)
                {
                    int i = 0;
                    Row newRow = new Row();
                    CreateCellForValue(columns[0], data.Barcode, newRow);
                    CreateCellForValue(columns[1], data.Client, newRow);
                    CreateCellForValue(columns[2], data.Content_Owner, newRow);
                    CreateCellForValue(columns[3], data.IsClientVisible.ToString(), newRow);
                    CreateCellForValue(columns[4], data.Material_Type, newRow);
                    CreateCellForValue(columns[5], data.Client_Barcode, newRow);
                    CreateCellForValue(columns[6], data.MaterialId, newRow);
                    CreateCellForValue(columns[7], data.Standard, newRow);
                    CreateCellForValue(columns[8], data.Class, newRow);
                    CreateCellForValue(columns[9], data.Title_Name, newRow);
                    CreateCellForValue(columns[10], data.Language, newRow);
                    //if (data.TimecodeItem != null)
                    //{
                    //    CreateCellForValue(columns[11], data.TimecodeItem[i].StartOfMessage, newRow);
                    //    CreateCellForValue(columns[12], data.TimecodeItem[i].EndOfMessage, newRow);
                    //    CreateCellForValue(columns[13], data.TimecodeItem[i].StartOfRecording, newRow);
                    //    CreateCellForValue(columns[14], data.TimecodeItem[i].EndOfRecording, newRow);
                    //}
                    //else
                    //{
                    //    CreateCellForValue(columns[11], string.Empty, newRow);
                    //    CreateCellForValue(columns[12], string.Empty, newRow);
                    //    CreateCellForValue(columns[13], string.Empty, newRow);
                    //    CreateCellForValue(columns[14], string.Empty, newRow);
                    //}
                    //CreateCellForValue(columns[15], ((data.AudioChannels != null) ? string.Join("|", data.AudioChannels.Select(x => x.ToString())) : ""), newRow);
                    ////CreateCellForValue(columns[18], data., newRow);
                    //  i++;
                    sheetData.AppendChild(newRow);

                }

                workbook.WorkbookPart.Workbook.Save();
            }
            //memoryStream.Flush();
            //memoryStream.Position = 0;

            //Response.ClearContent();
            //Response.Clear();
            //Response.Buffer = true;
            //Response.Charset = "";

            ////  NOTE: If you get an "HttpCacheability does not exist" error on the following line, make sure you have
            ////  manually added System.Web to this project's References.

            //Response.Cache.SetCacheability(System.Web.HttpCacheability.NoCache);
            //Response.AddHeader("content-disposition", "attachment; filename=" + filename);
            //Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            //byte[] data1 = new byte[memoryStream.Length];
            //memoryStream.Read(data1, 0, data1.Length);
            //memoryStream.Close();
            //Response.BinaryWrite(data1);
            //Response.Flush();

            ////  Feb2015: Needed to replace "Response.End();" with the following 3 lines, to make sure the Excel was fully written to the Response
            //System.Web.HttpContext.Current.Response.Flush();
            //System.Web.HttpContext.Current.Response.SuppressContent = true;
            //System.Web.HttpContext.Current.ApplicationInstance.CompleteRequest();
        }

        private void CreateCellForValue(string column, string columnValue, Row newRow, bool isLock = false)
        {
            Cell cell = new Cell();
            cell.DataType = CellValues.String;
            cell.CellValue = new CellValue(columnValue);
            if (isLock)
                cell.StyleIndex = 0;
            newRow.AppendChild(cell);
        }

        private string GetCellValue(SpreadsheetDocument doc, Cell cell)
        {
            if (cell == null || cell.CellValue == null)
                return null;

            string value = cell.CellValue.InnerText;
            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                return doc.WorkbookPart.SharedStringTablePart.SharedStringTable.ChildElements.GetItem(int.Parse(value)).InnerText;
            }
            return value;
        }

        private void AddStyle(WorkbookPart workbookPart, string id)
        {
            CellFormat lockFormat = new CellFormat() { ApplyProtection = true, Protection = new Protection() { Locked = true } };
            WorkbookStylesPart sp = workbookPart.GetPartsOfType<WorkbookStylesPart>().FirstOrDefault();

            if (sp == null)
                sp = workbookPart.AddNewPart<WorkbookStylesPart>();
            sp.Stylesheet = new Stylesheet();
            sp.Stylesheet.CellFormats = new CellFormats();
            sp.Stylesheet.CellFormats.AppendChild<CellFormat>(lockFormat);
            sp.Stylesheet.CellFormats.Count = UInt32Value.FromUInt32((uint)sp.Stylesheet.CellFormats.ChildElements.Count);
            sp.Stylesheet.Save();
        }

        private static void GenerateWorkbookStylesPart1Content(WorkbookStylesPart workbookStylesPart1)
        {
            Stylesheet stylesheet1 = new Stylesheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac x16r2" } };
            stylesheet1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            stylesheet1.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
            stylesheet1.AddNamespaceDeclaration("x16r2", "http://schemas.microsoft.com/office/spreadsheetml/2015/02/main");

            Fonts fonts1 = new Fonts() { Count = (UInt32Value)1U, KnownFonts = true };

            Font font1 = new Font();
            FontSize fontSize1 = new FontSize() { Val = 11D };
            DocumentFormat.OpenXml.Spreadsheet.Color color1 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Theme = (UInt32Value)1U };
            FontName fontName1 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering1 = new FontFamilyNumbering() { Val = 2 };
            FontScheme fontScheme1 = new FontScheme() { Val = FontSchemeValues.Minor };

            font1.Append(fontSize1);
            font1.Append(color1);
            font1.Append(fontName1);
            font1.Append(fontFamilyNumbering1);
            font1.Append(fontScheme1);

            fonts1.Append(font1);

            Fills fills1 = new Fills() { Count = (UInt32Value)2U };

            Fill fill1 = new Fill();
            PatternFill patternFill1 = new PatternFill() { PatternType = PatternValues.None };

            fill1.Append(patternFill1);

            Fill fill2 = new Fill();
            PatternFill patternFill2 = new PatternFill() { PatternType = PatternValues.Gray125 };

            fill2.Append(patternFill2);

            fills1.Append(fill1);
            fills1.Append(fill2);

            Borders borders1 = new Borders() { Count = (UInt32Value)1U };

            Border border1 = new Border();
            LeftBorder leftBorder1 = new LeftBorder();
            RightBorder rightBorder1 = new RightBorder();
            TopBorder topBorder1 = new TopBorder();
            BottomBorder bottomBorder1 = new BottomBorder();
            DiagonalBorder diagonalBorder1 = new DiagonalBorder();

            border1.Append(leftBorder1);
            border1.Append(rightBorder1);
            border1.Append(topBorder1);
            border1.Append(bottomBorder1);
            border1.Append(diagonalBorder1);

            borders1.Append(border1);

            CellStyleFormats cellStyleFormats1 = new CellStyleFormats() { Count = (UInt32Value)1U };
            CellFormat cellFormat1 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U };

            cellStyleFormats1.Append(cellFormat1);

            CellFormats cellFormats1 = new CellFormats() { Count = (UInt32Value)1U };
            CellFormat cellFormat2 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U };

            cellFormats1.Append(cellFormat2);

            CellStyles cellStyles1 = new CellStyles() { Count = (UInt32Value)1U };
            CellStyle cellStyle1 = new CellStyle() { Name = "Normal", FormatId = (UInt32Value)0U, BuiltinId = (UInt32Value)0U };

            cellStyles1.Append(cellStyle1);
            DocumentFormat.OpenXml.Office2010.Excel.DifferentialFormats differentialFormats1 = new DocumentFormat.OpenXml.Office2010.Excel.DifferentialFormats() { Count = (UInt32Value)0U };
            TableStyles tableStyles1 = new TableStyles() { Count = (UInt32Value)0U, DefaultTableStyle = "TableStyleMedium2", DefaultPivotStyle = "PivotStyleLight16" };

            StylesheetExtensionList stylesheetExtensionList1 = new StylesheetExtensionList();

            StylesheetExtension stylesheetExtension1 = new StylesheetExtension() { Uri = "{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" };
            stylesheetExtension1.AddNamespaceDeclaration("x14", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");
            SlicerStyles slicerStyles1 = new SlicerStyles() { DefaultSlicerStyle = "SlicerStyleLight1" };

            stylesheetExtension1.Append(slicerStyles1);

            StylesheetExtension stylesheetExtension2 = new StylesheetExtension() { Uri = "{9260A510-F301-46a8-8635-F512D64BE5F5}" };
            stylesheetExtension2.AddNamespaceDeclaration("x15", "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main");
            TimelineStyles timelineStyles1 = new TimelineStyles() { DefaultTimelineStyle = "TimeSlicerStyleLight1" };

            stylesheetExtension2.Append(timelineStyles1);

            stylesheetExtensionList1.Append(stylesheetExtension1);
            stylesheetExtensionList1.Append(stylesheetExtension2);

            stylesheet1.Append(fonts1);
            stylesheet1.Append(fills1);
            stylesheet1.Append(borders1);
            stylesheet1.Append(cellStyleFormats1);
            stylesheet1.Append(cellFormats1);
            stylesheet1.Append(cellStyles1);
            stylesheet1.Append(differentialFormats1);
            stylesheet1.Append(tableStyles1);
            stylesheet1.Append(stylesheetExtensionList1);

            workbookStylesPart1.Stylesheet = stylesheet1;
        }


    }
}

/*
 *  public void CreatePackage(string filePath)
        {
            using (SpreadsheetDocument package = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook))
            {
                WriteExcelFile(GenerateDataSet(), package);
            }
        }

        private DataSet GenerateDataSet()
        {
            DataTable table1 = new DataTable("patients");
            table1.Columns.Add("name");
            table1.Columns.Add("id");
            table1.Rows.Add("sam", 1);
            table1.Rows.Add("mark", 2);

            DataTable table2 = new DataTable("medications");
            table2.Columns.Add("id");
            table2.Columns.Add("medication");
            table2.Rows.Add(1, "atenolol");
            table2.Rows.Add(2, "amoxicillin");

            // Create a DataSet and put both tables in it.
            DataSet set = new DataSet("office");
            set.Tables.Add(table1);
            set.Tables.Add(table2);

            return set;

        }

        private static void WriteExcelFile(DataSet ds, SpreadsheetDocument spreadsheet)
        {

            spreadsheet.AddWorkbookPart();
            spreadsheet.WorkbookPart.Workbook = new DocumentFormat.OpenXml.Spreadsheet.Workbook();

            spreadsheet.WorkbookPart.Workbook.Append(new BookViews(new WorkbookView()));

            WorkbookStylesPart workbookStylesPart = spreadsheet.WorkbookPart.AddNewPart<WorkbookStylesPart>("rIdStyles");

            GenerateWorkbookStylesPart1Content(workbookStylesPart);
            uint worksheetNumber = 1;
            Sheets sheets = spreadsheet.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());
            foreach (DataTable dt in ds.Tables)
            {
                string worksheetName = dt.TableName;
                WorksheetPart newWorksheetPart = spreadsheet.WorkbookPart.AddNewPart<WorksheetPart>();
                Sheet sheet = new Sheet() { Id = spreadsheet.WorkbookPart.GetIdOfPart(newWorksheetPart), SheetId = worksheetNumber, Name = worksheetName };
                newWorksheetPart.Worksheet = new Worksheet(new SheetViews(new SheetView() { WorkbookViewId = 0, RightToLeft = true }), new SheetData());
                newWorksheetPart.Worksheet.Save();

                sheets.Append(sheet);

                WriteDataTableToExcelWorksheet(dt, newWorksheetPart);

                worksheetNumber++;
            }
            spreadsheet.WorkbookPart.Workbook.Save();
        }

    */