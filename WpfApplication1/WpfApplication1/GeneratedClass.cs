//TODO - replace with usings from the tool. 

//Keep this namespace
using DocumentFormat.OpenXml;
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
            string filename = "ETMetaData.xlsx";
            //  using (var workbook = SpreadsheetDocument.Create(destination, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook))
            using (var workbook = SpreadsheetDocument.Create( memoryStream, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook, true))
            {
                var workbookPart = workbook.AddWorkbookPart();
                workbook.WorkbookPart.Workbook = new Workbook();
                workbook.WorkbookPart.Workbook.Sheets = new Sheets();


                var sheetPart = workbook.WorkbookPart.AddNewPart<WorksheetPart>();
                var sheetData = new SheetData();
                sheetPart.Worksheet = new Worksheet(sheetData);

                Sheets sheets = workbook.WorkbookPart.Workbook.GetFirstChild<Sheets>();
                string relationshipId = workbook.WorkbookPart.GetIdOfPart(sheetPart);
                AddStyle(workbookPart, relationshipId);
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

    }
}
