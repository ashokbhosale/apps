using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using System.IO;
using System.Globalization;
using System.Text.RegularExpressions;

namespace Meridium.Excel
{
    public class OpenXmlClient
    {        

        public OpenXmlClient()
        {

        }

        #region "Functionality Similar to OLE DB" 
        public DataTable GetSchemaTable(string fileName)
        {

            DataTable schematbl = new DataTable();
            schematbl.Columns.Add("TABLE_NAME", Type.GetType("System.String"));
            if (!System.IO.File.Exists(fileName))  // if file itself not exists return blank table
                return schematbl;

            DataRow row = null;
            using (SpreadsheetDocument spreadSheetDocument = SpreadsheetDocument.Open(fileName, false))
            {

                WorkbookPart workbookPart = spreadSheetDocument.WorkbookPart;
                IEnumerable<Sheet> sheets = spreadSheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();
                if (sheets.ToList().Count > 0)
                {
                    sheets.ToList().ForEach(sheet => {
                        row = schematbl.NewRow();
                        row[0] = sheet.Name;
                        schematbl.Rows.Add(row);
                        schematbl.AcceptChanges();
                    });
                }
            }


            return schematbl;
        }

        public int ExecuteNonQuery(string query, string[] param, string filePath)
        {
            //first see query type
            //create /insert /delete /update

            if (query.ToUpper().Contains("CREATE"))
            {
                CreateTableInsideExcel(query, param, filePath);
            }
            else if (query.ToUpper().Contains("INSERT"))
            {
                InsertRecordInTableInsideExcel(query, param, filePath);
            }
            else if (query.ToUpper().Contains("DELETE"))
            {
                DeleteWorkSheetRows(query, param, filePath);
            }
            else
            {
                throw new Exception("Not Implemented");
            }


            return 1;
        }

        //for multiple insert statement
        public int ExecuteMultipleQuery(string filePath, string[] query)
        {

            MutipleInsertRecordInsideExcel(query, filePath);

            return 1;
        }
        public DataSet Fill(string query, string[] param, string filePath)
        {
            DataSet ds = new DataSet();
            string tablename = GetTableNameFromSelectStatement(query);
            DataTable dt = ReadWorkSheetFromExcel(filePath, tablename);
            dt.TableName = "Table";
            ds.Tables.Add(dt);
            return ds;
        }

        public DataSet FillWithName(string query, string[] param, string filePath)
        {
            DataSet ds = new DataSet();
            return ds;
        }

        #endregion


        #region "CREATE INSERT DELETE OPERATION IN EXCEL"

        public bool CreateTableInsideExcel(string query, string[] param, string filePath)
        {

            DataTable dt = GetDataTableFromCreateStatement(query, param);

            //  For each worksheet you want to create
            string worksheetName = dt.TableName;

            AddNewSpreadSheet(filePath, worksheetName);

            AddHeaderInWorkSheet(filePath, dt);

            return true;
        }

        private bool InsertRecordInTableInsideExcel(string query, string[] param, string filePath)
        {
            // create datatable from insert statement.
            DataTable tbl = GetDataTableFromInsertStatement(query, param);

            //first existing datatable for respective sheet 
            // append current insert statement data to existing 
            // send final datatable for saving.
            DataSet ds = Fill("SELECT * FROM " + tbl.TableName, null, filePath);
            DataTable dsFinal = ds.Tables[0];

            bool firstRecordisBlank = IsFirstRecordIsBlank(dsFinal);
            if (firstRecordisBlank)
                dsFinal.Rows.RemoveAt(0);  // remove blank row.

            //add new row with value coming with insert statement.
            DataRow drrow = dsFinal.NewRow();
            for (int i = 0; i < tbl.Columns.Count; i++)
            {
                if (Convert.ToString(tbl.Rows[0][i]) == null)
                    drrow[i] = "";
                else
                    drrow[i] = Convert.ToString(tbl.Rows[0][i]);
            }
            dsFinal.Rows.Add(drrow);
            dsFinal.AcceptChanges();

            ProcessWorkSheet(filePath, dsFinal);
            return true;
        }

        private static bool IsFirstRecordIsBlank(DataTable dsFinal)
        {
            bool firstRecordisBlank = true;
            for (int i = 0; i < dsFinal.Columns.Count; i++)
            {
                if (Convert.ToString(dsFinal.Rows[0][i]) != "")
                {
                    firstRecordisBlank = false;
                    break;
                }
            }

            return firstRecordisBlank && dsFinal.Rows.Count <= 1;
        }

        private bool DeleteWorkSheetRows(string query, string[] param, string filePath)
        {

            //get table name
            string tablename = GetTableNameFromDeleteStatement(query, param);

            //get dataset
            DataSet ds = Fill("SELECT * FROM " + tablename, null, filePath);
            DataTable dsFinal = ds.Tables[0];


            bool firstRecordisBlank = IsFirstRecordIsBlank(dsFinal); // if first record is blank 
            if (!firstRecordisBlank)
            {
                //clear rows
                dsFinal.Rows.Clear();

                // add blank row 
                dsFinal = AddBlankRowToTable(dsFinal);

                ProcessWorkSheet(filePath, dsFinal);
            }
            return true;
        }


        private bool MutipleInsertRecordInsideExcel(string[] query, string filePath)
        {
            //prepare datatable from multiple insert statement
            DataTable tbl = GetDataTableFromMutipleInsertStatement(query);
            ProcessWorkSheet(filePath, tbl);
            return true;
        }

        #endregion


        #region "Parse Query"
        private DataTable GetDataTableFromCreateStatement(string query, string[] param)
        {
            DataTable dt = new DataTable();
            //traverse throw the query and create table from it.
            string tbl = query.ToString().Substring(query.IndexOf("CREATE TABLE") + 13, query.IndexOf("("));
            tbl = tbl.Split('(')[0].Trim();
            string[] fields = ((query.ToString().Substring(query.IndexOf("(") + 1)).Split(')')[0].Trim()).Split(',');

            dt.TableName = tbl.Replace("[", "").Replace("]", "");
            for (int i = 0; i < fields.Length; i++)
            {
                string[] fieldDetail = fields[i].Trim().Split(' '); // seperating column  name and datatype.
                dt.Columns.Add(fieldDetail[0].Replace("[", "").Replace("]", ""));  // adding column detail
            }

            return dt;
        }

        private DataTable GetDataTableFromInsertStatement(string query, string[] param)
        {
            DataTable dt = new DataTable();
            DataRow drrow = dt.NewRow();

            query = query.ToUpper();
            //traverse throw the query and create table from it.
            string tbl = query.ToString().Substring(query.IndexOf("INSERT INTO") + 11, query.IndexOf("("));
            tbl = tbl.Split('(')[0].Trim();
            if (tbl.IndexOf("$") > 0)
                tbl = tbl.Substring(0, tbl.IndexOf("$"));

            string[] fields = ((query.ToString().Substring(query.IndexOf("(") + 1)).Split(')')[0].Trim()).Split(',');

            dt.TableName = tbl.Replace("[", "").Replace("]", "");
            for (int i = 0; i < fields.Length; i++)
            {
                dt.Columns.Add(fields[i].Trim().Replace("[", "").Replace("]", ""));  // adding column detail
            }

            string field = string.Empty;
            if (param == null)
            {
                field = query.ToString().Substring(query.LastIndexOf("VALUES(") + 7);
                field = field.Substring(0, field.LastIndexOf(")"));
                field = field.Replace("''", @"""");
                param = new string[] { };
                param = field.Replace("',", "~").Split('~');
            }
            for (int i = 0; i < param.Length; i++)
            {
                if (Convert.ToString(param[i]) == null)
                    drrow[i] = "";
                else
                    drrow[i] = Convert.ToString(param[i]).Trim().Replace("'", "");
            }

            dt.Rows.Add(drrow);
            dt.AcceptChanges();
            return dt;
        }


        private string GetTableNameFromDeleteStatement(string query, string[] param)
        {
            query = query.ToUpper();
            string tbl = query.ToString().Substring(query.IndexOf("FROM") + 4);
            tbl = tbl.Trim().Replace("[", "").Replace("]", "");
            if (tbl.IndexOf("$") > 0)
                tbl = tbl.Substring(0, tbl.IndexOf("$"));

            return tbl;
        }


        private DataTable GetDataTableFromMutipleInsertStatement(string[] query)
        {
            DataTable dt = null;

            if (query.Length >= 1)
                dt = GetDataTableFromInsertStatement(query[0], null);

            if (query.Length > 2)
            {
                DataTable tmpDt;
                for (int i = 1; i < query.Length; i++)
                {
                    tmpDt = GetDataTableFromInsertStatement(query[i], null);
                    dt.ImportRow(tmpDt.Rows[0]);
                }
            }

            return dt; //it contain all insert statement records
        }

        private string GetTableNameFromSelectStatement(string query)
        {

            query = query.ToUpper();
            //traverse throw the query and create table from it.
            string tablename = query.ToString().Substring(query.IndexOf("FROM") + 4);
            tablename = tablename.Trim().Replace("[", "").Replace("]", "").Replace("$", "");

            return tablename;
        }

        #endregion


        #region "Create SpreadSheet"
        public static bool isSheetsAlreadyExists(Sheets sheets, string sheetName)
        {

            foreach (Sheet sheet in sheets)
            {
                if (sheet.Name == sheetName)
                    return true;
            }

            return false;
        }

        public static void AppendSpreadSheet(string filePath, string sheetName)
        {
            if (!System.IO.File.Exists(filePath))
                throw new Exception("File Not Exists");

            using (Stream stream = File.Open(filePath, FileMode.Open))
            {
                // Open a SpreadsheetDocument based on a stream.
                SpreadsheetDocument spreadsheetDocument =
                        SpreadsheetDocument.Open(stream, true);


                Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>();
                if (isSheetsAlreadyExists(sheets, sheetName))
                {
                    spreadsheetDocument.Close();
                    stream.Close();
                    return;
                }

                // Add a new worksheet.
                WorksheetPart newWorksheetPart = spreadsheetDocument.WorkbookPart.AddNewPart<WorksheetPart>();
                newWorksheetPart.Worksheet = new Worksheet(new SheetData());
                newWorksheetPart.Worksheet.Save();

                string relationshipId = spreadsheetDocument.WorkbookPart.GetIdOfPart(newWorksheetPart);

                // Get a unique ID for the new worksheet.
                uint sheetId = 1;
                if (sheets.Elements<Sheet>().Count() > 0)
                {
                    sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
                }

                // Append the new worksheet and associate it with the workbook.
                Sheet sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = sheetName };
                sheets.Append(sheet);
                //ready to write column header


                spreadsheetDocument.WorkbookPart.Workbook.Save();

                // Close the document handle.
                spreadsheetDocument.Close();
                stream.Close();
            }
        }

        public static void AddNewSpreadSheet(string filepath, string sheetName)
        {
            if (System.IO.File.Exists(filepath))
            {
                AppendSpreadSheet(filepath, sheetName);
                return;
            }

            // Create a spreadsheet document by supplying the filepath.
            // By default, AutoSave = true, Editable = true, and Type = xlsx.
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(filepath, SpreadsheetDocumentType.Workbook))
            {

                // Add a WorkbookPart to the document.
                WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();
                workbookpart.Workbook = new Workbook();

                // Add a WorksheetPart to the WorkbookPart.
                WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet(new SheetData());

                // Add Sheets to the Workbook.
                Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.
                    AppendChild<Sheets>(new Sheets());


                // Append a new worksheet and associate it with the workbook.
                Sheet sheet = new Sheet()
                {
                    Id = spreadsheetDocument.WorkbookPart.
                    GetIdOfPart(worksheetPart),
                    SheetId = 1,
                    Name = sheetName
                };
                sheets.Append(sheet);

                //ready to write column header

                workbookpart.Workbook.Save();

                // Close the document.
                spreadsheetDocument.Close();
            }

        }

        #endregion


        #region "Process WorkSheet"

        private DataTable AddBlankRowToTable(DataTable tbl) // this will useful incase of worksheet with no row 
        {
            DataRow row = tbl.NewRow();
            for (int i = 0; i < tbl.Columns.Count; i++)
            {
                row[i] = "";
            }
            tbl.Rows.Add(row);
            tbl.AcceptChanges();
            return tbl;
        }
        private void AddHeaderInWorkSheet(string filePath, DataTable tbl)//it will take last worksheet as default to write blank rows
        {

            tbl = AddBlankRowToTable(tbl);
            ProcessWorkSheet(filePath, tbl);

        }

        private void ProcessWorkSheet(string filePath, DataTable tbl)
        {

            // get worksheetpart 
            using (Stream stream = File.Open(filePath, FileMode.Open))
            {
                // Open a SpreadsheetDocument based on a stream.
                SpreadsheetDocument spreadsheet =
                        SpreadsheetDocument.Open(stream, true);

                WorkbookPart wbPart = spreadsheet.WorkbookPart;

                // Assuming last id is , is id for worksheet .
                Sheets sheets = spreadsheet.WorkbookPart.Workbook.GetFirstChild<Sheets>();
                uint sheetId = 0;
                if (sheets.Elements<Sheet>().Count() > 0)
                {
                    sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() - 1;
                }
                WorksheetPart worksheetPart = wbPart.WorksheetParts.ToList()[Convert.ToInt32(sheetId)];
                WriteDataTableToExcelWorksheet(tbl, worksheetPart, CultureInfo.InvariantCulture);
                spreadsheet.WorkbookPart.Workbook.Save();

                spreadsheet.Close();
            }
        }

        private void WriteDataTableToExcelWorksheet(DataTable dt, WorksheetPart worksheetPart, CultureInfo culture)
        {
            OpenXmlWriter writer = OpenXmlWriter.Create(worksheetPart, Encoding.ASCII);
            writer.WriteStartElement(new Worksheet());
            writer.WriteStartElement(new SheetData());

            string cellValue = "";

            //  Create a Header Row in our Excel file, containing one header for each Column of data in our DataTable.
            //
            //  We'll also create an array, showing which type each column of data is (Text or Numeric), so when we come to write the actual
            //  cells of data, we'll know if to write Text values or Numeric cell values.
            int numberOfColumns = dt.Columns.Count;
            bool[] IsNumericColumn = new bool[numberOfColumns];
            bool[] IsDateColumn = new bool[numberOfColumns];

            string[] excelColumnNames = new string[numberOfColumns];
            for (int n = 0; n < numberOfColumns; n++)
                excelColumnNames[n] = GetExcelColumnName(n);

            //
            //  Create the Header row in our Excel Worksheet
            //
            uint rowIndex = 1;

            writer.WriteStartElement(new Row { RowIndex = rowIndex });
            for (int colInx = 0; colInx < numberOfColumns; colInx++)
            {
                DataColumn col = dt.Columns[colInx];
                AppendTextCell(excelColumnNames[colInx] + "1", col.ColumnName, ref writer);
                IsNumericColumn[colInx] = (col.DataType.FullName == "System.Decimal") || (col.DataType.FullName == "System.Int32") || (col.DataType.FullName == "System.Double") || (col.DataType.FullName == "System.Single");
                IsDateColumn[colInx] = (col.DataType.FullName == "System.DateTime");
            }
            writer.WriteEndElement();   //  End of header "Row"

            //
            //  Now, step through each row of data in our DataTable...
            //
            double cellNumericValue = 0;
            foreach (DataRow dr in dt.Rows)
            {
                // ...create a new row, and append a set of this row's data to it.
                ++rowIndex;

                writer.WriteStartElement(new Row { RowIndex = rowIndex });

                for (int colInx = 0; colInx < numberOfColumns; colInx++)
                {
                    cellValue = dr.ItemArray[colInx].ToString();
                    cellValue = ReplaceHexadecimalSymbols(cellValue);

                    // Create cell with data
                    if (IsNumericColumn[colInx])
                    {
                        //  For numeric cells, make sure our input data IS a number, then write it out to the Excel file.
                        //  If this numeric value is NULL, then don't write anything to the Excel file.
                        cellNumericValue = 0;
                        if (double.TryParse(cellValue, out cellNumericValue))
                        {
                            if (culture != null)
                            {
                                cellValue = cellNumericValue.ToString(culture);
                            }
                            else
                            {
                                cellValue = cellNumericValue.ToString();
                            }
                            AppendNumericCell(excelColumnNames[colInx] + rowIndex.ToString(), cellValue, ref writer);
                        }
                    }
                    else if (IsDateColumn[colInx])
                    {
                        //  This is a date value.
                        DateTime dtValue;
                        string strValue = cellValue;
                        if (DateTime.TryParse(cellValue, out dtValue))
                        {
                            if (culture != null)
                            {
                                if (dtValue.Day > 12)
                                {
                                    strValue = dtValue.ToString(culture);
                                }
                                else
                                {
                                    strValue = dtValue.ToString();
                                }
                            }
                        }
                        AppendTextCell(excelColumnNames[colInx] + rowIndex.ToString(), strValue, ref writer);
                    }
                    else
                    {
                        //  For text cells, just write the input data straight out to the Excel file.
                        AppendTextCell(excelColumnNames[colInx] + rowIndex.ToString(), cellValue, ref writer);
                    }
                }
                writer.WriteEndElement(); //  End of Row
            }
            writer.WriteEndElement(); //  End of SheetData
            writer.WriteEndElement(); //  End of worksheet

            writer.Close();
        }

        //  Convert a zero-based column index into an Excel column reference  (A, B, C.. Y, Y, AA, AB, AC... AY, AZ, B1, B2..)
        private string GetExcelColumnName(int columnIndex)
        {
            //  eg  (0) should return "A"
            //      (1) should return "B"
            //      (25) should return "Z"
            //      (26) should return "AA"
            //      (27) should return "AB"
            //      ..etc..
            char firstChar;
            char secondChar;
            char thirdChar;

            if (columnIndex < 26)
            {
                return ((char)('A' + columnIndex)).ToString();
            }

            if (columnIndex < 702)
            {
                firstChar = (char)('A' + (columnIndex / 26) - 1);
                secondChar = (char)('A' + (columnIndex % 26));

                return string.Format("{0}{1}", firstChar, secondChar);
            }

            int firstInt = columnIndex / 26 / 26;
            int secondInt = (columnIndex - firstInt * 26 * 26) / 26;
            if (secondInt == 0)
            {
                secondInt = 26;
                firstInt = firstInt - 1;
            }
            int thirdInt = (columnIndex - firstInt * 26 * 26 - secondInt * 26);

            firstChar = (char)('A' + firstInt - 1);
            secondChar = (char)('A' + secondInt - 1);
            thirdChar = (char)('A' + thirdInt);

            return string.Format("{0}{1}{2}", firstChar, secondChar, thirdChar);
        }

        private string StripURL(string str)
        {
            //<a href='www.google.com'>unknown</a>
            if (str.StartsWith("<a href='"))
            {
                str = str.Substring(str.IndexOf(">") + 1, (str.Length - (str.IndexOf(">") + 1) - 4));
            }
            return str;
        }

        private void AppendTextCell(string cellReference, string cellStringValue, ref OpenXmlWriter writer)
        {
            //  Add a new Excel Cell to our Row 
            writer.WriteElement(new Cell
            {
                CellValue = new CellValue(StripURL(cellStringValue)),
                CellReference = cellReference,
                DataType = CellValues.String
            });
        }

        private string ReplaceHexadecimalSymbols(string txt)
        {
            string r = "[\x00-\x08\x0B\x0C\x0E-\x1F]";
            return Regex.Replace(txt, r, "", RegexOptions.Compiled);
        }

        private void AppendNumericCell(string cellReference, string cellStringValue, ref OpenXmlWriter writer)
        {
            //  Add a new Excel Cell to our Row 
            writer.WriteElement(new Cell
            {
                CellValue = new CellValue(cellStringValue),
                CellReference = cellReference,
                DataType = CellValues.Number
            });
        }

        #endregion


        #region "Read Data from worksheet"
        public DataTable ReadWorkSheetFromExcel(string filePath, string workSheetName)
        {
            DataTable dt = new DataTable("Table");

            using (SpreadsheetDocument document = SpreadsheetDocument.Open(filePath, false))
            {
                WorkbookPart wbPart = document.WorkbookPart;
                //SharedStringTablePart sstpart = wbPart.GetPartsOfType<SharedStringTablePart>().First();
                //SharedStringTable sst = sstpart.SharedStringTable;

                string relId = wbPart.Workbook.Descendants<Sheet>().First(s => workSheetName.Equals(s.Name, StringComparison.OrdinalIgnoreCase)).Id;
                Worksheet wsp2 = (wbPart.GetPartById(relId) as WorksheetPart).Worksheet;

                IEnumerable<Cell> cells = wsp2.Descendants<Cell>();
                IEnumerable<Row> rows = wsp2.Descendants<Row>();

                Console.WriteLine("Row count = {0}", rows.LongCount());
                Console.WriteLine("Cell count = {0}", cells.LongCount());


                foreach (Row row in rows)
                {
                    foreach (Cell c in row.Elements<Cell>())
                    {
                        //first make sure there are enough columns
                        //in order to to that have to get column index
                        var parsed = ExtractColumnAndRowNumbers(c.CellReference);
                        AddColumns(dt, parsed.Item1);

                        String str = String.Empty;
                        //make sure we have a column in our table
                        if ((c.DataType != null) && (c.DataType == CellValues.SharedString))
                        {
                            int ssid = int.Parse(c.CellValue.Text);
                            //str = sst.ChildElements[ssid].InnerText;
                            // Console.WriteLine("Shared string {0}: {1}", ssid, str);
                        }
                        else if (c.CellValue != null)
                        {
                            str = c.CellValue.Text;
                            Console.WriteLine("Cell contents: {0}", c.CellValue.Text);
                        }

                        String cellData = $"Cell data:{str}, Ref:{c.CellReference}, MetaIndex:{c.CellMetaIndex}";
                        //We don't know if there are empty rows but if so we have to add them to data table
                        AddRows(dt, parsed.Item2);

                        dt.Rows[parsed.Item2 - 1][parsed.Item1 - 1] = str;
                    }
                }

            }

            //assiging first row as column name and remove it from row
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                dt.Columns[i].ColumnName = dt.Rows[0][i].ToString();
            }
            dt.Rows.RemoveAt(0);

            return dt;
        }

        private static (int, int) ExtractColumnAndRowNumbers(string columnName)
        {
            if (string.IsNullOrEmpty(columnName))
                return (-1, -1);

            Match match = Regex.Match(columnName, @"[\d]+");
            if (!match.Success)
                return (-1, -1);

            int row = Int32.Parse(match.Value);
            //now strip of digit part
            string columnReference = Regex.Replace(columnName.ToUpper(), @"[\d]", string.Empty).ToUpperInvariant();

            int col = 0;

            for (int i = 0; i < columnReference.Length; i++)
            {
                col *= 26;
                col += (columnReference[i] - 'A' + 1);
            }

            return (col, row);
        }

        private void AddRows(DataTable dt, int row)
        {
            if (dt.Rows.Count >= row)
                return;
            for (int i = 0; i < row - dt.Rows.Count; i++)
                dt.Rows.Add();
        }

        private void AddColumns(DataTable dt, int colNum)
        {
            if (dt.Columns.Count >= colNum)
                return; //we have enough

            for (int i = 0; i < colNum - dt.Columns.Count; i++)
                dt.Columns.Add();
        }

        #endregion

    }


}
