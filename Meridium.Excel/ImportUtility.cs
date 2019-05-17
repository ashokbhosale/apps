using System;
using System.IO;
using System.Data;
using System.Xml;
using System.Collections;
using System.Runtime.InteropServices;
using System.Text;

namespace Meridium.Excel
{
	/// <summary>
	/// Summary description for ImportUtility.
	/// </summary>
    public class ImportUtility
    {
        private const string SELECT_SQL = "SELECT {0} FROM {1}";
        private const string TABLE_NAME = "TABLE_NAME";
        private const string COLUMN_NAME = "COLUMN_NAME";
        private const string Y = "Yes";
        private const string N = "No";
        private const string DEFAULT_COLUMN_NAME = "F";

        public ImportUtility() { }

        private static string GetFileName(string filePath)
        {
            if (string.IsNullOrEmpty(ValidateFileName(filePath)))
                return string.Empty;
            FileInfo info = new FileInfo(filePath);
            return info.Name;
        }

        private static string GetDirectoryPath(string filePath)
        {
            if (string.IsNullOrEmpty(ValidateFileName(filePath)))
                return string.Empty;
            FileInfo info = new FileInfo(filePath);
            return info.Directory.FullName;
        }

        private static string ValidateFileName(string filePath)
        {
            String dir = Path.GetDirectoryName(filePath);
            String fileName = Path.GetFileName(filePath);

            char[] invalidPath = Path.GetInvalidPathChars();
            foreach (char c in invalidPath)
            {
                if (dir.IndexOf(c) == -1)
                    continue;
                throw new Exception("The path contains an invalid character.");
                //Meridium.Common.MeridiumMsgBox.MeridiumMsgBox.Show(StringResources.INVALID_CHAR, StringResources.INVALID_FILEPATH_NAME);
                //return string.Empty;
            }
            char[] invalidFile = Path.GetInvalidFileNameChars();
            foreach (char c in invalidFile)
            {
                if (fileName.IndexOf(c) == -1)
                    continue;
                //Meridium.Common.MeridiumMsgBox.MeridiumMsgBox.Show(StringResources.INVALID_CHAR, StringResources.INVALID_FILEPATH_NAME);
                //return string.Empty;
                  throw new Exception("The file name contains an invalid character.");
            }
            return filePath;
        }

        public static string[] GetExcelWorksheets(ImportType importType, string filePath, bool firstRowContainsHeaders)
        {
            throw new NotImplementedException("(US328120) - OleDb is not supported in DotNet Core");
            //            try
            //            {
            //                ExcelProperties properties = new ExcelProperties(firstRowContainsHeaders, true, importType == ImportType.CSV);
            //                DataTable table = ExcelUtility.GetOleDbSchemaTable(filePath, System.Data.OleDb.OleDbSchemaGuid.Tables, new[] { null, null, null, "TABLE" }, properties);
            //                ArrayList worksheets = new ArrayList();

            //                for (int i = 0; i < table.Rows.Count; i++)
            //                {
            //                    string tableName = table.Rows[i][2].ToString().Trim();

            //                    if (tableName.IndexOf(" ") > 0)
            //                    {
            //                        if (tableName[0].ToString() == "'")
            //                            tableName = tableName.Remove(0, 1);

            //                        if (tableName[tableName.Length - 1].ToString() == "'")
            //                            tableName = tableName.Remove(tableName.Length - 1, 1);
            //                    }

            //                    // worksheets must end in $
            //                    if (tableName[tableName.Length - 1] == '$')
            //                        worksheets.Add(tableName);
            //                }

            //                worksheets.Sort();
            //                string[] retVal = new string[worksheets.Count];
            //                worksheets.CopyTo(retVal);
            //                return retVal;
            //            }
            //            catch (Exception ex) { throw ex; }
            //            finally
            //            {
            ////                conn.Close();
            ////                conn.Dispose();
            //            }
        }

        public static string[] GetExcelColumnNames(ImportType importType, string filePath, bool firstRowContainsHeaders, string worksheetName)
        {
            throw new NotImplementedException("(US328120) - OleDb is not supported in DotNet Core");
//            try
//            {
//                ExcelProperties properties = new ExcelProperties(firstRowContainsHeaders, true, importType == ImportType.CSV);
//                DataTable table = ExcelUtility.GetOleDbSchemaTable(filePath, System.Data.OleDb.OleDbSchemaGuid.Columns, new[] { null, null, worksheetName, null }, properties);
//                ArrayList cols = new ArrayList();

//                try
//                {
//                    for (int i = 0; i < table.Rows.Count; i++)
//                    {
//                        //if(table.Rows[i][TABLE_NAME].ToString() == worksheetName)
//                        cols.Add(table.Rows[i][COLUMN_NAME].ToString());
//                    }
//                }
//                finally { table.Dispose(); }

//                string[] retVal = new string[cols.Count];
//                cols.CopyTo(retVal);
//                return retVal;
//            }
//            catch (Exception ex) { throw ex; }
//            finally
//            {
////                conn.Close();
////                conn.Dispose();
//            }
        }

        public static System.Data.DataSet ImportExcelData(ImportType importType, string fileName, bool firstRowContainsHeaders, string worksheetName)
        {
            string sql = String.Format(SELECT_SQL, "*", "[" + worksheetName + "]");
            ExcelProperties properties = new ExcelProperties(firstRowContainsHeaders, true, importType == ImportType.CSV);
            return ExcelUtility.Fill(fileName, sql, null, properties);
        }

        public static System.Data.DataSet ImportExcelData(ImportType importType, string fileName, bool firstRowContainsHeaders, string worksheetName, string[] columns)
        {
            string selectColumns = string.Empty;

            for (int i = 0; i < columns.Length; i++)
            {
                if (i == 0)
                    selectColumns = columns[i];
                else
                    selectColumns += "," + columns[i];
            }

            if (selectColumns.Trim() == string.Empty)
                selectColumns = "*";

            string sql = String.Format(SELECT_SQL, selectColumns, "[" + worksheetName + "]");

            ExcelProperties properties = new ExcelProperties(firstRowContainsHeaders, true, importType == ImportType.CSV);
            return ExcelUtility.Fill(fileName, sql, null, properties);
        }

        public static System.Data.DataSet ImportExcelData(ImportType importType, string fileName, bool firstRowContainsHeaders, string worksheetName, string startRange, string endRange)
        {
            string sql = String.Format(SELECT_SQL, "*", "[" + worksheetName + startRange + ":" + endRange + "]");
            ExcelProperties properties = new ExcelProperties(firstRowContainsHeaders, true, importType == ImportType.CSV);
            return ExcelUtility.Fill(fileName, sql, null, properties);
        }

        public static System.Data.DataSet ImportCSVData(string fileName, bool firstRowContainsHeaders)
        {
            string sql = String.Format(SELECT_SQL, "*", "[" + GetFileName(fileName) + "]");
            ExcelProperties properties = new ExcelProperties(firstRowContainsHeaders, true, true);
            return ExcelUtility.Fill(fileName, sql, null, properties);
        }
        [DllImport("shlwapi.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern bool PathCanonicalize([Out] StringBuilder dst, string src);
        public static System.Data.DataSet ImportTextData(string fileName, bool firstRowContainsHeaders, char colSeparator)
        {
            string contents = string.Empty;
            StreamReader reader = null;
            try
            {
                StringBuilder sb = new StringBuilder(Math.Max(260, 2 * fileName.Length));
                PathCanonicalize(sb, fileName);
                reader = File.OpenText(sb.ToString());
            }
            catch
            {
                //reader = File.OpenText(ValidateFileName(fileName));
            }

            try { contents = reader.ReadToEnd(); }
            finally
            {
                reader.Close();
                reader.Dispose();
            }

            int currentRow = 0;
            char[] rowSep = new char[] { (char)13, (char)10 };
            char[] colSep = new char[] { colSeparator };
            System.Data.DataSet dataSet = new System.Data.DataSet();
            DataTable dataTable = new DataTable();

            string[] rows = null;
            string[] cols = null;

            System.Text.RegularExpressions.Regex regex = new System.Text.RegularExpressions.Regex("\r\n", System.Text.RegularExpressions.RegexOptions.Multiline);

            rows = regex.Split(contents);

            ArrayList processedRows = new ArrayList();

            for (int i = 0; i < rows.Length; i++)
                if (rows[i].Trim() != string.Empty)
                    processedRows.Add(rows[i]);

            rows = new String[processedRows.Count];
            processedRows.CopyTo(rows);
            cols = rows[0].Split(colSep);

            for (int i = 0; i < cols.Length; i++)
            {
                int dupColCount = 1;

                if (firstRowContainsHeaders)
                {
                    while (dataTable.Columns.Contains(cols[i]))
                    {
                        cols[i] = cols[i] + "(" + dupColCount.ToString() + ")";
                        dupColCount++;
                    }
                    dataTable.Columns.Add(cols[i], typeof(string));
                }
                else
                {
                    dupColCount = 0;

                    while (dataTable.Columns.Contains(DEFAULT_COLUMN_NAME + (i + 1 + dupColCount)))
                        dupColCount++;

                    dataTable.Columns.Add(DEFAULT_COLUMN_NAME + (i + 1 + dupColCount), typeof(string));
                }
            }

            if (firstRowContainsHeaders)
                currentRow++;

            for (; currentRow < rows.Length; currentRow++)
            {
                cols = rows[currentRow].Split(colSep);
                object[] rowValues = new object[dataTable.Columns.Count];

                for (int i = 0; i < dataTable.Columns.Count; i++)
                {
                    if (i < cols.Length)
                        rowValues[i] = cols[i];
                    else
                        rowValues[i] = System.DBNull.Value;
                }
                dataTable.Rows.Add(rowValues);
            }

            dataSet.Tables.Add(dataTable);
            return dataSet;
        }

        public static System.Data.DataSet ImportXMLData(string fileName)
        {
            System.Data.DataSet dataSet = new System.Data.DataSet();
            try
            {
                StringBuilder sb = new StringBuilder(Math.Max(260, 2 * fileName.Length));
                PathCanonicalize(sb, fileName);
                dataSet.ReadXml(Path.GetFullPath(sb.ToString()));
            }
            catch
            {
                //dataSet.ReadXml(ValidateFileName(fileName));
            }
            return dataSet;
        }
    }

	public enum ImportType
	{
		Excel2002,
		CSV,
		XML
	}
}
