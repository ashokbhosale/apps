using System;
using System.Collections.Generic;
using System.Data;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using System.IO;
using System.Runtime.InteropServices;

namespace Meridium.Excel
{
	/// <summary>
	/// Summary description for ExportUtility.
	/// </summary>
	public class ExportUtility
	{
		private const string EXP_INVALID_CHARS		= "[^a-zA-Z0-9_]";
		private const string REPLACEMENT_CHAR		= "_";
		private const string USE_HEADER				= "Yes";
		private const string TEXT_DATA_TYPE			= "LongText";
		private const string NUMERIC_DATA_TYPE		= "Numeric";
		private const string DOUBLE_DATA_TYPE		= "Double";
		private const string DATE_DATA_TYPE			= "Date";
		private const string PARAM_IDENTIFIER		= "?";
		private const string COMMA					= ",";

		public ExportUtility()
		{
		}

        private static string PrepareColumnName(string col)
        {
            //Excel Driver Characteristics
            //Maximum column name length
            //Column names over 64 characters will produce an error. 
            //string tempColumnName = System.Text.RegularExpressions.Regex.Replace(col, EXP_INVALID_CHARS, REPLACEMENT_CHAR);

            string tempColumnName = string.Empty;

            if (!string.IsNullOrEmpty(col))
            {
                char[] array = col.ToCharArray();

                foreach (char c in array)
                {
                    // passing through if it is between 0-9, a-z, A-Z 
                    // or not an ASCII character, or if it is a space, an "'", a "?" or a "-";
                    // otherwise replacing with "_".
                    if ((c >= '0' && c <= '9') || (c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z')
                        || (c > 127) || (c == ' ') || (c == '\'') || (c == '?') || (c == '-') || (c == '(') || (c == ')'))
                    {
                        tempColumnName += c;
                    }
                    else
                        tempColumnName += "_";
                }
            }

            if (tempColumnName.Length > 64)
                tempColumnName = tempColumnName.Substring(0, 63);

            return "[" + tempColumnName.Trim() + "]";
        }

        public static void ExportToExcel(string destinationFileName, System.Data.DataSet dataSet)
        {
            ExcelUtility.Export(dataSet.Tables[0], destinationFileName, "Sheet1", null);

		}
        public static void ExportToExcel(string destinationFileName, System.Data.DataSet dataSet, string worksheetName)
        {
            ExcelUtility.Export(dataSet.Tables[0], destinationFileName, worksheetName, null);
            
        }

        public static void ExportToCsv(string destinationFileName, System.Data.DataSet dataSet)
		{
			ExportToCsv(destinationFileName, dataSet, ',');
		}

        public static void ExportToCsv(string destinationFileName, System.Data.DataSet dataSet, char delimiter)
		{
			ExportUtility.CreateFile(destinationFileName, ExportUtility.CreateCSVString(dataSet, delimiter));
		}

        public static void ExportToXml(string destinationFileName, System.Data.DataSet dataSet)
		{
			ExportUtility.CreateFile(destinationFileName, ExportUtility.CreateXMLString(dataSet));			
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
                //return string.Empty;
            }
            char[] invalidFile = Path.GetInvalidFileNameChars();
            foreach (char c in invalidFile)
            {
                if (fileName.IndexOf(c) == -1)
                    continue;
                throw new Exception("The file name contains an invalid character.");
                //return string.Empty;
            }
            return filePath;
        }
        [DllImport("shlwapi.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern bool PathCanonicalize([Out] StringBuilder dst, string src);
		private static void CreateFile(string fileName, string contents)
		{          
            System.IO.StreamWriter writer = null;
            try
            {
                StringBuilder sb = new StringBuilder(Math.Max(260, 2 * fileName.Length));
                PathCanonicalize(sb, fileName);
                string canonicalPath = sb.ToString();
                writer = File.CreateText(canonicalPath);
            }
            catch
            {
               // writer = File.CreateText(ValidateFileName(fileName));
            }
			try
			{
				writer.Write(contents);
			}			
			finally
			{
				writer.Close();	
			}			
		}

		/*		CreateXMLString
		 * Returns a string containing XML serialized data
		 * from the given dataset.
		 * 
		 * Modified 17 Apr 2006 nbaker
		 * Calls ProcessDatasetXml to strip timezone information
		 * from the serialized data.
		 */
        public static string CreateXMLString(System.Data.DataSet dataSet)
		{
			if(dataSet == null)
				return string.Empty;

			using (StringWriter writer = new StringWriter())
			{
				dataSet.WriteXml(writer,XmlWriteMode.WriteSchema);
                return Regex.Replace(writer.ToString(),
                @"(?<date>[\d-]*T[\d:\.]*)[+-]\d+:\d\d",
                "${date}", RegexOptions.Compiled);
			}
		}

        public static string CreateCSVString(System.Data.DataSet dataSet, char delimiter)
		{
			// add header values
			System.Text.StringBuilder rows = new System.Text.StringBuilder();
			for(int i = 0; i < dataSet.Tables[0].Columns.Count; i++)
			{
				if(i > 0)
					rows.Append(delimiter);

				string colName = dataSet.Tables[0].Columns[i].ColumnName.Replace(delimiter, '_');
				rows.Append(colName);
			}
			rows.Append("\r\n");


			//add row values
			for(int i = 0; i < dataSet.Tables[0].Rows.Count; i++)
			{
				for(int j = 0; j < dataSet.Tables[0].Columns.Count; j++)
				{
					if(j > 0)
						rows.Append(delimiter);

					string colVal = dataSet.Tables[0].Rows[i][j] == null ? string.Empty : dataSet.Tables[0].Rows[i][j].ToString().Replace(delimiter, ' ');
					rows.Append(colVal);
				}
				rows.Append("\r\n");
			}

			return rows.ToString();
		}

	}
}
