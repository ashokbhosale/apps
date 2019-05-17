using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Meridium.Excel;

namespace Meridium.Excel
{
    internal class ExcelHelper
    {

        [DllImport("shlwapi.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern bool PathCanonicalize([Out] StringBuilder dst, string src);
        private static string getCannolizePath(string filepath)
        {
            try
            {
                String filename = filepath;
                StringBuilder sb = new StringBuilder(Math.Max(260, 2 * filename.Length));
                PathCanonicalize(sb, filename);
                String Cannoicalize_Path = sb.ToString();
                return Cannoicalize_Path;
            }
            catch
            {
                throw new Exception("File path cannot be cannolize");
            }
        }              

        #region IExcelHelper Members
        public MemoryStream CreateExcelDocument(DataTable dt, CultureInfo culture)
        {
            var ds = dt.DataSet;
            if (ds == null)
            {
                ds = new DataSet();
                ds.Tables.Add(dt);
            }
            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
            {
                for (int j = 0; j < ds.Tables[0].Columns.Count; j++)
                {
                    var cellValue = Convert.ToString(ds.Tables[0].Rows[i][j]);
                    if (string.IsNullOrWhiteSpace(cellValue))
                        ds.Tables[0].Rows[i][j] = DBNull.Value;
                    else
                        ds.Tables[0].Rows[i][j] = cellValue;
                }
            }


            return ExcelDocumnent.CreateExcelDocument(ds, culture);
        }

        public bool Export(DataTable dt, string filepath, string tablename, ExcelProperties properties, CultureInfo culture)
        {
            filepath = getCannolizePath(filepath);

            if (culture == null)
                culture = new System.Globalization.CultureInfo("en-EN");


            dt.TableName = tablename;
            MemoryStream ms = CreateExcelDocument(dt, culture);
            FileStream file = new FileStream(filepath, FileMode.Create, FileAccess.Write);
            ms.WriteTo(file);
            file.Close();
            ms.Close();

            return true;
        }

        public int ExecuteNonQuery(string filepath, string query, string[] parameters, ExcelProperties properties)
        {
            filepath = getCannolizePath(filepath);
            //logger.Trace("ExecuteNonQuery(string filepath, string query, string[] parameters, ExcelProperties properties)");

            if (string.IsNullOrEmpty(filepath))
                throw new Exception(ExcelResources.DataSourceEmpty);
            if (string.IsNullOrEmpty(query))
                throw new Exception(ExcelResources.QueryEmpty);

            OpenXmlClient con = new OpenXmlClient();

            return con.ExecuteNonQuery(query, parameters, filepath); ;
        }

        public int ExecuteMultipleQuery(string filepath, string[] query, ExcelProperties properties)
        {
            filepath = getCannolizePath(filepath);

            if (string.IsNullOrEmpty
                (filepath))
                throw new Exception(ExcelResources.DataSourceEmpty);
            if (query.Length == 0)
                throw new Exception(ExcelResources.QueryEmpty);


            OpenXmlClient con = new OpenXmlClient();

            con.ExecuteMultipleQuery(filepath, query);

            return 1;
        }

        public DataSet Fill(string filepath, string query, string[] parameters, ExcelProperties properties)
        {
            filepath = getCannolizePath(filepath);
            //logger.Trace("Fill(string filepath, string query, string[] parameters, ExcelProperties properties)");

            if (string.IsNullOrEmpty(filepath))
                throw new Exception(ExcelResources.DataSourceEmpty);
            if (string.IsNullOrEmpty(query))
                throw new Exception(ExcelResources.QueryEmpty);

            OpenXmlClient con = new OpenXmlClient();
            DataSet ds = con.Fill(query, parameters, filepath);
            return ds;
        }


        public DataSet FillWithName(string filepath, string query, string[] parameters, string srcTable, ExcelProperties properties)
        {
            filepath = getCannolizePath(filepath);

            DataSet ds = Fill(filepath, query, parameters, properties);
            ds.Tables[0].TableName = srcTable;
            ds.AcceptChanges();
            return ds;
        }


        public DataTable GetOleDbSchemaTable(string filepath, Guid schema, string[] parameters, ExcelProperties properties)
        {
            filepath = getCannolizePath(filepath);

            if (string.IsNullOrEmpty(filepath))
                throw new Exception(ExcelResources.DataSourceEmpty);
            if (schema == null || string.IsNullOrEmpty(schema.ToString()))
                throw new Exception(ExcelResources.SchemaEmpty);


            OpenXmlClient openxmlcon = new OpenXmlClient();
            DataTable openxmltbl = openxmlcon.GetSchemaTable(filepath);
            return openxmltbl;
        }

        public void Dispose()
        {
        }

        #endregion

    }
}
