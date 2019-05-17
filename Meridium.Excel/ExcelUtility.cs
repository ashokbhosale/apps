using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Meridium.Excel
{
    public class ExcelUtility
    {
        public static bool Export(DataTable dt, string filepath, string tablename, ExcelProperties properties)
        {
            if (dt == null)
                throw new Exception(ExcelResources.DataTableNull);
            if (string.IsNullOrEmpty(filepath))
                throw new Exception(ExcelResources.DataSourceEmpty);

            ExcelHelper exlHelp = new ExcelHelper();
            var culture = new CultureInfo("en-US");
            bool rtn = exlHelp.Export(dt, filepath, tablename, properties, culture);

            return rtn;
        }

        public static DataSet Fill(string filepath, string query, object[] parameters)
        {
            if (string.IsNullOrEmpty(filepath))
                throw new Exception(ExcelResources.DataSourceEmpty);
            if (string.IsNullOrEmpty(query))
                throw new Exception(ExcelResources.QueryEmpty);

            ExcelHelper exlHelp = new ExcelHelper();
            string[] param = CastParameters(parameters);
            DataSet rtn = exlHelp.Fill(filepath, query, param, null);

            return rtn;
        }

        public static DataSet Fill(string filepath, string query, object[] parameters, ExcelProperties properties)
        {
            if (string.IsNullOrEmpty(filepath))
                throw new Exception(ExcelResources.DataSourceEmpty);
            if (string.IsNullOrEmpty(query))
                throw new Exception(ExcelResources.QueryEmpty);

            ExcelHelper exlHelp = new ExcelHelper();
            string[] param = CastParameters(parameters);
            DataSet rtn = exlHelp.Fill(filepath, query, param, properties);

            return rtn;
        }

        public static DataSet Fill(string filepath, string query, object[] parameters, string srcTable)
        {
            if (string.IsNullOrEmpty(filepath))
                throw new Exception(ExcelResources.DataSourceEmpty);
            if (string.IsNullOrEmpty(query))
                throw new Exception(ExcelResources.QueryEmpty);
            if (string.IsNullOrEmpty(srcTable))
                throw new Exception(ExcelResources.SourceTableEmpty);

            ExcelHelper exlHelp = new ExcelHelper();
            string[] param = CastParameters(parameters);
            DataSet rtn = exlHelp.FillWithName(filepath, query, param, srcTable, null);

            return rtn;
        }

        public static DataSet Fill(string filepath, string query, object[] parameters, string srcTable, ExcelProperties properties)
        {
            if (string.IsNullOrEmpty(filepath))
                throw new Exception(ExcelResources.DataSourceEmpty);
            if (string.IsNullOrEmpty(query))
                throw new Exception(ExcelResources.QueryEmpty);
            if (string.IsNullOrEmpty(srcTable))
                throw new Exception(ExcelResources.SourceTableEmpty);

            ExcelHelper exlHelp = new ExcelHelper();
            string[] param = CastParameters(parameters);
            DataSet rtn = exlHelp.FillWithName(filepath, query, param, srcTable, properties);

            return rtn;
        }

        public static int ExecuteNonQuery(string filepath, string query, object[] parameters)
        {
            if (string.IsNullOrEmpty(filepath))
                throw new Exception(ExcelResources.DataSourceEmpty);
            if (string.IsNullOrEmpty(query))
                throw new Exception(ExcelResources.QueryEmpty);

            ExcelHelper exlHelp = new ExcelHelper();
            string[] param = CastParameters(parameters);
            int rtn = exlHelp.ExecuteNonQuery(filepath, query, param, null);

            return rtn;
        }

        public static int ExecuteNonQuery(string filepath, string query, object[] parameters, ExcelProperties properties)
        {
            if (string.IsNullOrEmpty(filepath))
                throw new Exception(ExcelResources.DataSourceEmpty);
            if (string.IsNullOrEmpty(query))
                throw new Exception(ExcelResources.QueryEmpty);

            ExcelHelper exlHelp = new ExcelHelper();
            string[] param = CastParameters(parameters);
            int rtn = exlHelp.ExecuteNonQuery(filepath, query, param, properties);

            return rtn;
        }

        public static int ExecuteMultipleQuery(string filepath, string[] query)
        {
            if (string.IsNullOrEmpty(filepath))
                throw new Exception(ExcelResources.DataSourceEmpty);
            if (query == null)
                throw new Exception(ExcelResources.QueryArrayNull);
            foreach (string x in query)
                if (string.IsNullOrEmpty(x))
                    throw new Exception(ExcelResources.QueryArrayInstanceEmpty);

            ExcelHelper exlHelp = new ExcelHelper();
            int rtn = exlHelp.ExecuteMultipleQuery(filepath, query, null);

            return rtn;
        }

        public static int ExecuteMultipleQuery(string filepath, string[] query, ExcelProperties properties)
        {
            if (string.IsNullOrEmpty(filepath))
                throw new Exception(ExcelResources.DataSourceEmpty);
            if (query == null)
                throw new Exception(ExcelResources.QueryArrayNull);
            foreach (string x in query)
                if (string.IsNullOrEmpty(x))
                    throw new Exception(ExcelResources.QueryArrayInstanceEmpty);

            ExcelHelper exlHelp = new ExcelHelper();
            int rtn = exlHelp.ExecuteMultipleQuery(filepath, query, properties);

            return rtn;
        }

        public static DataTable GetOleDbSchemaTable(string filepath, Guid schema, string[] parameters)
        {
            if (string.IsNullOrEmpty(filepath))
                throw new Exception(ExcelResources.DataSourceEmpty);
            if (schema == null || string.IsNullOrEmpty(schema.ToString()))
                throw new Exception(ExcelResources.SchemaEmpty);

            ExcelHelper exlHelp = new ExcelHelper();
            DataTable rtn = exlHelp.GetOleDbSchemaTable(filepath, schema, parameters, null);

            return rtn;
        }

        public static DataTable GetOleDbSchemaTable(string filepath, Guid schema, string[] parameters, ExcelProperties properties)
        {
            if (string.IsNullOrEmpty(filepath))
                throw new Exception(ExcelResources.DataSourceEmpty);
            if (schema == null || string.IsNullOrEmpty(schema.ToString()))
                throw new Exception(ExcelResources.SchemaEmpty);

            ExcelHelper exlHelp = new ExcelHelper();
            DataTable rtn = exlHelp.GetOleDbSchemaTable(filepath, schema, parameters, properties);

            return rtn;
        }


        private static string[] CastParameters(object[] parameters)
        {
            string[] param = null;

            if (parameters == null)
                return param;

            param = new string[parameters.Length];
            for (int i = 0; i < parameters.Length; i++)
                param[i] = parameters[i] == DBNull.Value ? null : parameters[i].ToString();

            return param;
        }
    }
}
