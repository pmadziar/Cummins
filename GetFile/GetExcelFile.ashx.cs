using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Reflection;
using System.IO;
using System.Text;
using System.Data.OleDb;

namespace GetFile
{
    /// <summary>
    /// Summary description for GetExcelFile
    /// </summary>
    public class GetExcelFile : IHttpHandler
    {
        private const string resourceName = "GetFile.Template.AbsenceRights.xlsx";

        public void ProcessRequest(HttpContext context)
        {
            string tempFileName = null;
            try
            {
                context.Response.ContentType = "application/vnd.ms-excel";
                context.Response.Headers.Add("Content-Disposition", $"attachment; filename = test-{DateTime.Now.ToString("yyyyMMddHHmmss")}.xlsx;");

                tempFileName = saveTemplateAsTemp();
                writeExcelFile(tempFileName);
                byte[] content = File.ReadAllBytes(tempFileName);

                context.Response.BinaryWrite(content);
            }
            catch (Exception)
            {

                throw;
            }
            finally
            {
                if(!string.IsNullOrEmpty(tempFileName) && File.Exists(tempFileName))
                {
                    File.Delete(tempFileName);
                }
            }

        }

        public bool IsReusable
        {
            get
            {
                return false;
            }
        }

        private string saveTemplateAsTemp()
        {
            Assembly aa = this.GetType().Assembly;
            string tempDir = Path.GetTempPath();
            string tempFName = $"AbsenceRights-{Guid.NewGuid().ToString("D")}.xlsx";
            string ret = Path.Combine(tempDir, tempFName);
            byte[] bytes;

            using (Stream stream = aa.GetManifestResourceStream(resourceName))
            using (BinaryReader reader = new BinaryReader(stream))
            {
                bytes = reader.ReadBytes((int)stream.Length);
            }

            File.WriteAllBytes(ret, bytes);

            return ret;
        }

        private static void writeExcelFile(string fpath)
        {
            string connectionString = getConnectionString(fpath);

            using (OleDbConnection conn = new OleDbConnection(connectionString))
            {
                conn.Open();
                OleDbCommand cmd = new OleDbCommand();
                cmd.Connection = conn;

                cmd.CommandText = @"INSERT INTO [Input Absence rights$]
                ([First Name], [Last Name], [Employee In Date], [Employee Out Date], [Insz], [Date Of Birth], [Absence Class], [Absence Code], [Nilos Code], [Saldo], [From Date], [To Date]) 
VALUES('John','Doe',DATEVALUE('2014-01-01'),DATEVALUE('2015-01-01'),'InszX',DATEVALUE('1990-01-02'),'TestClass',1234,54321, 10, DATEVALUE('2017-11-01'),DATEVALUE('2017-11-11'));";
                cmd.ExecuteNonQuery();

                cmd.CommandText = @"INSERT INTO [Input Absence rights$]
                ([First Name], [Last Name], [Employee In Date], [Employee Out Date], [Insz], [Date Of Birth], [Absence Class], [Absence Code], [Nilos Code], [Saldo], [From Date], [To Date]) 
VALUES('Jane','Smith',DATEVALUE('2014-01-01'),DATEVALUE('2015-01-01'),'InszX',DATEVALUE('1990-01-02'),'TestClass',1234,54321, 10, DATEVALUE('2017-08-01'),DATEVALUE('2017-11-11'));";
                cmd.ExecuteNonQuery();

                conn.Close();
            }
        }

        private static string getConnectionString(string fpath)
        {
            Dictionary<string, string> props = new Dictionary<string, string>();

            // XLSX - Excel 2007, 2010, 2012, 2013
            props["Provider"] = "Microsoft.ACE.OLEDB.12.0;";
            props["Extended Properties"] = "Excel 12.0 XML";
            props["Data Source"] = fpath;


            StringBuilder sb = new StringBuilder();

            foreach (KeyValuePair<string, string> prop in props)
            {
                sb.Append(prop.Key);
                sb.Append('=');
                sb.Append(prop.Value);
                sb.Append(';');
            }

            return sb.ToString();
        }

    }
}