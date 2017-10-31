using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestXslWrite
{
    class Program
    {
        static void Main(string[] args)
        {
            const string templatePath = @"C:\Code\Cummins\GetFile\Template\AbsenceRights.xlsx";
            const string destDir = @"C:\Users\Administrator\Desktop\Excel";
            string xlsPath = $"{destDir}\\test-{DateTime.Now.ToString("yyyyMMddHHmmss")}.xlsx";

            File.WriteAllBytes(xlsPath, File.ReadAllBytes(templatePath));
            writeExcelFile(xlsPath);
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
VALUES('John','Doe',DATEVALUE('2014-01-01'),DATEVALUE('2015-01-01'),'InszX',DATEVALUE('1990-01-02'),'TestClass',VAL('1234'),54321, 10, DATEVALUE('2017-11-01'),DATEVALUE('2017-11-11'));";
                cmd.ExecuteNonQuery();

                cmd.CommandText = @"INSERT INTO [Input Absence rights$]
                ([First Name], [Last Name], [Employee In Date], [Employee Out Date], [Insz], [Date Of Birth], [Absence Class], [Absence Code], [Nilos Code], [Saldo], [From Date], [To Date]) 
VALUES('Jane','Smith',DATEVALUE('2014-01-01'),DATEVALUE('2015-01-01'),'InszX',DATEVALUE('1990-01-02'),'TestClass',VAL('1234'),54321, 10, DATEVALUE('2017-11-01'),DATEVALUE('2017-11-11'));";
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
