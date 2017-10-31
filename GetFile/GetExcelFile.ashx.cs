using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Reflection;
using System.IO;
using System.Text;

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
            context.Response.ContentType = "application/vnd.ms-excel";
            context.Response.Headers.Add("Content-Disposition", $"attachment; filename = test-{DateTime.Now.ToString("yyyyMMddHHmmss")}.xlsx;");

            string tempFileName = saveTemplateAsTemp();
            byte[] content = File.ReadAllBytes(tempFileName);

            context.Response.BinaryWrite(content);
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

        private string getConnectionString(string fpath)
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