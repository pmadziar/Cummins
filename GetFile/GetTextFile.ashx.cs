using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;

namespace GetFile
{
    /// <summary>
    /// Summary description for GetTextFile
    /// </summary>
    public class GetTextFile : IHttpHandler
    {

        public void ProcessRequest(HttpContext context)
        {
            string param = "NONE";
            try
            {
                string p = context.Request.QueryString["param"];
                if (!string.IsNullOrEmpty(p)) param = p;
            }
            catch { }
            context.Response.ContentType = "text/plain";
            context.Response.Headers.Add("Content-Disposition", "attachment; filename = text.txt;");
            context.Response.Write(GetText(param));
        }

        public bool IsReusable
        {
            get
            {
                return false;
            }
        }

        private string GetText(string param)
        {
            StringWriter sw = new StringWriter();
            sw.WriteLine($"Param is: {param}");
            sw.WriteLine(@"Lorem ipsum dolor sit amet, consectetur adipiscing elit. Nunc sollicitudin, quam ac dictum rhoncus, erat sem vehicula dui, a placerat odio tortor venenatis lacus. Pellentesque eget nibh finibus, bibendum mauris non, posuere est. Ut commodo leo malesuada, aliquam augue non, varius purus. Cras venenatis odio sed vulputate suscipit. Vivamus efficitur, metus quis consequat tempor, enim nulla molestie massa, non viverra lorem neque vitae quam. Sed et aliquet risus. Maecenas iaculis et augue sed aliquet. Nunc tincidunt dolor a leo tempus, ac viverra odio consectetur. Ut sem quam, elementum ut lorem vel, mattis aliquam massa. Quisque in ipsum ac tortor malesuada auctor.
Aenean quis consequat nibh, eu tincidunt elit. Mauris commodo ultricies elit, ac tincidunt neque congue non. Etiam imperdiet interdum condimentum. Duis euismod lacus vel magna lobortis, ac feugiat dolor pretium. Cras commodo mauris ac accumsan mollis. Nam ornare a odio sit amet vulputate. Mauris sed ligula a felis congue porttitor. Nulla pulvinar purus quis lectus dictum, at cursus libero porttitor. Nullam fringilla elit ut egestas dapibus. Integer lacinia, erat vitae dictum molestie, neque nibh efficitur elit, sit amet cursus sem nulla pharetra orci. Proin efficitur tortor quis magna vulputate, quis finibus nisl gravida. Curabitur pellentesque tortor vel eros bibendum, at maximus tellus consequat. Ut vehicula lorem orci, ut dignissim justo pharetra vel. Etiam luctus a odio bibendum dignissim.
Proin semper, libero et malesuada mattis, eros quam malesuada sapien, eu elementum orci risus sit amet felis. Donec non venenatis elit, eu viverra tortor. Duis ut elementum mi. Donec et eros ut risus faucibus porttitor vitae sed ante. Donec malesuada dictum dapibus. Quisque tortor odio, vestibulum quis mi at, elementum tincidunt erat. Aliquam non elit at magna venenatis venenatis eget nec massa. Fusce sagittis, lectus et pretium gravida, odio mi lobortis tortor, a consequat nunc leo vel sem. Fusce efficitur pellentesque pulvinar.");
            return sw.ToString();
        }
    }
}
