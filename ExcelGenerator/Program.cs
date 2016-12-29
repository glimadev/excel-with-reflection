using ExcelGenerator.VM.Model;
using ExcelGenerator.VM.Out;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelGenerator
{
    class Program
    {
        static void Main(string[] args)
        {
            ExportUserVM exportUserVM = new ExportUserVM();

            exportUserVM.Users.Add(new UserModelVM
            {
                FullName = "John",
                City = "New York",
                State = "NY",
                UserId = Guid.NewGuid().ToString(),
                CreatedDate = DateTime.Now
            });
            
            byte[] content = ExcelHelper.WriteTsv<ExportUserVM>(exportUserVM);

            string fileName = string.Format("{0:yyyy-MM-dd_HH_mm-ss}.xls", DateTime.Now);

            string directory = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);

            File.WriteAllBytes(directory + "/" + fileName, content);
        }

        // In a web application
        //private async Task<HttpResponseMessage> HttpResponseMessage(byte[] content)
        //{
        //    HttpResponseMessage response = new HttpResponseMessage(HttpStatusCode.OK);

        //    Stream stream = new MemoryStream(content);
        //    response.Content = new StreamContent(stream);
        //    response.Content.Headers.ContentDisposition = new System.Net.Http.Headers.ContentDispositionHeaderValue("attachment");
        //    response.Content.Headers.ContentDisposition.FileName = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");
        //    response.Content.Headers.ContentType = new MediaTypeHeaderValue("application/vnd.ms-excel");

        //    return response;
        //}
    }
}
