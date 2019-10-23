using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.tool.xml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Web;
using System.Web.Mvc;
using WebApplication6.Utility;

namespace WebApplication6.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult About()
        {
            return View();
        }
        public ActionResult PDF()
        {
            UriBuilder uriBuilder = new UriBuilder(Request.Url)
            {
                Path = Url.Action("DownloadPdf")
            };
            //string physicalPath = Server.MapPath(Server.UrlDecode(uriBuilder.ToString()));

            return Redirect(Url.Content(Server.UrlDecode(uriBuilder.ToString())));
        }
        public ActionResult DownloadPdf()
        {
            
            WebClient wc = new WebClient();
            //從網址下載Html字串
            UriBuilder uriBuilder = new UriBuilder(Request.Url)
            {
                Path = Url.Action("About")
            };

            string htmlText = wc.DownloadString(uriBuilder.ToString());
            byte[] pdfFile = this.ConvertHtmlTextToPDF(htmlText);

            return File(pdfFile, "application/pdf", "範例PDF檔.pdf");
        }

        public byte[] ConvertHtmlTextToPDF(string htmlText)
        {
            if (string.IsNullOrEmpty(htmlText))
            {
                return null;
            }
            //避免當htmlText無任何html tag標籤的純文字時，轉PDF時會掛掉，所以一律加上<p>標籤
            htmlText = "<p>" + htmlText + "</p>";

            MemoryStream outputStream = new MemoryStream();//要把PDF寫到哪個串流
            byte[] data = Encoding.UTF8.GetBytes(htmlText);//字串轉成byte[]
            MemoryStream msInput = new MemoryStream(data);
            Document doc = new Document();//要寫PDF的文件，建構子沒填的話預設直式A4
            PdfWriter writer = PdfWriter.GetInstance(doc, outputStream);
            //指定文件預設開檔時的縮放為100%
            PdfDestination pdfDest = new PdfDestination(PdfDestination.XYZ, 0, doc.PageSize.Height, 1f);
            //開啟Document文件 
            doc.Open();
            //使用XMLWorkerHelper把Html parse到PDF檔裡
            XMLWorkerHelper.GetInstance().ParseXHtml(writer, doc, msInput, null, Encoding.UTF8, new UnicodeFontFactory());
            //將pdfDest設定的資料寫到PDF檔
            PdfAction action = PdfAction.GotoLocalPage(1, pdfDest, writer);
            writer.SetOpenAction(action);
            doc.Close();
            msInput.Close();
            outputStream.Close();
            //回傳PDF檔案 
            return outputStream.ToArray();

        }
        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
        public ActionResult GeneratePDF()
        {
            return new Rotativa.ActionAsPdf("About");
        }
    }
}