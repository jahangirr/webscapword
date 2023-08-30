using System;
using System.Collections.Generic;
using System.EnterpriseServices;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;
using static System.Net.Mime.MediaTypeNames;
using System.Xml;
using webscapword.Models;
using HtmlAgilityPack;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;
using Microsoft.SqlServer.Server;
using System.Runtime.InteropServices;
using System.Runtime.Remoting.Contexts;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Drawing;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using DocumentFormat.OpenXml.Math;

namespace webscapword.Controllers
{
    

 
    
    public class HomeController : Controller
    {
        
        public ActionResult Index()
        {
            
            return View();
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }

        

       

        [HttpPost]
        public JsonResult IndexPost( )
        {
            HttpPostedFileBase file = Request.Files[0];
            retValue r = ReplaceTextInWordDoc( file);
            return Json(r);
        }


        private static async Task<string> CallUrl(string fullUrl)
        {
            
            var response = "";
            try
            {
                using(HttpClient client = new HttpClient())
                {
                    response = await client.GetStringAsync(fullUrl).ConfigureAwait(false);
                }
               
            }
            catch(Exception ex)
            {
                response += ex.Message;
            }
            
            return response;
        }

        public static string StripHTML(string input)
        {
            if (input == null)
            {
                return string.Empty;
            }
            return Regex.Replace(input, "<.*?>", String.Empty);

        }

        public string RemoveAllUnwantedElementNodes(string html)
        {
            try
            {

                HtmlDocument document = new HtmlAgilityPack.HtmlDocument();
                document.LoadHtml(html);

                if (document.DocumentNode.InnerHtml.Contains("<img"))
                {
                    foreach (var eachNode in document.DocumentNode.SelectNodes("//img"))
                    {
                        eachNode.Remove();
                        //eachNode.Attributes.Remove("src"); //This only removes the src Attribute from <img> tag
                    }
                }

                if (document.DocumentNode.InnerHtml.Contains("<head"))
                {
                    foreach (var eachHNode in document.DocumentNode.SelectNodes("//head"))
                    {
                        eachHNode.Remove();
                        //eachNode.Attributes.Remove("src"); //This only removes the src Attribute from <img> tag
                    }
                }

                if (document.DocumentNode.InnerHtml.Contains("<footer"))
                {
                    foreach (var eachFNode in document.DocumentNode.SelectNodes("//footer"))
                    {
                        eachFNode.Remove();
                        //eachNode.Attributes.Remove("src"); //This only removes the src Attribute from <img> tag
                    }
                }

                if (document.DocumentNode.InnerHtml.Contains("<body"))
                {
                    foreach (var eachBNode in document.DocumentNode.SelectNodes("//body"))
                    {

                        eachBNode.Attributes.Remove("src");
                        eachBNode.Attributes.Remove("class");
                        eachBNode.Attributes.Remove("style");
                    }
                }

                if (document.DocumentNode.InnerHtml.Contains("<script"))
                {
                    foreach (var eachSNode in document.DocumentNode.SelectNodes("//script"))
                    {

                        eachSNode.Remove();

                    }
                }

                if (document.DocumentNode.InnerHtml.Contains("<style"))
                {
                    foreach (var eachSNode in document.DocumentNode.SelectNodes("//style"))
                    {

                        eachSNode.Remove();

                    }
                }

                html = document.DocumentNode.InnerHtml;
                string sHTML = StripHTML(html);
                sHTML = "<html><body>" + sHTML + "</body></html>";
                return sHTML;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        
        public retValue ReplaceTextInWordDoc(HttpPostedFileBase file)
        {
            // URL: 

            retValue rValue = new retValue();
            rValue.showValues = new List<showValue>();
            List<string> offlineWordList = new List<string>();
            string url = "";

            string extension = System.IO.Path.GetExtension(file.FileName);


            string fname = System.IO.Path.Combine(Server.MapPath("~/Uploads/"), "111" + extension);
            string physicalPath = fname.Replace("\\", "/");
            
            if ( System.IO.File.Exists(physicalPath))
            {
                System.IO.File.Delete(physicalPath);
            }
            file.SaveAs(physicalPath);


            bool toBeBreak = false;

            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(physicalPath, false ))
            {
                bool found = false;
                foreach (var headerPart in wordDoc.MainDocumentPart.HeaderParts)
                {
                    //Gets the text in headers
                    foreach (var currentText in headerPart.RootElement.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>())
                    {
                        if (currentText.Text.Contains("http"))
                        {
                            int i = currentText.Text.IndexOf("http");
                            url = currentText.Text.Substring(i);
                            found = true;
                            break;
                        }
                    }
                    if(found == true)
                    {
                        break;
                    }
                }
                
                if(found == false)
                {
                    DocumentFormat.OpenXml.Wordprocessing.Body body = wordDoc.MainDocumentPart.Document.Body;
                    foreach (var para in body.Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>())
                    {
                        foreach (var run in para.Elements<Run>())
                        {

                            foreach (var text in run.Elements<Text>())
                            {
                                text.Text = Regex.Replace(text.Text, "<.*?>", String.Empty);

                                if (text.Text.Contains("http"))
                                {
                                    int i = text.Text.IndexOf("http");
                                    url = text.Text.Substring(i);
                                    toBeBreak = true;
                                    break;
                                }

                            }
                            if (toBeBreak)
                            {
                                break;
                            }
                        }
                        if (toBeBreak)
                        {
                            break;
                        }
                    }
                }
                
            }


            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(physicalPath, false))
            {
                DocumentFormat.OpenXml.Wordprocessing.Body body = wordDoc.MainDocumentPart.Document.Body;
                foreach (var para in body.Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>())
                {
                    foreach (var run in para.Elements<Run>())
                    {
                        foreach (var text in run.Elements<Text>())
                        {
                            foreach (var t in text.InnerText.Split(' '))
                                if (!offlineWordList.Contains(t.ToUpper().Trim()))
                                {
                                    offlineWordList.Add(t.ToUpper().Trim());
                                }

                        }
                    }
                }
            }

           
            var response = CallUrl(url).Result;
            rValue.url = url;


            HtmlDocument htmlDoc = new HtmlDocument();
            htmlDoc.LoadHtml(RemoveAllUnwantedElementNodes(response));
            List<string> onlineWordList = new List<string>();
            foreach (HtmlNode node in htmlDoc.DocumentNode.SelectNodes("//text()"))
            {
                foreach (var w in node.InnerText.Split(' '))
                    if (!onlineWordList.Contains(w.ToUpper().Trim()))
                    {
                        onlineWordList.Add(w.ToUpper().Trim());
                    }

            }

            foreach (var item in offlineWordList.Except(onlineWordList))
            {
                rValue.showValues.Add(new showValue() { missingText = item });
            }
            return rValue;
        }
    }
}