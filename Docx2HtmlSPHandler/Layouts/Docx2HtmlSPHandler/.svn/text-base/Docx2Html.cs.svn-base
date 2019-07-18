using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web;
using System.IO;
using System.Xml.Xsl;
using System.Xml;
using DocumentFormat.OpenXml.Packaging;
using Microsoft.SharePoint;
using System.IO.Packaging;
using Docx2HtmlSPHandler.Layouts.Docx2HtmlSPHandler.Converter;
using Microsoft.SharePoint.Utilities;
using System.Threading;
using DocumentFormat.OpenXml;

namespace Docx2HtmlSPHandler.Layouts.Docx2HtmlSPHandler
{
    public class Docx2Html : IHttpHandler
    {
        private string timeStamp;
        private string anchorRequest;
        private string URL = string.Empty;
        private SPUser usuario;
        private string buscar = string.Empty;
        private SPFile file;
        private SPFile fileHtml;
        private SPWeb spweb;
        private string redirect = string.Empty;
        private WordprocessingML2HTMLHandler word = new WordprocessingML2HTMLHandler();
        /// <summary>
        /// You will need to configure this handler in the web.config file of your 
        /// web and register it with IIS before being able to use it. For more information
        /// see the following link: http://go.microsoft.com/?linkid=8101007
        /// </summary>


        public bool IsReusable
        {
            // Return false in case your Managed Handler cannot be reused for another request.
            // Usually this would be false in case you have some state information preserved per request.
            get { return false; }
        }

        public void ProcessRequest(HttpContext context)
        {
            URL = "http://" + context.Request.ServerVariables["HTTP_HOST"] + "/sitios/ver";
            spweb = SPContext.Current.Web;
            SPSecurity.RunWithElevatedPrivileges(delegate()
            {
                try
                {
                    ArmarUrl(context);
                    if (file.Exists)
                    {
                        Llamada(context);
                        context.Response.ContentType = "text/html";
                        context.Response.Redirect(redirect, false);
                    }
                    else
                    {
                        LoggingService.LogError("Docx2HtmlSPHandlerLog", string.Concat("No se encontró el documento: ", context.Request.Url.OriginalString));
                        context.Response.ContentType = "text/htm";
                        context.Response.Redirect(URL + "/Menu/noexist.htm?doc=" + context.Request.Url.OriginalString.Split('/').Last(), false);
                    }
                }
                catch (Exception ex)
                {
                    LoggingService.LogError("Docx2HtmlSPHandlerLog", string.Concat("Error en la transformación | Mensaje: ", ex.Message, " StackTrace: " + ex.StackTrace));
                    //string html = new StreamReader(SPUtility.GetGenericSetupPath(@"Template\Layouts\Docx2HtmlSPHandler\redirect.htm")).ReadToEnd().Replace("http://www.google.com.ar", context.Request.Url.OriginalString);
                    context.Response.ContentType = "text/htm";
                    //context.Response.Write(html);
                    context.Response.Redirect("http://" + context.Request.ServerVariables["HTTP_HOST"] + "/sitios/ver/Menu/redirect.htm?url=" + context.Request.Url.OriginalString, false);
                }
            });
        }

        private void Llamada(HttpContext context)
        {
            using (SPSite site = new SPSite(URL))
            {
                using (SPWeb web = site.OpenWeb())
                {
                    usuario = site.SystemAccount;
                    fileHtml = web.GetFile(URL + "/html/" + timeStamp + ".html");
                    bool existeHtml = fileHtml.Exists;
                    int htmlVersion = 0;
                    if (existeHtml)
                    {
                        htmlVersion = fileHtml.Item["VersionDocx"] != null ? int.Parse(fileHtml.Item["VersionDocx"].ToString()) : 0;
                    }
                    if (!existeHtml)
                    {
                        Procesar(file, site, web);
                    }
                    else
                    {
                        if (existeHtml && file.MajorVersion > htmlVersion)
                        {
                            Procesar(file, site, web);
                        }
                    }
                }
            }

        }

        private void ArmarUrl(HttpContext context)
        {
            string url = context.Request.Url.OriginalString.Replace("html", string.Empty);
            if (context.Request.QueryString.HasKeys())
            {
                var queryString = context.Request.QueryString;
                string[] keys = queryString.AllKeys;
                for (int i = 0; i < keys.Length; i++)
                {
                    keys[i] = keys[i] + "=" + queryString[keys[i]];
                }
                buscar = String.Join("&", keys);
            }
            if (context.Request.Url.OriginalString.Contains("#"))
            {
                anchorRequest = context.Request.Url.OriginalString.Split('#').Last();
            }
            url = url.Contains("?") ? url.Substring(0, url.LastIndexOf("?")) : url;
            file = spweb.GetFile(url);
            word.timeStamp = file.Name.Split('.').First();
            timeStamp = word.timeStamp;
            redirect = URL + "/html/" + timeStamp + ".html" + (!string.IsNullOrEmpty(anchorRequest) ? "#" + anchorRequest : string.Empty) + (!string.IsNullOrEmpty(buscar) ? "?" + buscar : string.Empty);
        }

        private void Procesar(SPFile file, SPSite site, SPWeb web)
        {
            word.siteUrl = site.Url;
            string carpeta = word.InitializeMyPackage(Package.Open(file.OpenBinaryStream()));
            string html = RefactorHtml(word.htmlResult, file.OpenBinary());
            StreamWriter writer = new StreamWriter(carpeta + "/" + timeStamp + ".html");
            writer.Write(html);
            writer.Close();
            SubirArchivos(carpeta, web);
            DirectoryInfo di = new DirectoryInfo(carpeta);
            di.Delete(true);
        }

        private void SubirArchivos(string carpeta, SPWeb web)
        {
            web.AllowUnsafeUpdates = true;
            StreamReader html = new StreamReader(carpeta + "\\" + word.timeStamp + ".html");
            if (!fileHtml.Exists)
            {
                web.Files.Add(URL + "/html/" + timeStamp + ".html", html.BaseStream, usuario, usuario, DateTime.Now, DateTime.Now);
                fileHtml = web.GetFile(URL + "/html/" + timeStamp + ".html");
            }
            else
            {
                fileHtml.SaveBinary(html.BaseStream);
                fileHtml.Update();
            }
            html.Close();
            fileHtml.Item["VersionDocx"] = file.MajorVersion;
            fileHtml.Item.SystemUpdate();
            if (!web.GetFile(URL + "/css/" + word.timeStamp + "/styles.css").Exists)
            {
                EnsureParentFolder(web, URL + "/css/" + timeStamp + "/styles.css");
                StreamReader css = new StreamReader(carpeta + "\\" + "styles.css");
                web.Files.Add(URL + "/css/" + word.timeStamp + "/styles.css", css.BaseStream, usuario, usuario, DateTime.Now, DateTime.Now);
                css.Close();
            }
            if (Directory.Exists(carpeta + "/media"))
            {
                string[] files = Directory.GetFiles(carpeta + "/media");
                foreach (string pathImagen in files)
                {
                    string nombre = pathImagen.Split('\\').Last();
                    if (!web.GetFile(URL + "/img/" + word.timeStamp + "/" + nombre).Exists)
                    {
                        EnsureParentFolder(web, URL + "/img/" + word.timeStamp + "/" + nombre);
                        StreamReader imagen = new StreamReader(pathImagen);
                        web.Files.Add(URL + "/img/" + word.timeStamp + "/" + nombre, imagen.BaseStream, usuario, usuario, DateTime.Now, DateTime.Now);
                        imagen.Close();
                    }
                }
            }
        }

        /// <summary>
        /// Se asegura de que las carpetas padres existan.
        /// O sea que toda la ruta o ubicacion exista dentro del site indicado
        /// </summary>
        /// <param name="parentSite"></param>
        /// <param name="destinUrl"></param>
        /// <returns></returns>
        public string EnsureParentFolder(SPWeb parentSite, string destinUrl)
        {
            destinUrl = parentSite.GetFile(destinUrl).Url;

            int index = destinUrl.LastIndexOf("/");
            string parentFolderUrl = string.Empty;

            if (index > -1)
            {
                parentFolderUrl = destinUrl.Substring(0, index);

                SPFolder parentFolder = parentSite.GetFolder(parentFolderUrl);

                if (!parentFolder.Exists)
                {
                    SPFolder currentFolder = parentSite.RootFolder;

                    foreach (string folder in parentFolderUrl.Split('/'))
                    {
                        currentFolder = currentFolder.SubFolders.Add(folder);
                        try
                        {
                            currentFolder.Item["Title"] = folder;
                            currentFolder.Item.Update();
                        }
                        catch (Exception)
                        {
                        }
                    }
                }
            }
            return parentFolderUrl;
        }

        private string RefactorHtml(string htmlResult, byte[] byteArray)
        {
            string html = htmlResult.Replace("styles.css", "/sitios/ver/Css/" + word.timeStamp + "/styles.css").Replace("media/image", "/sitios/ver/Img/" + word.timeStamp + "/image");
            string archivoOrigen = word.timeStamp + ".docx";
            //html = html.Replace("<body>", "<body>\n<form><input type=\"hidden\" id=\"ArchivoOrigen\" value=\"" + archivoOrigen + "\" /></form>\n");
            OpenXmlElement[] oxe;
            HyperlinkRelationship[] hlr;
            using (MemoryStream mem = new MemoryStream())
            {
                mem.Write(byteArray, 0, (int)byteArray.Length);
                using (WordprocessingDocument wordDoc =
                    WordprocessingDocument.Open(mem, true))
                {
                    hlr = wordDoc.MainDocumentPart.HyperlinkRelationships.ToArray();
                    var docPart = wordDoc.MainDocumentPart;
                    oxe = docPart.Document.Descendants().Where(a => a.LocalName.Equals("hyperlink")).ToArray();
                }
            }
            Dictionary<string, string> sos = new Dictionary<string, string>();
            foreach (var fa in oxe.Select(a => a.GetAttributes()))
            {
                if (fa.Where(ds => ds.LocalName.Equals("id")).Count() > 0 && fa.Where(ds => ds.LocalName.Equals("anchor")).Count() > 0)
                {
                    sos.Add(fa.First(ds => ds.LocalName.Equals("id")).Value, "href=\"#" + fa.First(ds => ds.LocalName.Equals("anchor")).Value + "\"");
                }
            }
            List<string> anchorsUsados = new List<string>();
            foreach (var fe in hlr)
            {
                Uri uri = fe.Uri;
                if (uri.OriginalString.Contains(".docx"))
                {
                    if (sos.ContainsKey(fe.Id))
                    {
                        html = html.Replace(sos[fe.Id], sos[fe.Id].Insert(sos[fe.Id].LastIndexOf("#"), fe.Uri.AbsolutePath));
                    }
                    //anchorsUsados.Add(sos[fe.Id]);
                }
            }
            string[] result = html.Split('\n');
            for (int i = 0; i <= result.Length - 1; i++)
            {
                if (i != 0 && i != result.Length - 1)
                {
                    //if (result[i].Contains("<a") && result[i].Contains("href"))
                    //{
                    //    var tags = result[i].Split('\"');
                    //    string hrefValue = string.Empty;
                    //    for (int j = 0; j < tags.Length; j++)
                    //    {
                    //        if (tags[j].Contains("href"))
                    //        {
                    //            hrefValue = tags[j + 1];
                    //            break;
                    //        }
                    //    }
                    //    Uri uri = hlr.First(h => h.Id.Equals(sos[hrefValue])).Uri;
                    //    if (uri.OriginalString.Contains(".docx"))
                    //    {
                    //        //html = html.Replace("#" + sos[fe.Id], fe.Uri.AbsolutePath + "#" + sos[fe.Id]);                            
                    //        result[i] = result[i].Replace(hrefValue, uri.AbsolutePath + hrefValue);
                    //    }
                    //}
                    if (result[i].Contains(".docx"))
                    {
                        int index = result[i].LastIndexOf(".docx");
                        long tm = 0;
                        bool parseo = long.TryParse(result[i].Substring(index - "20110807084607459".Length, "20110807084607459".Length), out tm);
                        if (parseo)
                        {
                            result[i] = result[i].Replace(".docx", ".docxhtml");
                        }
                    }
                }
            }
            return String.Join("\n", result).Replace("<body>", "<body>\n<form><input type=\"hidden\" id=\"ArchivoOrigen\" value=\"" + file.Url.Split('/').First() + "/" + archivoOrigen + "\" /></form>\n");
        }
    }
}

//if (result[i].Contains("</span>"))
//{
//    result[i] = result[i].Replace("</span>", "</span><!-- ");
//}
//if (result[i - 1].Contains("</span><!--"))
//{
//    result[i] = result[i].Insert(result[i].IndexOf("<"), " -->");
//}                    
//if (result[i].Contains("<a") && result[i].Contains("href"))
//{
//    var tags = result[i].Split('\"');
//    string hrefValue = string.Empty;
//    for (int j = 0; j < tags.Length; j++)
//    {
//        if (tags[j].Contains("href"))
//        {
//            hrefValue = tags[j + 1];
//            break;
//        }
//    }
//    if (!hrefValue.Contains(WordprocessingML2HTMLHandler.timeStamp))
//    {
//        if (hrefValue.Contains("#"))
//        {
//            var atributo = oxe.First(s => s.GetAttributes().Where(a => a.LocalName.Equals("anchor") && a.Value.Equals(hrefValue.Substring(1))).Count() == 1);
//            var aid = atributo.GetAttributes().Where(s => s.LocalName.Equals("id"));
//            string id = aid.Count() > 0 ? aid.First().Value : string.Empty;
//            string anchor = atributo.GetAttributes().First(s => s.LocalName.Equals("anchor")).Value;
//            if (!string.IsNullOrEmpty(id))
//            {
//                Uri uri = hlr.First(s => s.Id.Equals(id)).Uri;
//                if (uri.OriginalString.Contains(".docx"))
//                {
//                    result[i] = result[i].Replace(hrefValue, uri.AbsolutePath + "html" + "#" + anchor);
//                    if (!result[i].Contains(WordprocessingML2HTMLHandler.timeStamp))
//                    {
//                        if (result[i].Contains("target=\"_self\""))
//                        {
//                            result[i] = result[i].Replace("target=\"_self\"", string.Empty);
//                        }
//                    }
//                }
//            }
//        }
//        else
//        {
//            if (hrefValue.Split('/').Last().Length == "20110807084607459.docx".Length && hrefValue.Split('/').Last().EndsWith(".docx"))
//            {
//                result[i] = result[i].Replace(".docx", ".docxhtml");
//            }
//        }
//    }
//}