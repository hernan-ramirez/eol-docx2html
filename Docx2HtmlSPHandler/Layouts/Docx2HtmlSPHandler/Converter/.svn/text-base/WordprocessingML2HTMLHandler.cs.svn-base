using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Xsl;
using System.Xml;
using System.Xml.XPath;
using System.IO.Packaging;
using System.IO;
using System.Data;
using Microsoft.SharePoint.Utilities;
using System.Net;

namespace Docx2HtmlSPHandler.Layouts.Docx2HtmlSPHandler.Converter
{
    public class WordprocessingML2HTMLHandler
    {
        public string htmlResult;
        public string timeStamp;
        public MyPackage myPackage;
        public string siteUrl;
        public string InitializeMyPackage(Package paquete)
        {
            string strMainPartUri = "";
            myPackage = new MyPackage();
            using (myPackage.package = paquete)
            {
                PackageRelationshipCollection relationshipCollection = myPackage.package.GetRelationships();
                myPackage.siteUrl = this.siteUrl;
                // najiti Uri pro MainPart
                foreach (PackageRelationship rel in relationshipCollection)
                {
                    if (rel.RelationshipType == Constants.Relationships.Main)
                    {
                        strMainPartUri = rel.TargetUri.ToString();
                        myPackage.uriMainPart = new Uri(AddSlash(rel.TargetUri.ToString()), UriKind.Relative);
                        break;
                    }
                }
                int iLastSlash = strMainPartUri.LastIndexOf("/");
                myPackage.strMainPartDirectory = AddSlash(strMainPartUri.Substring(0, iLastSlash));

                if (myPackage.uriMainPart == null)
                {
                    myPackage.package.Close();

                    return string.Empty;
                }
                if (!myPackage.package.PartExists(myPackage.uriMainPart))
                {
                    myPackage.package.Close();

                    return string.Empty;
                }

                // Uri pro MainPart_rels                
                int iLastSlashPos = strMainPartUri.LastIndexOf('/');
                string strRelationsUri = strMainPartUri.Insert(iLastSlashPos, "/_rels") + ".rels";
                myPackage.uriMainPart_Rels = new Uri(AddSlash(strRelationsUri), UriKind.Relative);

                // najiti Uri pro dalsi casti
                if (myPackage.package.PartExists(myPackage.uriMainPart_Rels))
                {
                    PackagePart mainRelPart = myPackage.package.GetPart(myPackage.uriMainPart_Rels);
                    XPathDocument xPathDocMainRel = new XPathDocument(mainRelPart.GetStream());
                    XPathNavigator xn = xPathDocMainRel.CreateNavigator();
                    XPathNavigator xnPartPath;
                    string strPartPath;
                    // styles
                    xnPartPath = xn.SelectSingleNode(String.Format("//*[@Type ='{0}']/@Target", Constants.Relationships.Styles));
                    if (xnPartPath != null)
                    {
                        strPartPath = xnPartPath.ToString().Replace("/word/", string.Empty);
                        myPackage.uriStyles = new Uri(myPackage.strMainPartDirectory + AddSlash(strPartPath), UriKind.Relative);
                    }
                    // numbering
                    xnPartPath = xn.SelectSingleNode(String.Format("//*[@Type ='{0}']/@Target", Constants.Relationships.Numbering));
                    if (xnPartPath != null)
                    {
                        strPartPath = xnPartPath.ToString().Replace("/word/", string.Empty);
                        myPackage.uriNumbering = new Uri(myPackage.strMainPartDirectory + AddSlash(strPartPath), UriKind.Relative);
                    }
                    // theme
                    xnPartPath = xn.SelectSingleNode(String.Format("//*[@Type ='{0}']/@Target", Constants.Relationships.Theme));
                    if (xnPartPath != null)
                    {
                        strPartPath = xnPartPath.ToString().Replace("/word/", string.Empty);
                        myPackage.uriTheme = new Uri(myPackage.strMainPartDirectory + AddSlash(strPartPath), UriKind.Relative);
                    }
                    // settings
                    xnPartPath = xn.SelectSingleNode(String.Format("//*[@Type ='{0}']/@Target", Constants.Relationships.Settings));
                    if (xnPartPath != null)
                    {
                        strPartPath = xnPartPath.ToString().Replace("/word/", string.Empty);
                        myPackage.uriSettings = new Uri(myPackage.strMainPartDirectory + AddSlash(strPartPath), UriKind.Relative);
                    }
                    // footnotes
                    xnPartPath = xn.SelectSingleNode(String.Format("//*[@Type ='{0}']/@Target", Constants.Relationships.Footnotes));
                    if (xnPartPath != null)
                    {
                        strPartPath = xnPartPath.ToString().Replace("/word/", string.Empty);
                        myPackage.uriFootnotes = new Uri(myPackage.strMainPartDirectory + AddSlash(strPartPath), UriKind.Relative);
                    }
                    // endnotes
                    xnPartPath = xn.SelectSingleNode(String.Format("//*[@Type ='{0}']/@Target", Constants.Relationships.Endnotes));
                    if (xnPartPath != null)
                    {
                        strPartPath = xnPartPath.ToString().Replace("/word/", string.Empty);
                        myPackage.uriEndnotes = new Uri(myPackage.strMainPartDirectory + AddSlash(strPartPath), UriKind.Relative);
                    }
                }
                DataBindStyles();
                ConvertProcess(paquete);                
            }
            return myPackage.strOutputDirectory;
        }

        private void DataBindStyles()
        {
            XPathDocument xdocStyles;
            PackagePart p = myPackage.package.GetPart(myPackage.uriStyles);
            xdocStyles = new XPathDocument(p.GetStream());
            XPathNavigator xNav = xdocStyles.CreateNavigator();
            XmlNamespaceManager xMan = new XmlNamespaceManager(xNav.NameTable);
            xMan.AddNamespace("w", Constants.NameSpaces.W);
            DataTable dtStyles = new DataTable();
            dtStyles.Columns.Add(new DataColumn("styleId", typeof(string)));
            dtStyles.Columns.Add(new DataColumn("styleName", typeof(string)));
            DataRow dr = dtStyles.NewRow();
            dr["styleId"] = "none";
            dr["styleName"] = "-- choose style --";
            dtStyles.Rows.Add(dr);
            foreach (XPathNavigator xStyle in xNav.Select("//w:style[@w:type = 'paragraph']", xMan))
            {
                DataRow drStyle = dtStyles.NewRow();
                drStyle["styleId"] = xStyle.GetAttribute("styleId", Constants.NameSpaces.W);
                drStyle["styleName"] = xStyle.SelectSingleNode("w:name/@w:val", xMan).ToString();
                dtStyles.Rows.Add(drStyle);
            }
        }

        private static string AddSlash(string str)
        {
            // metoda GetPart(Uri partUri) chce, aby Uri zacinalo lomitkem
            if (!str.StartsWith("/"))
            {
                str = "/" + str;
            }
            return str;
        }

        public void Transform(XPathDocument xPathDoc, string strXslPath)
        {
            try
            {                
                XslCompiledTransform xslTrans = new XslCompiledTransform();                
                xslTrans.Load(strXslPath);
                DirectoryInfo di = Directory.CreateDirectory(myPackage.strOutputDirectory);
                di.Refresh();
                FileStream fs = new FileStream(Path.Combine(myPackage.strOutputDirectory, "index.html"), FileMode.OpenOrCreate);
                using (MultiXmlTextWriter multiWriter = new MultiXmlTextWriter(fs, Encoding.UTF8))
                {
                    XsltArgumentList args = new XsltArgumentList();
                    multiWriter.Formatting = Formatting.Indented;
                    // kdyz ma hlavni cast ralations part, pripojim k transformaci objekt urn:convertor.helper
                    Helper hlpr = new Helper(myPackage);
                    if (myPackage.package.PartExists(myPackage.uriMainPart_Rels))
                    {
                        args.AddExtensionObject("urn:convertor.helper", hlpr);
                    }
                    string strDividingStyles = "dividingStyle";
                    myPackage.dividingStyle = strDividingStyles;
                    args.AddParam("DivideBySection", String.Empty, myPackage.bDivideBySections);
                    args.AddParam("DividingStyle", String.Empty, myPackage.dividingStyle);
                    xslTrans.Transform(xPathDoc, args, multiWriter);
                    fs.Dispose();
                    StreamReader file = File.OpenText(Path.Combine(myPackage.strOutputDirectory, "index.html"));
                    htmlResult = file.ReadToEnd();
                    file.Dispose();
                    File.Delete(Path.Combine(myPackage.strOutputDirectory, "index.html"));
                    //multiWriter.Close();                    
                }
            }
            catch (Exception e)
            {

            }
        }

        private void ConvertProcess(Package paquete)
        {
            using (myPackage.package = paquete)
            {
                // transformace MainPart
                Random random = new Random();
                int id = random.Next(int.MaxValue);
                myPackage.strOutputDirectory = (@"C:\Docx2HtmlTemp\") + this.timeStamp;
                myPackage.strStylesheetsPath = SPUtility.GetGenericSetupPath(@"Template\Layouts\Docx2HtmlSPHandler\Stylesheets\") + "Styles.xslt";
                PackagePart mainPart = myPackage.package.GetPart(myPackage.uriMainPart);
                XPathDocument xPathDoc = new XPathDocument(mainPart.GetStream());
                Transform(xPathDoc, SPUtility.GetGenericSetupPath(@"Template\Layouts\Docx2HtmlSPHandler\Stylesheets\") + "MainPart.xslt");
                MakeCssFile();
                myPackage.package.Close();
            }
        }

        public void MakeCssFile()
        {
            if (myPackage.uriStyles != null)
            {
                PackagePart stylePart = myPackage.package.GetPart(myPackage.uriStyles);
                XPathDocument xdocStyle = new XPathDocument(stylePart.GetStream());
                XslCompiledTransform xslTrans = new XslCompiledTransform();
                xslTrans.Load(myPackage.strStylesheetsPath);
                XmlTextWriter xWriter = new XmlTextWriter(Path.Combine(myPackage.strOutputDirectory, "styles.css"), Encoding.UTF8);
                XsltArgumentList args = new XsltArgumentList();
                Helper hlpr = new Helper(myPackage);
                args.AddExtensionObject("urn:convertor.helper", hlpr);
                args.AddParam("DefaultStyleParagraph", String.Empty, hlpr.DefaultStyleParagraph);
                args.AddParam("DefaultStyleRun", String.Empty, hlpr.DefaultStyleRun);
                args.AddParam("DefaultStyleTable", String.Empty, hlpr.DefaultStyleTable);
                args.AddParam("DefaultStyleNumbering", String.Empty, hlpr.DefaultStyleNumbering);
                xslTrans.Transform(xdocStyle, args, xWriter);
                xWriter.Close();
            }
        }
    }
}
