using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Text;
using System.Xml;
using System.Xml.XPath;
using System.Xml.Xsl;
using System.Configuration;
using System.Collections;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
namespace Docx2HtmlSPHandler.Layouts.Docx2HtmlSPHandler.Converter
{
    public class Helper
    {
        // string strPartUri;
        PackagePart relationsPart;
        XPathDocument xdocRel;
        XPathDocument xdocNum;
        XPathDocument xdocSettings;
        XPathDocument xdocTheme;
        XPathDocument xdocStyles;
        XPathDocument xdocFootnotes;
        XPathDocument xdocEndnotes;
        public string DefaultStyleParagraph = String.Empty;
        public string DefaultStyleRun = String.Empty;
        public string DefaultStyleTable = String.Empty;
        public string DefaultStyleNumbering = String.Empty;
        DataTable dtNumberedStyles;
        public MyPackage hMyPackage;
        public Helper(MyPackage myPackage)
        {
            hMyPackage = myPackage;
            if (!myPackage.package.PartExists(myPackage.uriMainPart_Rels))
            {
                return;
            }
            // relations
            relationsPart = myPackage.package.GetPart(myPackage.uriMainPart_Rels);
            xdocRel = new XPathDocument(relationsPart.GetStream());
            // numbering
            try
            {
                PackagePart p = myPackage.package.GetPart(myPackage.uriNumbering);
                xdocNum = new XPathDocument(p.GetStream());
            }
            catch
            {
                //Console.WriteLine("Info: Numbering part not found.");
            }
            // settings
            try
            {
                PackagePart p = myPackage.package.GetPart(myPackage.uriSettings);
                xdocSettings = new XPathDocument(p.GetStream());
            }
            catch
            {
                //Console.WriteLine("Info: Settings part not found.");
            }
            // theme
            try
            {
                PackagePart p = myPackage.package.GetPart(myPackage.uriTheme);
                xdocTheme = new XPathDocument(p.GetStream());
            }
            catch
            {
                //Console.WriteLine("Info: Theme part not found.");
            }
            // styles
            try
            {
                PackagePart p = myPackage.package.GetPart(myPackage.uriStyles);
                xdocStyles = new XPathDocument(p.GetStream());
            }
            catch
            {
                //Console.WriteLine("Info: Styles part not found.");
            }
            // footnotes
            try
            {
                PackagePart p = myPackage.package.GetPart(myPackage.uriFootnotes);
                xdocFootnotes = new XPathDocument(p.GetStream());
            }
            catch
            {
                //Console.WriteLine("Info: Footnotes part not found.");
            }
            // endnotes
            try
            {
                PackagePart p = myPackage.package.GetPart(myPackage.uriEndnotes);
                xdocEndnotes = new XPathDocument(p.GetStream());
            }
            catch
            {
                //Console.WriteLine("Info: Endnotes part not found.");
            }

            XPathNavigator xNavStyles = xdocStyles.CreateNavigator();
            if (xNavStyles != null)
            {
                // ziskani defaultnich stylu ... to nechapu, proc je nectu az pri transformaci styles.xml
                XmlNamespaceManager xMan = new XmlNamespaceManager(xNavStyles.NameTable);
                xMan.AddNamespace("w", Constants.NameSpaces.W);
                XPathNavigator xeParagraph = xNavStyles.SelectSingleNode("//w:style[@w:type = 'paragraph' and @w:default='1'][1]", xMan);
                XPathNavigator xeRun = xNavStyles.SelectSingleNode("//w:style[@w:type = 'character' and @w:default='1'][1]", xMan);
                XPathNavigator xeTable = xNavStyles.SelectSingleNode("//w:style[@w:type = 'table' and @w:default='1'][1]", xMan);
                XPathNavigator xeNumbering = xNavStyles.SelectSingleNode("//w:style[@w:type = 'numbering' and @w:default='1'][1]", xMan);
                if (xeParagraph != null)
                {
                    DefaultStyleParagraph = xeParagraph.GetAttribute("styleId", Constants.NameSpaces.W);
                }
                if (xeRun != null)
                {
                    DefaultStyleRun = xeRun.GetAttribute("styleId", Constants.NameSpaces.W);
                }
                if (xeTable != null)
                {
                    DefaultStyleTable = xeTable.GetAttribute("styleId", Constants.NameSpaces.W);
                }
                if (xeNumbering != null)
                {
                    DefaultStyleNumbering = xeNumbering.GetAttribute("styleId", Constants.NameSpaces.W);
                }
                // získání stylù s èíslováním
                if (xdocNum != null)
                {
                    dtNumberedStyles = new DataTable();
                    dtNumberedStyles.Columns.Add(new DataColumn("numId"));
                    dtNumberedStyles.Columns.Add(new DataColumn("styleId"));
                    dtNumberedStyles.Columns.Add(new DataColumn("level", typeof(int)));
                    dtNumberedStyles.Columns.Add(new DataColumn("counter", typeof(int)));
                    dtNumberedStyles.PrimaryKey = new DataColumn[] { dtNumberedStyles.Columns["styleId"] };
                    //dtNumberedStyles.PrimaryKey = new DataColumn[] { dtNumberedStyles.Columns["numId"], dtNumberedStyles.Columns["styleId"],dtNumberedStyles.Columns["level"]};
                    XPathNavigator xNavNumb = xdocNum.CreateNavigator();
                    XPathNodeIterator xniStylesWithNumbering = xNavStyles.Select("//w:style[w:pPr/w:numPr/w:numId]", xMan);
                    foreach (XPathNavigator xe in xniStylesWithNumbering)
                    {
                        string strStyleId = xe.GetAttribute("styleId", Constants.NameSpaces.W);
                        string strNumId = xe.SelectSingleNode("w:pPr/w:numPr/w:numId/@w:val", xMan).ToString();
                        XPathNavigator xNavNumInstance = xNavNumb.SelectSingleNode(String.Format("/w:numbering/w:num[@w:numId='{0}']", strNumId), xMan);
                        if (xNavNumInstance == null)
                        {
                            // numbering with this numId not defined
                            continue;
                        }
                        string strAbstractNumId = xNavNumInstance.SelectSingleNode("w:abstractNumId/@w:val", xMan).ToString();
                        XPathNavigator xpnNumFmt = xNavNumb.SelectSingleNode(String.Format("/w:numbering/w:abstractNum[@w:abstractNumId='{0}']/w:lvl[@w:ilvl='0']/w:numFmt/@w:val", strAbstractNumId), xMan);
                        string strNumFmt = "";
                        if (xpnNumFmt != null)
                        {
                            strNumFmt = xpnNumFmt.ToString();
                        }
                        XPathNavigator xpnLevel = xe.SelectSingleNode("w:pPr/w:numPr/w:ilvl/@w:val", xMan);
                        int iLevel;
                        if (xpnLevel == null)
                        {
                            iLevel = 0;
                        }
                        else
                        {
                            iLevel = Convert.ToInt16(xpnLevel.ToString());
                        }
                        DataRow dr = dtNumberedStyles.NewRow();
                        dr["numId"] = strNumId;
                        dr["styleId"] = strStyleId;
                        dr["level"] = iLevel;
                        if (strNumFmt == "bullet")
                        {
                            dr["counter"] = -1;
                        }
                        else
                        {
                            dr["counter"] = 0;
                        }
                        dtNumberedStyles.Rows.Add(dr);
                    }
                }
            }
        }

        public string ProcessImage(string strRelationId)
        {
            XPathNavigator xNav = xdocRel.CreateNavigator();
            XPathNavigator xe = xNav.SelectSingleNode(String.Format("//*[@Id ='{0}']", strRelationId));
            string strTarget = xe.GetAttribute("Target", "");
            string strTargetMode = xe.GetAttribute("TargetMode", "");
            if (strTargetMode != "External")
            {
                PackagePart imgPart = hMyPackage.package.GetPart(new Uri(hMyPackage.strMainPartDirectory + "/" + strTarget, UriKind.Relative));
                int iPos = strTarget.LastIndexOf("/");
                string strImgName = strTarget.Substring(iPos, strTarget.Length - iPos);
                Stream readStream = imgPart.GetStream(FileMode.Open, FileAccess.Read);
                Directory.CreateDirectory(Path.Combine(hMyPackage.strOutputDirectory, "media"));
                if (strTarget.EndsWith(".wmf"))
                {
                    Image img = Image.FromStream(readStream);
                    string strNewTarget = strTarget.Substring(0, strTarget.Length - 4) + "_converted.png";
                    int iSlash = strNewTarget.LastIndexOf("/");
                    string strNewImgName = strNewTarget.Substring(iSlash, strNewTarget.Length - iSlash);
                    img.Save(Path.Combine(hMyPackage.strOutputDirectory, "media" + strNewImgName), ImageFormat.Png);
                    strTarget = strNewTarget;
                }
                else
                {
                    FileStream writeStream = new FileStream(Path.Combine(hMyPackage.strOutputDirectory, "media" + strImgName), FileMode.Create, FileAccess.Write);
                    //img.Save
                    int Length = 256;
                    Byte[] buffer = new Byte[Length];
                    int bytesRead = readStream.Read(buffer, 0, Length);
                    // write the required bytes
                    while (bytesRead > 0)
                    {
                        writeStream.Write(buffer, 0, bytesRead);
                        bytesRead = readStream.Read(buffer, 0, Length);
                    }

                    writeStream.Close();
                }
                readStream.Close();
            }
            return strTarget;
        }

        public string GetListInfo(int iListId, int iLevel, string strProperty)
        {
            XPathNavigator xNavNum = xdocNum.CreateNavigator();
            XmlNamespaceManager xMan = new XmlNamespaceManager(xNavNum.NameTable);
            xMan.AddNamespace("w", Constants.NameSpaces.W);
            XPathNavigator xe1 = xNavNum.SelectSingleNode(String.Format("/w:numbering/w:num[@w:numId='{0}']", iListId), xMan);
            if (xe1 == null)
            {
                // error - w:num with this w:numId not found
                return "unknown";
            }
            string strAbstractNumId = xe1.SelectSingleNode("w:abstractNumId/@w:val", xMan).ToString();
            XPathNavigator xNavAbsNum = xNavNum.SelectSingleNode(String.Format("/w:numbering/w:abstractNum[@w:abstractNumId='{0}']", strAbstractNumId), xMan);
            switch (strProperty)
            {
                case "type":
                    {
                        string strType = xNavAbsNum.SelectSingleNode(String.Format("w:lvl[@w:ilvl='{0}']/w:numFmt/@w:val", iLevel), xMan).ToString();
                        if (strType == "decimal" ||
                            strType == "lowerLetter" ||
                            strType == "lowerRoman" ||
                            strType == "upperLetter" ||
                            strType == "upperRoman"
                            ) return "numbered";
                        else
                        {
                            return "bullet";
                        }
                    }
                case "ind":
                    {
                        return xNavAbsNum.SelectSingleNode(String.Format("w:lvl[@w:ilvl='{0}']/w:pPr/w:ind/@w:left", iLevel), xMan).ToString();
                    }
                case "numType":
                    {
                        return xNavAbsNum.SelectSingleNode(String.Format("w:lvl[@w:ilvl='{0}']/w:numFmt/@w:val", iLevel), xMan).ToString();
                    }
                case "start":
                    {
                        return xNavAbsNum.SelectSingleNode(String.Format("w:lvl[@w:ilvl='{0}']/w:start/@w:val", iLevel), xMan).ToString();
                    }
                default:
                    {
                        return "unknown";
                    }
            }
        }

        public int GetDefaultTabStops()
        {
            XPathNavigator xNavSettings = xdocSettings.CreateNavigator();
            XmlNamespaceManager xMan = new XmlNamespaceManager(xNavSettings.NameTable);
            xMan.AddNamespace("w", Constants.NameSpaces.W);
            return int.Parse(xNavSettings.SelectSingleNode("/w:settings/w:defaultTabStop/@w:val", xMan).ToString());
        }

        public string GetThemeFont(string strType)
        {
            XPathNavigator xNavTheme = xdocTheme.CreateNavigator();
            XmlNamespaceManager xMan = new XmlNamespaceManager(xNavTheme.NameTable);
            xMan.AddNamespace("a", Constants.NameSpaces.A);
            if (strType == "major")
            {
                return xNavTheme.SelectSingleNode("/a:theme/a:themeElements/a:fontScheme/a:majorFont/a:latin/@typeface", xMan).ToString();
            }
            else
            {
                return xNavTheme.SelectSingleNode("/a:theme/a:themeElements/a:fontScheme/a:minorFont/a:latin/@typeface", xMan).ToString();
            }
        }

        public string GetHyperlinkTarget(string strRelationId)
        {
            XPathNavigator xNav = xdocRel.CreateNavigator();
            XPathNavigator xe = xNav.SelectSingleNode(String.Format("//*[@Id ='{0}']", strRelationId));
            if (xe == null) return "targetnotfound";
            return xe.GetAttribute("Target", "");
        }

        public bool GetBullettingDefinedInStyle(string strStyle)
        {
            if (dtNumberedStyles == null)
            {
                return false;
            }
            DataRow drv = dtNumberedStyles.Rows.Find(strStyle);
            if (drv != null && (int)drv["counter"] == -1)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public string GetNumberingDefinedInStyle(string strStyle)
        {
            string strNumber = String.Empty;
            if (dtNumberedStyles == null)
            {
                return strNumber;
            }
            DataRow dr = dtNumberedStyles.Rows.Find(strStyle);
            if (dr == null)
            {
                return strNumber;
            }
            string strNumId = (string)dr["numId"];
            int iLevel = (int)dr["level"];
            if ((int)dr["counter"] != -1)
            {
                dr["counter"] = (int)dr["counter"] + 1;
            }

            // vytvorim cislo
            dtNumberedStyles.DefaultView.Sort = "level";
            foreach (DataRowView drv in dtNumberedStyles.DefaultView)
            {
                if ((string)drv["numId"] == strNumId)
                {
                    if ((int)drv["counter"] == -1)
                    {
                        // styl je seznam, ale odrazkovej
                        strNumber = "bullet";
                    }
                    else
                    {
                        // snizim vsechny nizsi urovne daneho cislovani
                        if ((int)drv["level"] > iLevel)
                        {
                            drv["counter"] = 0;
                        }
                        // vytvorim cislo
                        else
                        {
                            if ((int)drv["counter"] == 0)
                            {
                                continue;
                            }
                            strNumber += drv["counter"] + ".";
                        }
                    }
                }
            }
            return strNumber;
        }

        public string GetFootnoteText(string strId)
        {
            XPathNavigator xNav = xdocFootnotes.CreateNavigator();
            XmlNamespaceManager xMan = new XmlNamespaceManager(xNav.NameTable);
            xMan.AddNamespace("w", Constants.NameSpaces.W);
            XPathNavigator xeFootnote = xNav.SelectSingleNode(String.Format("/w:footnotes/w:footnote[@w:id ='{0}']", strId), xMan);
            XPathNodeIterator xpi = xeFootnote.Select(".//w:t", xMan);
            StringBuilder sb = new StringBuilder();
            foreach (XPathNavigator xpn in xpi)
            {
                sb.Append(xpn.ToString());
            }
            return sb.ToString();
        }

        public string GetEndnoteText(string strId)
        {
            XPathNavigator xNav = xdocEndnotes.CreateNavigator();
            XmlNamespaceManager xMan = new XmlNamespaceManager(xNav.NameTable);
            xMan.AddNamespace("w", Constants.NameSpaces.W);
            XPathNavigator xeEndnote = xNav.SelectSingleNode(String.Format("/w:endnotes/w:endnote[@w:id ='{0}']", strId), xMan);
            XPathNodeIterator xpi = xeEndnote.Select(".//w:t", xMan);
            StringBuilder sb = new StringBuilder();
            foreach (XPathNavigator xpn in xpi)
            {
                sb.Append(xpn.ToString());
            }
            return sb.ToString();
        }
    }
}
