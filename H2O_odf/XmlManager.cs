using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
//using System.Windows.Forms;
using System.Xml;

namespace H2O__
{
    class XmlManager
    {
        public FileNode root;
        private const string header_style = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
        private const string header_fo = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";
        private const string header_office = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
        private const string header_text = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
        private static int numP = 1;
        private static int numSpan = 1;



        public XmlManager()
        {
        }

        public void CreateODT()
        {
            string path = @"C:\Users\park\source\repos\H2O_odf\H2O_odf";

            root = new FolderNode("New File");
            root.path = path + @"\" + root.name;


            //Create File

            string[] list1 = { "Configurations2", "META-INF", "Thumbnails" };
            
            foreach(string s in list1)
            {
                FolderNode node = new FolderNode(s, root);
            }

            FolderNode Configurations2 = (FolderNode)root.child["Configurations2"];

            string[] list2 = { "accelerator", "floater", "images", "menubar", "popupmenu", "progressbar", "statusbar", "toolbar", "toolpanel" };

            foreach (string s in list2)
            {
                FolderNode node = new FolderNode(s, Configurations2);
            }

            FolderNode images = (FolderNode)Configurations2.child["images"];

            FolderNode Bitmaps = new FolderNode("Bitmaps", images);


            //CreateXML

            string[] file1 = { "content.xml", "manifest.xml", "meta.xml", "mimetype", "settings.xml", "styles.xml" };

            foreach (string s in file1)
            { 
                XmlNode node = new XmlNode(s, root);
                
                node.LoadXml(path + @"\data" + @"\" + node.name);
            }

            FolderNode META_INF = (FolderNode)root.child["META-INF"];

            XmlNode manifest = new XmlNode("manifest.xml", META_INF);
            manifest.LoadXml(path + @"\data" + @"\META-INF" + @"\" + manifest.name);


            FolderNode accelerator = (FolderNode)Configurations2.child["accelerator"];

            XmlNode current = new XmlNode("current.xml", accelerator);
            current.LoadXml(path + @"\data" + @"\Configurations2" + @"\accelerator" + @"\" + current.name);

        }

        public void SaveODT(FileNode node)
        {
            if (node.GetType() == typeof(FolderNode))
            {
                DirectoryInfo di = new DirectoryInfo(node.path);
                if(!di.Exists) { di.Create(); }
            }
            else if (node.GetType() == typeof(XmlNode))
            {
                if(((XmlNode)node).name == "mimetype")
                {
                    StreamWriter sw = File.CreateText(((XmlNode)node).path);
                    sw.Write("application/vnd.oasis.opendocument.text");
                    sw.Close();
                }
                else
                {
                    ((XmlNode)node).doc.Save(((XmlNode)node).path);
                }
            }

            if (node.child == null)
                return;

            foreach(FileNode child in node.child.Values)
            {
                SaveODT(child);
            }

        }



        public void SetFontSize(string name, float size)
        {
            {
                XmlNode content = (XmlNode)root.child["content.xml"];

                XmlDocument doc = content.doc;

                XmlNodeList list = doc.GetElementsByTagName("style", header_style);

                foreach (XmlElement e in list)
                {
                    string check_name = e.GetAttribute("name", header_style);

                    if (check_name.Equals(name))
                    {
                        XmlElement e1 = (XmlElement)e.GetElementsByTagName("text-properties", header_style).Item(0);
                        if (e1 == null)
                        {
                            e1 = doc.CreateElement("style:text-properties", header_style);
                        }
                        e1.SetAttribute("font-size", header_fo, size + "pt");
                        e1.SetAttribute("font-size-asian", header_fo, size + "pt");
                        e1.SetAttribute("font-size-complex", header_fo, size + "pt");

                        e.AppendChild(e1);
                        break;
                    }
                }
            }
        }

        public void SetFontColor(string name, string color)
        {
            {
                XmlNode content = (XmlNode)root.child["content.xml"];

                XmlDocument doc = content.doc;

                XmlNodeList list = doc.GetElementsByTagName("style", header_style);

                ((XmlElement)doc.GetElementsByTagName("document-content", header_office).Item(0)).SetAttribute("xmlns:loext", "urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0");

                foreach (XmlElement e in list)
                {
                    string check_name = e.GetAttribute("name", header_style);

                    if (check_name.Equals(name))
                    {
                        XmlElement e1 = (XmlElement)e.GetElementsByTagName("text-properties", header_style).Item(0);
                        if (e1 == null)
                        {
                            e1 = doc.CreateElement("style:text-properties", header_style);
                        }
                        e1.SetAttribute("color", header_fo, color);
                        e1.SetAttribute("opacity", "urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0", "100%");

                        e.AppendChild(e1);
                        break;
                    }
                }
            }
        }

        public void SetLetterSpace(string name, float space)
        {
            {
                XmlNode content = (XmlNode)root.child["content.xml"];

                XmlDocument doc = content.doc;

                XmlNodeList list = doc.GetElementsByTagName("style", header_style);

                foreach (XmlElement e in list)
                {
                    string check_name = e.GetAttribute("name", header_style);

                    if (check_name.Equals(name))
                    {
                        XmlElement e1 = (XmlElement)e.GetElementsByTagName("text-properties", header_style).Item(0);
                        if (e1 == null)
                        {
                            e1 = doc.CreateElement("style:text-properties", header_style);
                        }
                        e1.SetAttribute("letter-spacing", header_fo, space+"cm");

                        e.AppendChild(e1);
                        break;
                    }
                }
            }
        }

        public void SetFont(string name, string font)
        {
            {
                XmlNode content = (XmlNode)root.child["content.xml"];

                XmlDocument doc = content.doc;

                XmlNodeList list = doc.GetElementsByTagName("font-face", header_style);
                for(int i = 0; i < list.Count; i++)
                {
                    string fontname = ((XmlElement)list.Item(i)).GetAttribute("name", header_style);
                    if (fontname.Equals(font))
                        break;
                    if(i == list.Count - 1)
                    {
                        XmlNodeList fontlist = doc.GetElementsByTagName("font-face-decls", header_office);
                        XmlElement addfont = (XmlElement)fontlist.Item(0);
                        XmlElement e1 = doc.CreateElement("style:font-face", header_style);

                        e1.SetAttribute("name", header_style, font);
                        e1.SetAttribute("font-family", "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0", font);
                        e1.SetAttribute("font-family-generic", header_style, "system");
                        e1.SetAttribute("font-pitch", header_style, "variable");

                        addfont.AppendChild(e1);

                    }

                }

                list = doc.GetElementsByTagName("style", header_style);

                foreach (XmlElement e in list)
                {
                    string check_name = e.GetAttribute("name", header_style);

                    if (check_name.Equals(name))
                    {
                        XmlElement e1 = (XmlElement)e.GetElementsByTagName("text-properties", header_style).Item(0);
                        if (e1 == null)
                        {
                            e1 = doc.CreateElement("style:text-properties", header_style);
                        }
                        e1.SetAttribute("font-name", header_style, font);

                        e.AppendChild(e1);
                        break;
                    }
                }

                XmlNode styles = (XmlNode)root.child["styles.xml"];
                doc = styles.doc;

                for (int i = 0; i < list.Count; i++)
                {
                    string fontname = ((XmlElement)list.Item(i)).GetAttribute("name", header_style);
                    if (fontname.Equals(font))
                        break;
                    if (i == list.Count - 1)
                    {
                        XmlNodeList fontlist = doc.GetElementsByTagName("font-face-decls", header_office);
                        XmlElement addfont = (XmlElement)fontlist.Item(0);
                        XmlElement e1 = doc.CreateElement("style:font-face", header_style);

                        e1.SetAttribute("name", header_style, font);
                        e1.SetAttribute("font-family", "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0", font);
                        e1.SetAttribute("font-family-generic", header_style, "system");
                        e1.SetAttribute("font-pitch", header_style, "variable");

                        addfont.AppendChild(e1);

                    }
                }
            }
        }

        public void SetBold(string name)
        {
            XmlNode content = (XmlNode)root.child["content.xml"];

            XmlDocument doc = content.doc;

            XmlNodeList list =  doc.GetElementsByTagName("style", header_style);

            foreach(XmlElement e in list)
            {
                string check_name = e.GetAttribute("name", header_style);

                if (check_name.Equals(name))
                {
                    XmlElement e1 = (XmlElement)e.GetElementsByTagName("text-properties", header_style).Item(0);
                    if (e1 == null)
                    {
                        e1 = doc.CreateElement("style:text-properties", header_style);
                    }
                    e1.SetAttribute("font-weight", header_fo, "bold");
                    e1.SetAttribute("font-weight-asian", header_style, "bold");
                    e1.SetAttribute("font-weight-complex", header_style, "bold");

                    e.AppendChild(e1);
                    break;
                }
            }
        }

        public void SetItalic(string name)
        {
            XmlNode content = (XmlNode)root.child["content.xml"];

            XmlDocument doc = content.doc;

            XmlNodeList list = doc.GetElementsByTagName("style", header_style);

            foreach (XmlElement e in list)
            {
                string check_name = e.GetAttribute("name", header_style);

                if (check_name.Equals(name))
                {
                    XmlElement e1 = (XmlElement)e.GetElementsByTagName("text-properties", header_style).Item(0);
                    if (e1 == null)
                    {
                        e1 = doc.CreateElement("style:text-properties", header_style);
                    }
                    e1.SetAttribute("font-style", header_fo, "italic");
                    e1.SetAttribute("font-style-asian", header_style, "italic");
                    e1.SetAttribute("font-style-complex", header_style, "italic");

                    e.AppendChild(e1);
                    break;
                }
            }
        }

        public void SetUnderline(string name, string style = "solid", string color = "font-color")
        {
            XmlNode content = (XmlNode)root.child["content.xml"];

            XmlDocument doc = content.doc;

            XmlNodeList list = doc.GetElementsByTagName("style", header_style);

            foreach (XmlElement e in list)
            {
                string check_name = e.GetAttribute("name", header_style);

                if (check_name.Equals(name))
                {
                    XmlElement e1 = (XmlElement)e.GetElementsByTagName("text-properties", header_style).Item(0);
                    if (e1 == null)
                    {
                        e1 = doc.CreateElement("style:text-properties", header_style);
                    }
                    e1.SetAttribute("text-underline-style", header_style, style);
                    e1.SetAttribute("text-underline-width", header_style, "auto");
                    e1.SetAttribute("text-underline-color", header_style, color);

                    e.AppendChild(e1);
                    break;
                }
            }
        }

        public void SetOverline(string name, string style = "solid", string color = "font-color")
        {
            XmlNode content = (XmlNode)root.child["content.xml"];

            XmlDocument doc = content.doc;

            XmlNodeList list = doc.GetElementsByTagName("style", header_style);

            foreach (XmlElement e in list)
            {
                string check_name = e.GetAttribute("name", header_style);

                if (check_name.Equals(name))
                {
                    XmlElement e1 = (XmlElement)e.GetElementsByTagName("text-properties", header_style).Item(0);
                    if (e1 == null)
                    {
                        e1 = doc.CreateElement("style:text-properties", header_style);
                    }
                    e1.SetAttribute("text-overline-style", header_style, style);
                    e1.SetAttribute("text-overline-width", header_style, "auto");
                    e1.SetAttribute("text-overline-color", header_style, color);

                    e.AppendChild(e1);
                    break;
                }
            }
        }

        public void SetThroughline(string name, string style = "solid")
        {
            XmlNode content = (XmlNode)root.child["content.xml"];

            XmlDocument doc = content.doc;

            XmlNodeList list = doc.GetElementsByTagName("style", header_style);

            foreach (XmlElement e in list)
            {
                string check_name = e.GetAttribute("name", header_style);

                if (check_name.Equals(name))
                {
                    XmlElement e1 = (XmlElement)e.GetElementsByTagName("text-properties", header_style).Item(0);
                    if (e1 == null)
                    {
                        e1 = doc.CreateElement("style:text-properties", header_style);
                    }
                    e1.SetAttribute("text-line-through-style", header_style, style);
                    e1.SetAttribute("text-line-through-type", header_style, "single");

                    e.AppendChild(e1);
                    break;
                }
            }
        }

        public void SetOutline(string name)
        {
            XmlNode content = (XmlNode)root.child["content.xml"];

            XmlDocument doc = content.doc;

            XmlNodeList list = doc.GetElementsByTagName("style", header_style);

            foreach (XmlElement e in list)
            {
                string check_name = e.GetAttribute("name", header_style);

                if (check_name.Equals(name))
                {
                    XmlElement e1 = (XmlElement)e.GetElementsByTagName("text-properties", header_style).Item(0);
                    if (e1 == null)
                    {
                        e1 = doc.CreateElement("style:text-properties", header_style);
                    }
                    e1.SetAttribute("text-outline", header_style, "true");

                    e.AppendChild(e1);
                    break;
                }
            }
        }

        public void SetShadow(string name)
        {
            XmlNode content = (XmlNode)root.child["content.xml"];

            XmlDocument doc = content.doc;

            XmlNodeList list = doc.GetElementsByTagName("style", header_style);

            foreach (XmlElement e in list)
            {
                string check_name = e.GetAttribute("name", header_style);

                if (check_name.Equals(name))
                {
                    XmlElement e1 = (XmlElement)e.GetElementsByTagName("text-properties", header_style).Item(0);
                    if (e1 == null)
                    {
                        e1 = doc.CreateElement("style:text-properties", header_style);
                    }
                    e1.SetAttribute("text-shadow", header_fo, "1pt 1pt");

                    e.AppendChild(e1);
                    break;
                }
            }
        }

        public void SetRelief(string name, string type = "embossed")
        {
            XmlNode content = (XmlNode)root.child["content.xml"];

            XmlDocument doc = content.doc;

            XmlNodeList list = doc.GetElementsByTagName("style", header_style);

            foreach (XmlElement e in list)
            {
                string check_name = e.GetAttribute("name", header_style);

                if (check_name.Equals(name))
                {
                    XmlElement e1 = (XmlElement)e.GetElementsByTagName("text-properties", header_style).Item(0);
                    if (e1 == null)
                    {
                        e1 = doc.CreateElement("style:text-properties", header_style);
                    }
                    e1.SetAttribute("font-relief", header_style, type);//engraved, embossed

                    e.AppendChild(e1);
                    break;
                }
            }
        }

        public void SetSuper(string name)
        {
            XmlNode content = (XmlNode)root.child["content.xml"];

            XmlDocument doc = content.doc;

            XmlNodeList list = doc.GetElementsByTagName("style", header_style);

            foreach (XmlElement e in list)
            {
                string check_name = e.GetAttribute("name", header_style);

                if (check_name.Equals(name))
                {
                    XmlElement e1 = (XmlElement)e.GetElementsByTagName("text-properties", header_style).Item(0);
                    if (e1 == null)
                    {
                        e1 = doc.CreateElement("style:text-properties", header_style);
                    }
                    e1.SetAttribute("text-position", header_style, "super 58%");
                    e.AppendChild(e1);
                    break;
                }
            }
        }

        public void SetSub(string name)
        {
            XmlNode content = (XmlNode)root.child["content.xml"];

            XmlDocument doc = content.doc;

            XmlNodeList list = doc.GetElementsByTagName("style", header_style);

            foreach (XmlElement e in list)
            {
                string check_name = e.GetAttribute("name", header_style);

                if (check_name.Equals(name))
                {
                    XmlElement e1 = (XmlElement)e.GetElementsByTagName("text-properties", header_style).Item(0);
                    if (e1 == null)
                    {
                        e1 = doc.CreateElement("style:text-properties", header_style);
                    }
                    e1.SetAttribute("text-position", header_style, "sub 58%");
                    e.AppendChild(e1);
                    break;
                }
            }
        }

        public void SetPMargin(string name, float left, float right, float top, float bottom)
        {
            XmlNode content = (XmlNode)root.child["content.xml"];

            XmlDocument doc = content.doc;

            XmlNodeList list = doc.GetElementsByTagName("style", header_style);

            foreach (XmlElement e in list)
            {
                string check_name = e.GetAttribute("name", header_style);

                if (check_name.Equals(name))
                {
                    XmlElement e1 = (XmlElement)e.GetElementsByTagName("paragraph-properties", header_style).Item(0);
                    if (e1 == null)
                    {
                        e1 = doc.CreateElement("style:paragraph-properties", header_style);
                    }
                    e1.SetAttribute("margin-left", header_fo, left+"cm");
                    e1.SetAttribute("margin-right", header_fo, right + "cm");
                    e1.SetAttribute("margin-top", header_fo, top + "cm");
                    e1.SetAttribute("margin-bottom", header_fo, bottom + "cm");

                    e.PrependChild(e1);
                    break;
                }
            }
        }

        public void SetPIndent(string name, float ident)
        {
            XmlNode content = (XmlNode)root.child["content.xml"];

            XmlDocument doc = content.doc;

            XmlNodeList list = doc.GetElementsByTagName("style", header_style);

            foreach (XmlElement e in list)
            {
                string check_name = e.GetAttribute("name", header_style);

                if (check_name.Equals(name))
                {
                    XmlElement e1 = (XmlElement)e.GetElementsByTagName("paragraph-properties", header_style).Item(0);
                    if (e1 == null)
                    {
                        e1 = doc.CreateElement("style:paragraph-properties", header_style);
                    }
                    e1.SetAttribute("text-indent", header_fo, ident + "cm");

                    e.PrependChild(e1);
                    break;
                }
            }
        }

        public void SetPAlign(string name, string align)
        {
            XmlNode content = (XmlNode)root.child["content.xml"];

            XmlDocument doc = content.doc;

            XmlNodeList list = doc.GetElementsByTagName("style", header_style);

            foreach (XmlElement e in list)
            {
                string check_name = e.GetAttribute("name", header_style);

                if (check_name.Equals(name))
                {
                    XmlElement e1 = (XmlElement)e.GetElementsByTagName("paragraph-properties", header_style).Item(0);
                    if (e1 == null)
                    {
                        e1 = doc.CreateElement("style:paragraph-properties", header_style);
                    }
                    e1.SetAttribute("text-align", header_fo, align);
                    e1.SetAttribute("justify-single-word", header_style, "false");


                    e.PrependChild(e1);
                    break;
                }
            }
        }
        
        public string AddContentP(string text)
        {
            string pname = "P" + numP.ToString();
            XmlNode content = (XmlNode)root.child["content.xml"];
            XmlDocument doc = content.doc;

            XmlNodeList list = doc.GetElementsByTagName("automatic-styles", header_office);
            XmlElement e = (XmlElement)list.Item(0);
            
            XmlElement e1 = doc.CreateElement("style:style", header_style);

            e1.SetAttribute("name", header_style, pname);
            e1.SetAttribute("family", header_style, "paragraph");
            e1.SetAttribute("parent-style-name", header_style, "Standard");

            e.AppendChild(e1);
            
            list = doc.GetElementsByTagName("text", header_office);

            e = (XmlElement)list.Item(0);
            
            XmlElement text_element = doc.CreateElement("text:p", header_text);

            text_element.SetAttribute("style-name", header_text, pname);
            text_element.InnerText = text;
            e.AppendChild(text_element);

            numP++;

            return pname;
        }

        public void AddContentP(int p_number, string text)
        {
            XmlNode content = (XmlNode)root.child["content.xml"];
            XmlDocument doc = content.doc;

            XmlNodeList list = doc.GetElementsByTagName("p", header_text);

            XmlText xmltext = doc.CreateTextNode(text);
            list.Item(p_number).AppendChild(xmltext);

        }

        public string AddContentSpan(string pname, string text)
        {
            string spanname = "T" + numSpan.ToString();
            XmlNode content = (XmlNode)root.child["content.xml"];
            XmlDocument doc = content.doc;

            XmlNodeList list = doc.GetElementsByTagName("automatic-styles", header_office);
            XmlElement e = (XmlElement)list.Item(0);

            XmlElement e1 = doc.CreateElement("style:style", header_style);

            e1.SetAttribute("name", header_style, spanname);
            e1.SetAttribute("family", header_style, "text");

            e.AppendChild(e1);

            list = doc.GetElementsByTagName("p", header_text);

            foreach(XmlElement element in list)
            {
                string name_check = element.GetAttribute("style-name", header_text);
                if (name_check.Equals(pname))
                {
                    XmlElement text_element = doc.CreateElement("text:span", header_text);

                    text_element.SetAttribute("style-name", header_text, spanname);
                    text_element.InnerText = text;
                    element.AppendChild(text_element);
                    numSpan++;
                    break;
                }
            }
            return spanname;
        }

        public void SetPageLayout(float left, float right, float top, float bottom)
        {
            XmlNode styles = (XmlNode)root.child["styles.xml"];

            XmlDocument doc = styles.doc;

            XmlElement pagelayout = (XmlElement)doc.GetElementsByTagName("page-layout-properties", header_style).Item(0);
            pagelayout.SetAttribute("margin-left", header_fo, left + "cm");
            pagelayout.SetAttribute("margin-right", header_fo, right +"cm");
            pagelayout.SetAttribute("margin-top", header_fo, top + "cm");
            pagelayout.SetAttribute("margin-bottom", header_fo, bottom + "cm");   
        }

        public void SetStandardMargin(float left, float right, float top, float bottom)
        {
            XmlNode styles = (XmlNode)root.child["styles.xml"];

            XmlDocument doc = styles.doc;

            XmlNodeList list = doc.GetElementsByTagName("style", header_style);
            foreach(XmlElement e in list)
            {
                string stylename = e.GetAttribute("name", header_style);
                if (stylename.Equals("Standard"))
                {
                    XmlElement standardmargin = (XmlElement)e.GetElementsByTagName("paragraph-properties", header_style).Item(0);

                    standardmargin.SetAttribute("margin-left", header_fo, left + "cm");
                    standardmargin.SetAttribute("margin-right", header_fo, right + "cm");
                    standardmargin.SetAttribute("margin-top", header_fo, top + "cm");
                    standardmargin.SetAttribute("margin-bottom", header_fo, bottom + "cm");
                }
            }
        }

        public void Update_Text(string text)
        {
            //update content
            XmlNodeList con_list = content.GetElementsByTagName("p", content_text);

            foreach (XmlElement e in con_list)
            {
                e.InnerText = text;
            }

            //update meta
            int paragraph_count = text.Split("\n").Length;
            int word_count = text.Split(" ").Length;
            int character_count = text.Length;

            XmlNodeList meta_list = meta.GetElementsByTagName("document-statistic", meta_meta);

            foreach (XmlElement e in meta_list)
            {
                XmlAttribute xa = e.GetAttributeNode("paragraph-count", meta_meta);
                xa.Value = paragraph_count.ToString();

                xa = e.GetAttributeNode("word-count", meta_meta);
                xa.Value = word_count.ToString();

                xa = e.GetAttributeNode("character-count", meta_meta);
                xa.Value = character_count.ToString();
            }

        }

        static string content_text = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
        static string meta_meta = "urn:oasis:names:tc:opendocument:xmlns:meta:1.0";
        private string mimetype;

        private XmlDocument content;
        private XmlDocument manifest;
        private XmlDocument meta;
        private XmlDocument settings;
        private XmlDocument styles;

        private XmlDocument manifest_META;

        private XmlDocument current;

        /*
        public void Save_Document(string path)
        {
            XmlWriterSettings wrs = new XmlWriterSettings();
            wrs.Indent = false;
            wrs.NewLineChars = string.Empty;

            XmlWriter wr = XmlWriter.Create(path + @"\content.xml", wrs);

            content.Save(wr);
            manifest.Save(path + @"\manifest.rdf");
            meta.Save(path + @"\meta.xml");

            StreamWriter sw = File.CreateText(path + @"\mimetype");
            sw.Write(mimetype);
            sw.Close();

            settings.Save(path + @"\settings.xml");
            styles.Save(path + @"\styles.xml");

            manifest_META.Save(path + @"\META-INF" + @"\manifest.xml");
            current.Save(path + @"\Configurations2" + @"\accelerator" + @"\current.xml");
        }
        */

        /*
        public void Create_Document(string path)
        {
            Create_content(path);
            Create_manifest(path);
            Create_meta(path);
            Create_settings(path);
            Create_styles(path);
            Create_manifest_META(path);
            Create_current(path);
            Create_mimetype();
        }

        public void Create_content(string path)
        {
            path += @"\content.xml";
            content = new XmlDocument();
            content.Load(path);
        }

        public void Create_manifest(string path)
        {
            path += @"\manifest.xml";
            manifest = new XmlDocument();
            manifest.Load(path);
        }

        public void Create_meta(string path)
        {
            path += @"\meta.xml";
            meta = new XmlDocument();
            meta.Load(path);
        }

        public void Create_settings(string path)
        {
            path += @"\settings.xml";
            settings = new XmlDocument();
            settings.Load(path);
        }

        public void Create_styles(string path)
        {
            path += @"\styles.xml";
            styles = new XmlDocument();
            styles.Load(path);
        }

        public void Create_manifest_META(string path)
        {
            path += @"\META-INF" + @"\manifest.xml";
            manifest_META = new XmlDocument();
            manifest_META.Load(path);

        }

        public void Create_current(string path)
        {
            path += @"\Configurations2" + @"\accelerator" + @"\current.xml";
            current = new XmlDocument();

            try
            {
                current.Load(path);
            }
            catch(Exception e)
            {
                XmlElement root = current.CreateElement("root");
                current.AppendChild(root);
            }

        }


        public void Create_mimetype()
        {
           mimetype = "application/vnd.oasis.opendocument.text";
        }

        */

        /*
        public void Create_content(string path)
        {
            XmlDocument xml = new XmlDocument();

            //Layer1

            XmlElement root = xml.CreateElement("office:document-content", "urn:oasis:names:tc:opendocument:xmlns:office:1.0");

            
            root.SetAttribute("xmlns:office", "urn:oasis:names:tc:opendocument:xmlns:office:1.0");
            root.SetAttribute("xmlns:style", "urn:oasis:names:tc:opendocument:xmlns:style:1.0");
            root.SetAttribute("xmlns:text", "urn:oasis:names:tc:opendocument:xmlns:text:1.0");
            root.SetAttribute("xmlns:table", "urn:oasis:names:tc:opendocument:xmlns:table:1.0");
            root.SetAttribute("xmlns:draw", "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0");
            root.SetAttribute("xmlns:fo", "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0");
            root.SetAttribute("xmlns:xlink", "http://www.w3.org/1999/xlink");
            root.SetAttribute("xmlns:dc", "http://purl.org/dc/elements/1.1/");
            root.SetAttribute("xmlns:meta", "urn:oasis:names:tc:opendocument:xmlns:meta:1.0");
            root.SetAttribute("xmlns:number", "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0");
            root.SetAttribute("xmlns:svg", "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0");
            root.SetAttribute("xmlns:chart", "urn:oasis:names:tc:opendocument:xmlns:chart:1.0");
            root.SetAttribute("xmlns:dr3d", "urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0");
            root.SetAttribute("xmlns:math", "http://www.w3.org/1998/Math/MathML");
            root.SetAttribute("xmlns:form", "urn:oasis:names:tc:opendocument:xmlns:form:1.0");
            root.SetAttribute("xmlns:script", "urn:oasis:names:tc:opendocument:xmlns:script:1.0");
            root.SetAttribute("xmlns:ooo", "http://openoffice.org/2004/office");
            root.SetAttribute("xmlns:ooow", "http://openoffice.org/2004/writer");
            root.SetAttribute("xmlns:oooc", "http://openoffice.org/2004/calc");
            root.SetAttribute("xmlns:dom", "http://www.w3.org/2001/xml-events");
            root.SetAttribute("xmlns:xforms", "http://www.w3.org/2002/xforms");
            root.SetAttribute("xmlns:xsd", "http://www.w3.org/2001/XMLSchema");
            root.SetAttribute("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance");
            root.SetAttribute("xmlns:rpt", "http://openoffice.org/2005/report");
            root.SetAttribute("xmlns:of", "urn:oasis:names:tc:opendocument:xmlns:of:1.2");
            root.SetAttribute("xmlns:xhtml", "http://www.w3.org/1999/xhtml");
            root.SetAttribute("xmlns:grddl", "http://www.w3.org/2003/g/data-view#");
            root.SetAttribute("xmlns:tableooo", "http://openoffice.org/2009/table");
            root.SetAttribute("xmlns:textooo", "http://openoffice.org/2013/office");
            root.SetAttribute("xmlns:field", "urn:openoffice:names:experimental:ooo-ms-interop:xmlns:field:1.0");

            XmlAttribute xa = xml.CreateAttribute("office:version", "urn:oasis:names:tc:opendocument:xmlns:office:1.0");
            xa.Value = "1.2";
            root.SetAttributeNode(xa);


            // Layer2

            XmlElement e1 = xml.CreateElement("office:scripts", "urn:oasis:names:tc:opendocument:xmlns:office:1.0");
            root.AppendChild(e1);

            XmlElement e2 = xml.CreateElement("office:font-face-decls", "urn:oasis:names:tc:opendocument:xmlns:office:1.0");
            root.AppendChild(e2);

            //

            XmlElement e21 = xml.CreateElement("style:font-face", "urn:oasis:names:tc:opendocument:xmlns:style:1.0");
            e21.SetAttribute("name", "urn:oasis:names:tc:opendocument:xmlns:style:1.0", "Lucida Sans1");
            e21.SetAttribute("font-family", "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0", "'Lucida Sans'");
            e21.SetAttribute("font-family-generic", "urn:oasis:names:tc:opendocument:xmlns:style:1.0", "swiss");
            e2.AppendChild(e21);

            XmlElement e22 = xml.CreateElement("style:font-face", "urn:oasis:names:tc:opendocument:xmlns:style:1.0");
            e22.SetAttribute("name", "urn:oasis:names:tc:opendocument:xmlns:style:1.0", "바탕");
            e22.SetAttribute("font-family", "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0", "바탕");
            e22.SetAttribute("font-family-generic", "urn:oasis:names:tc:opendocument:xmlns:style:1.0", "roman");
            e22.SetAttribute("font-pitch", "urn:oasis:names:tc:opendocument:xmlns:style:1.0", "variable");
            e2.AppendChild(e22);

            XmlElement e23 = xml.CreateElement("style:font-face", "urn:oasis:names:tc:opendocument:xmlns:style:1.0");
            e23.SetAttribute("name", "urn:oasis:names:tc:opendocument:xmlns:style:1.0", "Arial");
            e23.SetAttribute("font-family", "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0", "Arial");
            e23.SetAttribute("font-family-generic", "urn:oasis:names:tc:opendocument:xmlns:style:1.0", "swiss");
            e23.SetAttribute("font-pitch", "urn:oasis:names:tc:opendocument:xmlns:style:1.0", "variable");
            e2.AppendChild(e23);

            XmlElement e24 = xml.CreateElement("style:font-face", "urn:oasis:names:tc:opendocument:xmlns:style:1.0");
            e24.SetAttribute("name", "urn:oasis:names:tc:opendocument:xmlns:style:1.0", "Lucida Sans");
            e24.SetAttribute("font-family", "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0", "'Lucida Sans'");
            e24.SetAttribute("font-family-generic", "urn:oasis:names:tc:opendocument:xmlns:style:1.0", "system");
            e24.SetAttribute("font-pitch", "urn:oasis:names:tc:opendocument:xmlns:style:1.0", "variable");
            e2.AppendChild(e24);

            XmlElement e25 = xml.CreateElement("style:font-face", "urn:oasis:names:tc:opendocument:xmlns:style:1.0");
            e25.SetAttribute("name", "urn:oasis:names:tc:opendocument:xmlns:style:1.0", "바탕1");
            e25.SetAttribute("font-family", "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0", "바탕");
            e25.SetAttribute("font-family-generic", "urn:oasis:names:tc:opendocument:xmlns:style:1.0", "system");
            e25.SetAttribute("font-pitch", "urn:oasis:names:tc:opendocument:xmlns:style:1.0", "variable");
            e2.AppendChild(e25);


            XmlElement e3 = xml.CreateElement("office:automatic-styles", "urn:oasis:names:tc:opendocument:xmlns:office:1.0");
            root.AppendChild(e3);

            XmlElement e4 = xml.CreateElement("office:body", "urn:oasis:names:tc:opendocument:xmlns:office:1.0");
            root.AppendChild(e4);

            XmlElement e41 = xml.CreateElement("office:text", "urn:oasis:names:tc:opendocument:xmlns:office:1.0");
            e4.AppendChild(e41);

            //

            XmlElement e411 = xml.CreateElement("text:sequence-decls", "urn:oasis:names:tc:opendocument:xmlns:text:1.0");
            e41.AppendChild(e411);

            //

            XmlElement e4111 = xml.CreateElement("text:sequence-decl", "urn:oasis:names:tc:opendocument:xmlns:text:1.0");
            e4111.SetAttribute("display-outline-level", "urn:oasis:names:tc:opendocument:xmlns:text:1.0", "0");
            e4111.SetAttribute("name", "urn:oasis:names:tc:opendocument:xmlns:text:1.0", "Illustration");
            e411.AppendChild(e4111);
            XmlElement e4112 = xml.CreateElement("text:sequence-decl", "urn:oasis:names:tc:opendocument:xmlns:text:1.0");
            e4112.SetAttribute("display-outline-level", "urn:oasis:names:tc:opendocument:xmlns:text:1.0", "0");
            e4112.SetAttribute("name", "urn:oasis:names:tc:opendocument:xmlns:text:1.0", "Table");
            e411.AppendChild(e4112);
            XmlElement e4113 = xml.CreateElement("text:sequence-decl", "urn:oasis:names:tc:opendocument:xmlns:text:1.0");
            e4113.SetAttribute("display-outline-level", "urn:oasis:names:tc:opendocument:xmlns:text:1.0", "0");
            e4113.SetAttribute("name", "urn:oasis:names:tc:opendocument:xmlns:text:1.0", "Text");
            e411.AppendChild(e4113);
            XmlElement e4114 = xml.CreateElement("text:sequence-decl", "urn:oasis:names:tc:opendocument:xmlns:text:1.0");
            e4114.SetAttribute("display-outline-level", "urn:oasis:names:tc:opendocument:xmlns:text:1.0", "0");
            e4114.SetAttribute("name", "urn:oasis:names:tc:opendocument:xmlns:text:1.0", "Drawing");
            e411.AppendChild(e4114);


            XmlElement e412 = xml.CreateElement("text:p", "urn:oasis:names:tc:opendocument:xmlns:text:1.0");
            e412.SetAttribute("style-name", "urn:oasis:names:tc:opendocument:xmlns:text:1.0", "Standard");
            e41.AppendChild(e412);

            xml.AppendChild(root);

            xml.InsertBefore(xml.CreateXmlDeclaration("1.0", "UTF-8", null), root);

            xml.Save(path);
        }

        public void Create_manifest(string path)
        {
            XmlDocument xml = new XmlDocument();

            XmlElement root = xml.CreateElement("rdf:RDF", "http://www.w3.org/1999/02/22-rdf-syntax-ns#");

            XmlElement e1 = xml.CreateElement("rdf:Description", "http://www.w3.org/1999/02/22-rdf-syntax-ns#");
            e1.SetAttribute("about", "http://www.w3.org/1999/02/22-rdf-syntax-ns#", "styles.xml");
            root.AppendChild(e1);

            XmlElement e11 = xml.CreateElement("rdf:type", "http://www.w3.org/1999/02/22-rdf-syntax-ns#");
            e11.SetAttribute("resource", "http://www.w3.org/1999/02/22-rdf-syntax-ns#", "http://docs.oasis-open.org/ns/office/1.2/meta/odf#StylesFile");
            e1.AppendChild(e11);

            XmlElement e2 = xml.CreateElement("rdf:Description", "http://www.w3.org/1999/02/22-rdf-syntax-ns#");
            e2.SetAttribute("about", "http://www.w3.org/1999/02/22-rdf-syntax-ns#", "");
            root.AppendChild(e2);

            XmlElement e21 = xml.CreateElement("ns0:hasPart", "http://docs.oasis-open.org/ns/office/1.2/meta/pkg#");
            e21.SetAttribute("xmlns:ns0", "http://docs.oasis-open.org/ns/office/1.2/meta/pkg#");
            e21.SetAttribute("resource", "http://www.w3.org/1999/02/22-rdf-syntax-ns#", "styles.xml");
            e2.AppendChild(e21);

            XmlElement e3 = xml.CreateElement("rdf:Description", "http://www.w3.org/1999/02/22-rdf-syntax-ns#");
            e3.SetAttribute("about", "http://www.w3.org/1999/02/22-rdf-syntax-ns#", "content.xml");
            root.AppendChild(e3);

            XmlElement e31 = xml.CreateElement("rdf:type", "http://www.w3.org/1999/02/22-rdf-syntax-ns#");
            e31.SetAttribute("resource", "http://www.w3.org/1999/02/22-rdf-syntax-ns#", "http://docs.oasis-open.org/ns/office/1.2/meta/odf#ContentFile");
            e3.AppendChild(e31);

            XmlElement e4 = xml.CreateElement("rdf:Description", "http://www.w3.org/1999/02/22-rdf-syntax-ns#");
            e4.SetAttribute("about", "http://www.w3.org/1999/02/22-rdf-syntax-ns#", "");
            root.AppendChild(e4);

            XmlElement e41 = xml.CreateElement("ns0:hasPart", "http://docs.oasis-open.org/ns/office/1.2/meta/pkg#");
            e41.SetAttribute("xmlns:ns0", "http://docs.oasis-open.org/ns/office/1.2/meta/pkg#");
            e41.SetAttribute("resource", "http://www.w3.org/1999/02/22-rdf-syntax-ns#", "content.xml");
            e4.AppendChild(e41);

            XmlElement e5 = xml.CreateElement("rdf:Description", "http://www.w3.org/1999/02/22-rdf-syntax-ns#");
            e5.SetAttribute("about", "http://www.w3.org/1999/02/22-rdf-syntax-ns#", "");
            root.AppendChild(e5);

            XmlElement e51 = xml.CreateElement("rdf:type", "http://www.w3.org/1999/02/22-rdf-syntax-ns#");
            e51.SetAttribute("resource", "http://www.w3.org/1999/02/22-rdf-syntax-ns#", "http://docs.oasis-open.org/ns/office/1.2/meta/pkg#Document");
            e5.AppendChild(e51);

            xml.AppendChild(root);

            xml.InsertBefore(xml.CreateXmlDeclaration("1.0", "utf-8", null), root);

            xml.Save(path);
        }

        public void Create_meta(string path)
        {
            XmlDocument xml = new XmlDocument();

            XmlElement root = xml.CreateElement("office:document-meta", "urn:oasis:names:tc:opendocument:xmlns:office:1.0");

            root.SetAttribute("xmlns:office", "urn:oasis:names:tc:opendocument:xmlns:office:1.0");
            root.SetAttribute("xmlns:xlink", "http://www.w3.org/1999/xlink");
            root.SetAttribute("xmlns:dc", "http://purl.org/dc/elements/1.1/");
            root.SetAttribute("xmlns:meta", "urn:oasis:names:tc:opendocument:xmlns:meta:1.0");
            root.SetAttribute("xmlns:ooo", "http://openoffice.org/2004/office");
            root.SetAttribute("xmlns:grddl", "http://www.w3.org/2003/g/data-view#");
            root.SetAttribute("xmlns:textooo", "http://openoffice.org/2013/office");
            root.SetAttribute("version", "urn:oasis:names:tc:opendocument:xmlns:office:1.0", "1.2");

            XmlElement e1 = xml.CreateElement("office:meta", "urn:oasis:names:tc:opendocument:xmlns:office:1.0");
            root.AppendChild(e1);

            XmlElement e11 = xml.CreateElement("meta:initial-creator", "urn:oasis:names:tc:opendocument:xmlns:meta:1.0");
            e11.InnerText = "yujin cha";
            e1.AppendChild(e11);

            XmlElement e12 = xml.CreateElement("meta:creation-date", "urn:oasis:names:tc:opendocument:xmlns:meta:1.0");
            e12.InnerText = "2021-05-29T03:32:13.76";
            e1.AppendChild(e12);

            XmlElement e13 = xml.CreateElement("meta:document-statistic", "urn:oasis:names:tc:opendocument:xmlns:meta:1.0");
            e13.SetAttribute("table-count", "urn:oasis:names:tc:opendocument:xmlns:meta:1.0", "0");
            e13.SetAttribute("image-count", "urn:oasis:names:tc:opendocument:xmlns:meta:1.0", "0");
            e13.SetAttribute("object-count", "urn:oasis:names:tc:opendocument:xmlns:meta:1.0", "0");
            e13.SetAttribute("page-count", "urn:oasis:names:tc:opendocument:xmlns:meta:1.0", "1");
            e13.SetAttribute("paragraph-count", "urn:oasis:names:tc:opendocument:xmlns:meta:1.0", "0");
            e13.SetAttribute("word-count", "urn:oasis:names:tc:opendocument:xmlns:meta:1.0", "0");
            e13.SetAttribute("character-count", "urn:oasis:names:tc:opendocument:xmlns:meta:1.0", "0");
            e1.AppendChild(e13);

            XmlElement e14 = xml.CreateElement("meta:generator", "urn:oasis:names:tc:opendocument:xmlns:meta:1.0");
            e14.InnerText = "OpenOffice/4.1.10$Win32 OpenOffice.org_project/4110m2$Build-9807";
            e1.AppendChild(e14);

            xml.AppendChild(root);

            xml.InsertBefore(xml.CreateXmlDeclaration("1.0", "UTF-8", null), root);

            xml.Save(path);
        }



        public void Create_settings(string path)
        {
            XmlDocument xml = new XmlDocument();

            XmlElement root = xml.CreateElement("office:document-settings", "urn:oasis:names:tc:opendocument:xmlns:office:1.0");

            root.SetAttribute("xmlns:office", "urn:oasis:names:tc:opendocument:xmlns:office:1.0");
            root.SetAttribute("xmlns:xlink", "http://www.w3.org/1999/xlink");
            root.SetAttribute("xmlns:config", "urn:oasis:names:tc:opendocument:xmlns:config:1.0");
            root.SetAttribute("xmlns:ooo", "http://openoffice.org/2004/office");
            root.SetAttribute("xmlns:textooo", "http://openoffice.org/2013/office");
            root.SetAttribute("version", "urn:oasis:names:tc:opendocument:xmlns:office:1.0", "1.2");

            XmlElement setting = xml.CreateElement("office:settings", "urn:oasis:names:tc:opendocument:xmlns:office:1.0");
            root.AppendChild(setting);

            XmlElement config1 = xml.CreateElement("config:config-item-set", "urn:oasis:names:tc:opendocument:xmlns:config:1.0");
            config1.SetAttribute("name", "urn:oasis:names:tc:opendocument:xmlns:config:1.0", "ooo:view-settings");
            setting.AppendChild(config1);


            XmlElement e1 = xml.CreateElement("config:config-item", "urn:oasis:names:tc:opendocument:xmlns:config:1.0");
            e1.SetAttribute("name", "urn:oasis:names:tc:opendocument:xmlns:config:1.0", "ViewAreaTop");
            e1.SetAttribute("type", "urn:oasis:names:tc:opendocument:xmlns:config:1.0", "int");
            e1.InnerText = "0";
            config1.AppendChild(e1);

            XmlElement e2 = xml.CreateElement("config:config-item", "urn:oasis:names:tc:opendocument:xmlns:config:1.0");
            e2.SetAttribute("name", "urn:oasis:names:tc:opendocument:xmlns:config:1.0", "ViewAreaLeft");
            e2.SetAttribute("type", "urn:oasis:names:tc:opendocument:xmlns:config:1.0", "int");
            e2.InnerText = "0";
            config1.AppendChild(e2);

            XmlElement e3 = xml.CreateElement("config:config-item", "urn:oasis:names:tc:opendocument:xmlns:config:1.0");
            e3.SetAttribute("name", "urn:oasis:names:tc:opendocument:xmlns:config:1.0", "ViewAreaWidth");
            e3.SetAttribute("type", "urn:oasis:names:tc:opendocument:xmlns:config:1.0", "int");
            e3.InnerText = "28259";
            config1.AppendChild(e3);

            XmlElement e4 = xml.CreateElement("config:config-item", "urn:oasis:names:tc:opendocument:xmlns:config:1.0");
            e4.SetAttribute("name", "urn:oasis:names:tc:opendocument:xmlns:config:1.0", "ViewAreaHeight");
            e4.SetAttribute("type", "urn:oasis:names:tc:opendocument:xmlns:config:1.0", "int");
            e4.InnerText = "14951";
            config1.AppendChild(e4);

            XmlElement e5 = xml.CreateElement("config:config-item", "urn:oasis:names:tc:opendocument:xmlns:config:1.0");
            e5.SetAttribute("name", "urn:oasis:names:tc:opendocument:xmlns:config:1.0", "ShowRedlineChanges");
            e5.SetAttribute("type", "urn:oasis:names:tc:opendocument:xmlns:config:1.0", "boolean");
            e5.InnerText = "true";
            config1.AppendChild(e5);

            XmlElement e6 = xml.CreateElement("config:config-item", "urn:oasis:names:tc:opendocument:xmlns:config:1.0");
            e6.SetAttribute("name", "urn:oasis:names:tc:opendocument:xmlns:config:1.0", "InBrowseMode");
            e6.SetAttribute("type", "urn:oasis:names:tc:opendocument:xmlns:config:1.0", "boolean");
            e6.InnerText = "false";
            config1.AppendChild(e6);

            XmlElement e7 = xml.CreateElement("config:config-item-map-indexed", "urn:oasis:names:tc:opendocument:xmlns:config:1.0");
            e7.SetAttribute("name", "urn:oasis:names:tc:opendocument:xmlns:config:1.0", "Views");
            config1.AppendChild(e7);


            XmlElement e71 = xml.CreateElement("config:config-item-map-entry", "urn:oasis:names:tc:opendocument:xmlns:config:1.0");
            e7.AppendChild(e71);

            string con = "urn:oasis:names:tc:opendocument:xmlns:config:1.0";
            XmlElement e711 = xml.CreateElement("config:config-item", con);
            e711.SetAttribute("name", con, "ViewId");
            e711.SetAttribute("type", con, "string");
            e711.InnerText = "view2";
            e71.AppendChild(e711);

            XmlElement e712 = xml.CreateElement("config:config-item", con);
            e712.SetAttribute("name", con, "ViewLeft");
            e712.SetAttribute("type", con, "int");
            e712.InnerText = "5629";
            e71.AppendChild(e712);

            XmlElement e713 = xml.CreateElement("config:config-item", con);
            e713.SetAttribute("name", con, "ViewTop");
            e713.SetAttribute("type", con, "int");
            e713.InnerText = "3002";
            e71.AppendChild(e713);

            XmlElement e714 = xml.CreateElement("config:config-item", con);
            e714.SetAttribute("name", con, "VisibleLeft");
            e714.SetAttribute("type", con, "int");
            e714.InnerText = "0";
            e71.AppendChild(e714);

            XmlElement e715 = xml.CreateElement("config:config-item", con);
            e715.SetAttribute("name", con, "VisibleTop");
            e715.SetAttribute("type", con, "int");
            e715.InnerText = "0";
            e71.AppendChild(e715);

            XmlElement e716 = xml.CreateElement("config:config-item", con);
            e716.SetAttribute("name", con, "VisibleRight");
            e716.SetAttribute("type", con, "int");
            e716.InnerText = "28258";
            e71.AppendChild(e716);

            XmlElement e717 = xml.CreateElement("config:config-item", con);
            e717.SetAttribute("name", con, "VisibleBottom");
            e717.SetAttribute("type", con, "int");
            e717.InnerText = "14949";
            e71.AppendChild(e717);

            XmlElement e718 = xml.CreateElement("config:config-item", con);
            e718.SetAttribute("name", con, "ZoomType");
            e718.SetAttribute("type", con, "short");
            e718.InnerText = "0";
            e71.AppendChild(e718);

            XmlElement e719 = xml.CreateElement("config:config-item", con);
            e719.SetAttribute("name", con, "ViewLayoutColumns");
            e719.SetAttribute("type", con, "short");
            e719.InnerText = "0";
            e71.AppendChild(e719);

            XmlElement e7110 = xml.CreateElement("config:config-item", con);
            e7110.SetAttribute("name", con, "ViewLayoutBookMode");
            e7110.SetAttribute("type", con, "boolean");
            e7110.InnerText = "false";
            e71.AppendChild(e7110);

            XmlElement e7111 = xml.CreateElement("config:config-item", con);
            e7111.SetAttribute("name", con, "ZoomFactor");
            e7111.SetAttribute("type", con, "short");
            e7111.InnerText = "100";
            e71.AppendChild(e7111);

            XmlElement e7112 = xml.CreateElement("config:config-item", con);
            e7112.SetAttribute("name", con, "IsSelectedFrame");
            e7112.SetAttribute("type", con, "boolean");
            e7112.InnerText = "false";
            e71.AppendChild(e7112);



            XmlElement config2 = xml.CreateElement("config:config-item-set", "urn:oasis:names:tc:opendocument:xmlns:config:1.0");
            config2.SetAttribute("name", "urn:oasis:names:tc:opendocument:xmlns:config:1.0", "ooo:configuration-settings");
            setting.AppendChild(config2);

            XmlElement c1 = xml.CreateElement("config:config-item", con);
            c1.SetAttribute("name", con, "CurrentDatabaseDataSource");
            c1.SetAttribute("type", con, "string");
            config2.AppendChild(c1);

            XmlElement c2 = xml.CreateElement("config:config-item", con);
            c2.SetAttribute("name", con, "ConsiderTextWrapOnObjPos");
            c2.SetAttribute("type", con, "boolean");
            c2.InnerText = "false";
            config2.AppendChild(c2);

            XmlElement c3 = xml.CreateElement("config:config-item", con);
            c3.SetAttribute("name", con, "AddParaTableSpacing");
            c3.SetAttribute("type", con, "boolean");
            c3.InnerText = "true";
            config2.AppendChild(c3);

            XmlElement c4 = xml.CreateElement("config:config-item", con);
            c4.SetAttribute("name", con, "PrintReversed");
            c4.SetAttribute("type", con, "boolean");
            c4.InnerText = "false";
            config2.AppendChild(c4);

            XmlElement c5 = xml.CreateElement("config:config-item", con);
            c5.SetAttribute("name", con, "PrintRightPages");
            c5.SetAttribute("type", con, "boolean");
            c5.InnerText = "true";
            config2.AppendChild(c5);

            XmlElement c6 = xml.CreateElement("config:config-item", con);
            c6.SetAttribute("name", con, "UseOldNumbering");
            c6.SetAttribute("type", con, "boolean");
            c6.InnerText = "false";
            config2.AppendChild(c6);

            XmlElement c7 = xml.CreateElement("config:config-item", con);
            c7.SetAttribute("name", con, "PrintProspectRTL");
            c7.SetAttribute("type", con, "boolean");
            c7.InnerText = "false";
            config2.AppendChild(c7);

            XmlElement c8 = xml.CreateElement("config:config-item", con);
            c8.SetAttribute("name", con, "PrintTables");
            c8.SetAttribute("type", con, "boolean");
            c8.InnerText = "true";
            config2.AppendChild(c8);

            XmlElement c9 = xml.CreateElement("config:config-item", con);
            c9.SetAttribute("name", con, "CurrentDatabaseCommandType");
            c9.SetAttribute("type", con, "int");
            c9.InnerText = "0";
            config2.AppendChild(c9);

            XmlElement c10 = xml.CreateElement("config:config-item", con);
            c10.SetAttribute("name", con, "DoNotJustifyLinesWithManualBreak");
            c10.SetAttribute("type", con, "boolean");
            c10.InnerText = "false";
            config2.AppendChild(c10);

            XmlElement c11 = xml.CreateElement("config:config-item", con);
            c11.SetAttribute("name", con, "AlignTabStopPosition");
            c11.SetAttribute("type", con, "boolean");
            c11.InnerText = "true";
            config2.AppendChild(c11);

            XmlElement c12 = xml.CreateElement("config:config-item", con);
            c12.SetAttribute("name", con, "PrinterSetup");
            c12.SetAttribute("type", con, "base64Binary");
            config2.AppendChild(c12);

            XmlElement c13 = xml.CreateElement("config:config-item", con);
            c13.SetAttribute("name", con, "CurrentDatabaseCommand");
            c13.SetAttribute("type", con, "string");
            config2.AppendChild(c13);

            XmlElement c14 = xml.CreateElement("config:config-item", con);
            c14.SetAttribute("name", con, "UseFormerTextWrapping");
            c14.SetAttribute("type", con, "boolean");
            c14.InnerText = "false";
            config2.AppendChild(c14);

            XmlElement c15 = xml.CreateElement("config:config-item", con);
            c15.SetAttribute("name", con, "TableRowKeep");
            c15.SetAttribute("type", con, "boolean");
            c15.InnerText = "false";
            config2.AppendChild(c15);

            XmlElement c16 = xml.CreateElement("config:config-item", con);
            c16.SetAttribute("name", con, "AddFrameOffsets");
            c16.SetAttribute("type", con, "boolean");
            c16.InnerText = "false";
            config2.AppendChild(c16);

            XmlElement c17 = xml.CreateElement("config:config-item", con);
            c17.SetAttribute("name", con, "PrintEmptyPages");
            c17.SetAttribute("type", con, "boolean");
            c17.InnerText = "true";
            config2.AppendChild(c17);

            XmlElement c18 = xml.CreateElement("config:config-item", con);
            c18.SetAttribute("name", con, "FieldAutoUpdate");
            c18.SetAttribute("type", con, "boolean");
            c18.InnerText = "true";
            config2.AppendChild(c18);

            XmlElement c19 = xml.CreateElement("config:config-item", con);
            c19.SetAttribute("name", con, "OutlineLevelYieldsNumbering");
            c19.SetAttribute("type", con, "boolean");
            c19.InnerText = "false";
            config2.AppendChild(c19);

            XmlElement c20 = xml.CreateElement("config:config-item", con);
            c20.SetAttribute("name", con, "PrintDrawings");
            c20.SetAttribute("type", con, "boolean");
            c20.InnerText = "true";
            config2.AppendChild(c20);

            XmlElement c21 = xml.CreateElement("config:config-item", con);
            c21.SetAttribute("name", con, "PrintTextPlaceholder");
            c21.SetAttribute("type", con, "boolean");
            c21.InnerText = "false";
            config2.AppendChild(c21);

            XmlElement c22 = xml.CreateElement("config:config-item", con);
            c22.SetAttribute("name", con, "LinkUpdateMode");
            c22.SetAttribute("type", con, "short");
            c22.InnerText = "1";
            config2.AppendChild(c22);

            XmlElement c23 = xml.CreateElement("config:config-item", con);
            c23.SetAttribute("name", con, "PrintPaperFromSetup");
            c23.SetAttribute("type", con, "boolean");
            c23.InnerText = "false";
            config2.AppendChild(c23);

            XmlElement c24 = xml.CreateElement("config:config-item", con);
            c24.SetAttribute("name", con, "PrintLeftPages");
            c24.SetAttribute("type", con, "boolean");
            c24.InnerText = "true";
            config2.AppendChild(c24);

            XmlElement c25 = xml.CreateElement("config:config-item", con);
            c25.SetAttribute("name", con, "AddParaTableSpacingAtStart");
            c25.SetAttribute("type", con, "boolean");
            c25.InnerText = "true";
            config2.AppendChild(c25);

            XmlElement c26 = xml.CreateElement("config:config-item", con);
            c26.SetAttribute("name", con, "DoNotResetParaAttrsForNumFont");
            c26.SetAttribute("type", con, "boolean");
            c26.InnerText = "false";
            config2.AppendChild(c26);

            XmlElement c27 = xml.CreateElement("config:config-item", con);
            c27.SetAttribute("name", con, "AllowPrintJobCancel");
            c27.SetAttribute("type", con, "boolean");
            c27.InnerText = "true";
            config2.AppendChild(c27);

            XmlElement c28 = xml.CreateElement("config:config-item", con);
            c28.SetAttribute("name", con, "IgnoreFirstLineIndentInNumbering");
            c28.SetAttribute("type", con, "boolean");
            c28.InnerText = "false";
            config2.AppendChild(c28);

            XmlElement c29 = xml.CreateElement("config:config-item", con);
            c29.SetAttribute("name", con, "ChartAutoUpdate");
            c29.SetAttribute("type", con, "boolean");
            c29.InnerText = "true";
            config2.AppendChild(c29);

            XmlElement c30 = xml.CreateElement("config:config-item", con);
            c30.SetAttribute("name", con, "TabAtLeftIndentForParagraphsInList");
            c30.SetAttribute("type", con, "boolean");
            c30.InnerText = "false";
            config2.AppendChild(c30);

            XmlElement c31 = xml.CreateElement("config:config-item", con);
            c31.SetAttribute("name", con, "PrintHiddenText");
            c31.SetAttribute("type", con, "boolean");
            c31.InnerText = "false";
            config2.AppendChild(c31);

            XmlElement c32 = xml.CreateElement("config:config-item", con);
            c32.SetAttribute("name", con, "LoadReadonly");
            c32.SetAttribute("type", con, "boolean");
            c32.InnerText = "false";
            config2.AppendChild(c32);

            XmlElement c33 = xml.CreateElement("config:config-item", con);
            c33.SetAttribute("name", con, "SaveGlobalDocumentLinks");
            c33.SetAttribute("type", con, "boolean");
            c33.InnerText = "false";
            config2.AppendChild(c33);

            XmlElement c34 = xml.CreateElement("config:config-item", con);
            c34.SetAttribute("name", con, "PrintAnnotationMode");
            c34.SetAttribute("type", con, "short");
            c34.InnerText = "0";
            config2.AppendChild(c34);

            XmlElement c35 = xml.CreateElement("config:config-item", con);
            c35.SetAttribute("name", con, "ApplyUserData");
            c35.SetAttribute("type", con, "boolean");
            c35.InnerText = "true";
            config2.AppendChild(c35);

            XmlElement c36 = xml.CreateElement("config:config-item", con);
            c36.SetAttribute("name", con, "UnxForceZeroExtLeading");
            c36.SetAttribute("type", con, "boolean");
            c36.InnerText = "false";
            config2.AppendChild(c36);

            XmlElement c37 = xml.CreateElement("config:config-item", con);
            c37.SetAttribute("name", con, "PrintBlackFonts");
            c37.SetAttribute("type", con, "boolean");
            c37.InnerText = "false";
            config2.AppendChild(c37);

            XmlElement c38 = xml.CreateElement("config:config-item", con);
            c38.SetAttribute("name", con, "RedlineProtectionKey");
            c38.SetAttribute("type", con, "base64Binary");
            config2.AppendChild(c38);

            XmlElement c39 = xml.CreateElement("config:config-item", con);
            c39.SetAttribute("name", con, "PrintProspect");
            c39.SetAttribute("type", con, "boolean");
            c39.InnerText = "false";
            config2.AppendChild(c39);

            XmlElement c40 = xml.CreateElement("config:config-item", con);
            c40.SetAttribute("name", con, "ProtectForm");
            c40.SetAttribute("type", con, "boolean");
            c40.InnerText = "false";
            config2.AppendChild(c40);

            XmlElement c41 = xml.CreateElement("config:config-item", con);
            c41.SetAttribute("name", con, "UpdateFromTemplate");
            c41.SetAttribute("type", con, "boolean");
            c41.InnerText = "true";
            config2.AppendChild(c41);

            XmlElement c42 = xml.CreateElement("config:config-item", con);
            c42.SetAttribute("name", con, "AddParaSpacingToTableCells");
            c42.SetAttribute("type", con, "boolean");
            c42.InnerText = "true";
            config2.AppendChild(c42);

            XmlElement c43 = xml.CreateElement("config:config-item", con);
            c43.SetAttribute("name", con, "TabsRelativeToIndent");
            c43.SetAttribute("type", con, "boolean");
            c43.InnerText = "rue";
            config2.AppendChild(c43);

            XmlElement c44 = xml.CreateElement("config:config-item", con);
            c44.SetAttribute("name", con, "IgnoreTabsAndBlanksForLineCalculation");
            c44.SetAttribute("type", con, "boolean");
            c44.InnerText = "false";
            config2.AppendChild(c44);

            XmlElement c45 = xml.CreateElement("config:config-item", con);
            c45.SetAttribute("name", con, "PrinterName");
            c45.SetAttribute("type", con, "string");
            config2.AppendChild(c45);

            XmlElement c46 = xml.CreateElement("config:config-item", con);
            c46.SetAttribute("name", con, "UseOldPrinterMetrics");
            c46.SetAttribute("type", con, "boolean");
            c46.InnerText = "false";
            config2.AppendChild(c46);

            XmlElement c47 = xml.CreateElement("config:config-item", con);
            c47.SetAttribute("name", con, "IsKernAsianPunctuation");
            c47.SetAttribute("type", con, "boolean");
            c47.InnerText = "false";
            config2.AppendChild(c47);

            XmlElement c48 = xml.CreateElement("config:config-item", con);
            c48.SetAttribute("name", con, "PrintPageBackground");
            c48.SetAttribute("type", con, "boolean");
            c48.InnerText = "true";
            config2.AppendChild(c48);

            XmlElement c49 = xml.CreateElement("config:config-item", con);
            c49.SetAttribute("name", con, "ClipAsCharacterAnchoredWriterFlyFrames");
            c49.SetAttribute("type", con, "boolean");
            c49.InnerText = "false";
            config2.AppendChild(c49);

            XmlElement c50 = xml.CreateElement("config:config-item", con);
            c50.SetAttribute("name", con, "IsLabelDocument");
            c50.SetAttribute("type", con, "boolean");
            c50.InnerText = "false";
            config2.AppendChild(c50);

            XmlElement c51 = xml.CreateElement("config:config-item", con);
            c51.SetAttribute("name", con, "PrintGraphics");
            c51.SetAttribute("type", con, "boolean");
            c51.InnerText = "true";
            config2.AppendChild(c51);

            XmlElement c52 = xml.CreateElement("config:config-item", con);
            c52.SetAttribute("name", con, "PrintSingleJobs");
            c52.SetAttribute("type", con, "boolean");
            c52.InnerText = "false";
            config2.AppendChild(c52);

            XmlElement c54 = xml.CreateElement("config:config-item", con);
            c54.SetAttribute("name", con, "DoNotCaptureDrawObjsOnPage");
            c54.SetAttribute("type", con, "boolean");
            c54.InnerText = "false";
            config2.AppendChild(c54);

            XmlElement c55 = xml.CreateElement("config:config-item", con);
            c55.SetAttribute("name", con, "PrinterIndependentLayout");
            c55.SetAttribute("type", con, "string");
            c55.InnerText = "high-resolution";
            config2.AppendChild(c55);

            XmlElement c56 = xml.CreateElement("config:config-item", con);
            c56.SetAttribute("name", con, "UseFormerObjectPositioning");
            c56.SetAttribute("type", con, "boolean");
            c56.InnerText = "false";
            config2.AppendChild(c56);

            XmlElement c57 = xml.CreateElement("config:config-item", con);
            c57.SetAttribute("name", con, "PrintFaxName");
            c57.SetAttribute("type", con, "string");
            config2.AppendChild(c57);

            XmlElement c58 = xml.CreateElement("config:config-item", con);
            c58.SetAttribute("name", con, "CharacterCompressionType");
            c58.SetAttribute("type", con, "short");
            c58.InnerText = "0";
            config2.AppendChild(c58);

            XmlElement c59 = xml.CreateElement("config:config-item", con);
            c59.SetAttribute("name", con, "AddExternalLeading");
            c59.SetAttribute("type", con, "boolean");
            c59.InnerText = "true";
            config2.AppendChild(c59);

            XmlElement c60 = xml.CreateElement("config:config-item", con);
            c60.SetAttribute("name", con, "MathBaselineAlignment");
            c60.SetAttribute("type", con, "boolean");
            c60.InnerText = "true";
            config2.AppendChild(c60);

            XmlElement c61 = xml.CreateElement("config:config-item", con);
            c61.SetAttribute("name", con, "UseFormerLineSpacing");
            c61.SetAttribute("type", con, "boolean");
            c61.InnerText = "false";
            config2.AppendChild(c61);

            XmlElement c62 = xml.CreateElement("config:config-item", con);
            c62.SetAttribute("name", con, "PrintControls");
            c62.SetAttribute("type", con, "boolean");
            c62.InnerText = "true";
            config2.AppendChild(c62);

            XmlElement c63 = xml.CreateElement("config:config-item", con);
            c63.SetAttribute("name", con, "SaveVersionOnClose");
            c63.SetAttribute("type", con, "boolean");
            c63.InnerText = "false";
            config2.AppendChild(c63);


            xml.AppendChild(root);

            xml.InsertBefore(xml.CreateXmlDeclaration("1.0", "UTF-8", null), root);

            xml.Save(path);
        }

        public void Create_styles(string path)
        {
            XmlDocument xml = new XmlDocument();

            string office = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
            string style = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
            string text = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
            string table = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
            string draw = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
            string fo = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";
            string xlink = "http://www.w3.org/1999/xlink";
            string dc = "http://purl.org/dc/elements/1.1/";
            string meta = "urn:oasis:names:tc:opendocument:xmlns:meta:1.0";
            string number = "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0";
            string svg = "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0";
            string chart = "urn:oasis:names:tc:opendocument:xmlns:chart:1.0";
            string dr3d = "urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0";
            string math = "http://www.w3.org/1998/Math/MathML";
            string form = "urn:oasis:names:tc:opendocument:xmlns:form:1.0";
            string script = "urn:oasis:names:tc:opendocument:xmlns:script:1.0";
            string ooo = "http://openoffice.org/2004/office";
            string ooow = "http://openoffice.org/2004/writer";
            string oooc = "http://openoffice.org/2004/calc";
            string dom = "http://www.w3.org/2001/xml-events";
            string rpt = "http://openoffice.org/2005/report";
            string of = "urn:oasis:names:tc:opendocument:xmlns:of:1.2";
            string xhtml = "http://www.w3.org/1999/xhtml";
            string grddl = "http://www.w3.org/2003/g/data-view#";
            string tableooo = "http://openoffice.org/2009/table";
            string textooo = "http://openoffice.org/2013/office";



            XmlElement root = xml.CreateElement("office:document-styles", office);

            root.SetAttribute("xmlns:office", office);
            root.SetAttribute("xmlns:style", style);
            root.SetAttribute("xmlns:text", text);
            root.SetAttribute("xmlns:table", table);
            root.SetAttribute("xmlns:draw", draw);
            root.SetAttribute("xmlns:fo", fo);
            root.SetAttribute("xmlns:xlink", xlink);
            root.SetAttribute("xmlns:dc", dc);
            root.SetAttribute("xmlns:meta", meta);
            root.SetAttribute("xmlns:number", number);
            root.SetAttribute("xmlns:svg", svg);
            root.SetAttribute("xmlns:chart", chart);
            root.SetAttribute("xmlns:dr3d", dr3d);
            root.SetAttribute("xmlns:math", math);
            root.SetAttribute("xmlns:form", form);
            root.SetAttribute("xmlns:script", script);
            root.SetAttribute("xmlns:ooo", ooo);
            root.SetAttribute("xmlns:ooow", ooow);
            root.SetAttribute("xmlns:oooc", oooc);
            root.SetAttribute("xmlns:dom", dom);
            root.SetAttribute("xmlns:rpt", rpt);
            root.SetAttribute("xmlns:of", of);
            root.SetAttribute("xmlns:xhtml", xhtml);
            root.SetAttribute("xmlns:grddl", grddl);
            root.SetAttribute("xmlns:tableooo", tableooo);
            root.SetAttribute("xmlns:textooo", textooo);

            XmlElement e_font = xml.CreateElement("office:font-face-decls", office);
            root.AppendChild(e_font);

            XmlElement a1 = xml.CreateElement("style:font-face", style);
            a1.SetAttribute("name", style, "Lucida Sans1");
            a1.SetAttribute("font-family", svg, "'Lucida Sans'");
            a1.SetAttribute("font-family-generic", style, "swiss");
            e_font.AppendChild(a1);

            XmlElement a2 = xml.CreateElement("style:font-face", style);
            a2.SetAttribute("name", style, "바탕");
            a2.SetAttribute("font-family", svg, "바탕");
            a2.SetAttribute("font-family-generic", style, "roman");
            a2.SetAttribute("font-pitch", style, "variable");
            e_font.AppendChild(a2);

            XmlElement a3 = xml.CreateElement("style:font-face", style);
            a3.SetAttribute("name", style, "Arial");
            a3.SetAttribute("font-family", svg, "Arial");
            a3.SetAttribute("font-family-generic", style, "swiss");
            a3.SetAttribute("font-pitch", style, "variable");
            e_font.AppendChild(a3);

            XmlElement a4 = xml.CreateElement("style:font-face", style);
            a4.SetAttribute("name", style, "Lucida Sans");
            a4.SetAttribute("font-family", svg, "'Lucida Sans'");
            a4.SetAttribute("font-family-generic", style, "system");
            a4.SetAttribute("font-pitch", style, "variable");
            e_font.AppendChild(a4);

            XmlElement a5 = xml.CreateElement("style:font-face", style);
            a5.SetAttribute("name", style, "바탕1");
            a5.SetAttribute("font-family", svg, "바탕");
            a5.SetAttribute("font-family-generic", style, "system");
            a5.SetAttribute("font-pitch", style, "variable");
            e_font.AppendChild(a5);



            XmlElement e_style = xml.CreateElement("office:styles", office);
            root.AppendChild(e_style);


            XmlElement b1 = xml.CreateElement("style:default-style", style);
            b1.SetAttribute("family", style, "graphic");
            e_style.AppendChild(b1);

            XmlElement b2 = xml.CreateElement("style:default-style", style);
            b2.SetAttribute("family", style, "paragraph");
            e_style.AppendChild(b2);

            XmlElement b3 = xml.CreateElement("style:default-style", style);
            b3.SetAttribute("family", style, "table");
            e_style.AppendChild(b3);

            XmlElement b4 = xml.CreateElement("style:default-style", style);
            b4.SetAttribute("family", style, "table-row");
            e_style.AppendChild(b4);

            XmlElement b5 = xml.CreateElement("style:default-style", style);
            b5.SetAttribute("family", style, "paragraph");
            e_style.AppendChild(b5);

            XmlElement b6 = xml.CreateElement("style:default-style", style);
            b6.SetAttribute("family", style, "graphic");
            e_style.AppendChild(b6);

            XmlElement b7 = xml.CreateElement("style:default-style", style);
            b7.SetAttribute("family", style, "graphic");
            e_style.AppendChild(b7);

            XmlElement b8 = xml.CreateElement("style:default-style", style);
            b8.SetAttribute("family", style, "graphic");
            e_style.AppendChild(b8);

            XmlElement b9 = xml.CreateElement("style:default-style", style);
            b9.SetAttribute("family", style, "graphic");
            e_style.AppendChild(b9);

            XmlElement b10 = xml.CreateElement("style:default-style", style);
            b10.SetAttribute("family", style, "graphic");
            e_style.AppendChild(b10);

            XmlElement b11 = xml.CreateElement("style:default-style", style);
            b11.SetAttribute("family", style, "graphic");
            e_style.AppendChild(b11);

            XmlElement b12 = xml.CreateElement("style:default-style", style);
            b12.SetAttribute("family", style, "graphic");
            e_style.AppendChild(b12);

            XmlElement b13 = xml.CreateElement("style:default-style", style);
            b13.SetAttribute("family", style, "graphic");
            e_style.AppendChild(b13);

            XmlElement b14 = xml.CreateElement("style:default-style", style);
            b14.SetAttribute("family", style, "graphic");
            e_style.AppendChild(b14);



            XmlElement e_automatic = xml.CreateElement("office:automatic-styles", office);
            root.AppendChild(e_automatic);

            XmlElement e_master = xml.CreateElement("office:master-styles", office);
            root.AppendChild(e_master);

            xml.AppendChild(root);

            xml.InsertBefore(xml.CreateXmlDeclaration("1.0", "UTF-8", null), root);

            xml.Save(path);
        }
        */
    }
}
