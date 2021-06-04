using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace H2O__
{
    public class ParagraphManager
    {
        public ParagraphManager()
        {

        }

        /*
        public void SetBold(string name)
        {
            XmlNode content = (XmlNode)root.child["content.xml"];

            XmlDocument doc = content.doc;

            XmlNodeList list = doc.GetElementsByTagName("style", header_style);

            foreach (XmlElement e in list)
            {
                string check_name = e.GetAttribute("name", header_style);

                if (check_name.Equals(name))
                {
                    XmlElement e1 = null;
                    string type = e.GetAttribute("family", header_style);

                    e1 = (XmlElement)e.GetElementsByTagName(type + "-properties", header_style).Item(0);

                    if (e1 == null)
                    {
                        e1 = doc.CreateElement("style:" + type + "-properties", header_style);
                    }
                    e1.SetAttribute("font-weight", header_fo, "bold");
                    e1.SetAttribute("font-weight-asian", header_style, "bold");
                    e1.SetAttribute("font-weight-complex", header_style, "bold");

                    e.AppendChild(e1);
                    break;
                }
            }
        }
        */
    }
}
