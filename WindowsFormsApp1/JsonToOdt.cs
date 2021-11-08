using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Windows.Forms;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml;

namespace WindowsFormsApp1
{
    public class JsonToOdt
    {
        JObject json;

        XmlManager xm;

        FuncToXml.TextToXml ttx;
        FuncToXml.GridToXml gtx;

        public JsonToOdt()
        {
            xm = new XmlManager();

            ttx = new FuncToXml.TextToXml();
            gtx = new FuncToXml.GridToXml();
        }

        public void Run()
        {
            CreateFolder();

            for (int s = 0; s < json["bodyText"]["sectionList"].Count(); s++)
                for (int i = 0; i < json["bodyText"]["sectionList"][s]["paragraphList"].Count(); i++)
                {
                    bool zeroCheck = i == 0 ? true : false;

                    JToken nowJson = json["bodyText"]["sectionList"][s]["paragraphList"][i];
                    JToken docJson = json["docInfo"];

                    ttx.Run(xm, nowJson, docJson, zeroCheck);
                    gtx.Run(xm, nowJson, docJson, zeroCheck);
                }

            //ttx.Run(xm, json);

            SaveODT();
        }

        private void SaveODT()
        {
            xm.SaveODT(xm.root);
        }

        private void CreateFolder()
        {
            xm.CreateODT();
        }

        public void setJson(JObject j)
        {
            this.json = j;
        }
    }
}
