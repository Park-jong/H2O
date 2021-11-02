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
        JObject jsonT;

        XmlManager xm;

        FuncToXml.TextToXml ttx;

        public JsonToOdt()
        {
            xm = new XmlManager();

            ttx = new FuncToXml.TextToXml();
        }

        public void Run()
        {
            CreateFolder();

            ttx.Run(xm, json);


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

        public void setJsonT(JObject jT)
        {
            this.jsonT = jT;
        }

    }
}
