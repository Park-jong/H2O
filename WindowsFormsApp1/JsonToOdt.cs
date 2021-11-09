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

            double pageMarginLeft = Math.Round(json["bodyText"]["sectionList"][0]["paragraphList"][0]["controlList"][0]["pageDef"]["leftMargin"].Value<double>() * 0.01 * 0.0352778, 3);
            double pageMarginRight = Math.Round(json["bodyText"]["sectionList"][0]["paragraphList"][0]["controlList"][0]["pageDef"]["rightMargin"].Value<double>() * 0.01 * 0.0352778, 3);
            double pageMarginTop = Math.Round(json["bodyText"]["sectionList"][0]["paragraphList"][0]["controlList"][0]["pageDef"]["topMargin"].Value<double>() * 0.01 * 0.0352778, 3);
            double pageMarginBottom = Math.Round(json["bodyText"]["sectionList"][0]["paragraphList"][0]["controlList"][0]["pageDef"]["bottomMargin"].Value<double>() * 0.01 * 0.0352778, 3);
            double pageMarginHeader = Math.Round(json["bodyText"]["sectionList"][0]["paragraphList"][0]["controlList"][0]["pageDef"]["headerMargin"].Value<double>() * 0.01 * 0.0352778, 3);
            double pageMarginFooter = Math.Round(json["bodyText"]["sectionList"][0]["paragraphList"][0]["controlList"][0]["pageDef"]["footerMargin"].Value<double>() * 0.01 * 0.0352778, 3);
            //머리 꼬리 없으면
            xm.SetPageLayout(pageMarginLeft, pageMarginRight, pageMarginTop + pageMarginHeader, pageMarginBottom + pageMarginFooter);
            //머리 꼬리 있으면
            /*수정예정
            {
                xm.ContentXml = false;
                xm.SetPageLayout(pageMarginLeft, pageMarginRight, pageMarginTop, pageMarginBottom);
                {
                    xm.SetHeader(pageMarginHeader, 0, 0, 0);
                    string name = xm.AddHeader();
                    name = xm.AddHeader();
                    name = xm.AddHeaderFooterSpan(name, "123");
                    name = xm.AddHeader();
                    name = xm.AddHeaderFooterSpan(name, "456");
                    xm.SetBold(name);
                }
                {
                    xm.SetFooter(pageMarginHeader, 0, 0, 0);
                    xm.AddFooter();
                    string name = xm.AddFooter();
                    xm.AddHeaderFooterSpan(name, "789");
                    name = xm.AddFooter();
                    xm.AddHeaderFooterSpan(name, "1011");
                }
                xm.ContentXml = true;
            }*/
            xm.ContentXml = true;
            // 문단 내 텍스트 별로 subcontent 만들기
            // p 생성



            for (int s = 0; s < json["bodyText"]["sectionList"].Count(); s++)
                for (int i = 0; i < json["bodyText"]["sectionList"][s]["paragraphList"].Count(); i++)
                {
                    bool zeroCheck = i == 0 ? true : false;

                    JToken nowJson = json["bodyText"]["sectionList"][s]["paragraphList"][i];
                    JToken docJson = json["docInfo"];

                    ttx.Run(xm, nowJson, docJson, zeroCheck);
                    gtx.Run(xm, nowJson, docJson, zeroCheck);
                }

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
