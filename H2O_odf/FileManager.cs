using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
//using System.Windows.Forms;
using System.IO;
using System.IO.Compression;
using System.Linq;

namespace H2O__
{
    class FileManager
    {
        public FileManager()
        {
            Initial_Folder();
        }

        private string appPath;
        private string fileName;
        private string loadName;
        private string jsonpath = @"C:\Users\park\source\repos\H2O_odf\H2O_odf\test.json";

        private void Initial_Folder()
        {
            appPath = @"C:\Users\park\source\repos\H2O_odf\H2O_odf";

            StreamReader file = File.OpenText(jsonpath);
            JsonTextReader reader = new JsonTextReader(file);


            JObject json = (JObject)JToken.ReadFrom(reader);

            //fileName = @"\TestODT";
            //loadName = @"\data";

            XmlManager xm = new XmlManager();

            xm.CreateODT();

            make(xm, json);

            xm.SaveODT(xm.root);


            try
            {
                ZipFile.CreateFromDirectory(appPath + @"\New File", appPath + @"\New.odt");
            }
            catch (IOException)
            {
                File.Delete(appPath + @"\New.odt");
                ZipFile.CreateFromDirectory(appPath + @"\New File", appPath + @"\New.odt");
            }

            //xm.Create_Document(appPath + loadName);

            //xm.Update_Text("Test Test one two one two");

            //xm.Save_Document(appPath + fileName);
        }

        private void Initial_XML()
        {

        }

        private void make(XmlManager xm, JObject json)
        {
            // p 생성
            for (int i = 0; i < json["BodyText"]["Section_0"]["HWPTAG_PARA_LINE_SEG"]["PARA LINE SEG"].Count(); i++)
            {
                string pcontent; // p text
                try
                {
                    pcontent = json["BodyText"]["Section_0"]["HWPTAG_PARA_TEXT"]["PARA"]["PARA " + i]["Text"].ToString();
                }
                catch(Exception)
                {
                    xm.AddContentP(); // para null 인 경우 처리
                    continue;
                }

                int shpaeID;
                int spancount = json["BodyText"]["Section_0"]["PARAMETER_List"]["PARA_" + i + "_HWPTAG_PARA_CHAR_SHAPE"]["PositonShapeIdPairList"].Count(); // p 안 text style count
                int pstyle = json["BodyText"]["Section_0"]["PARAMETER_List"]["PARA_" + i + "_HWPTAG_PARA_CHAR_SHAPE"]["PositonShapeIdPairList"]["PositonShapeIdPairList_0"]["ShapeId"].Value<int>();
                int current_position;
                int next_position;
                string name = "P" + (i + 1);

                // span 생성

                for (int j = 0; j < spancount; j++)
                {
                    current_position = json["BodyText"]["Section_0"]["PARAMETER_List"]["PARA_" + i + "_HWPTAG_PARA_CHAR_SHAPE"]["PositonShapeIdPairList"]["PositonShapeIdPairList_" + j]["Position"].Value<int>();
                    string subcontent;
                    int currentstyle = json["BodyText"]["Section_0"]["PARAMETER_List"]["PARA_" + i + "_HWPTAG_PARA_CHAR_SHAPE"]["PositonShapeIdPairList"]["PositonShapeIdPairList_" + j]["ShapeId"].Value<int>();
                    if (j < spancount - 1)
                    {
                        next_position = json["BodyText"]["Section_0"]["PARAMETER_List"]["PARA_" + i + "_HWPTAG_PARA_CHAR_SHAPE"]["PositonShapeIdPairList"]["PositonShapeIdPairList_" + (j + 1)]["Position"].Value<int>();
                        if (i == 0)
                        {
                            current_position = current_position - 16 < 0 ? 0 : current_position - 16;
                            next_position -= 16;
                        }
                        subcontent = pcontent.Substring(current_position, next_position - current_position);
                    }
                    else
                    {
                        if (i == 0 && j != 0)
                        {
                            current_position -= 16;
                        }
                        subcontent = pcontent.Substring(current_position);
                    }

                    if (j == 0)
                        name = xm.AddContentP(subcontent);
                    else if (currentstyle == pstyle)
                        xm.AddContentP(i, subcontent);
                    else
                        name = xm.AddContentSpan(name, subcontent);

                    //sytle 적용

                    bool bold = json["DocInfo 2"]["HWPTAG_CHAR_SHAPE"]["CHAR_SHAPE"]["CHAR_SHAPE_" + currentstyle]["Property"]["isBold"].Value<bool>();
                    bool italic = json["DocInfo 2"]["HWPTAG_CHAR_SHAPE"]["CHAR_SHAPE"]["CHAR_SHAPE_" + currentstyle]["Property"]["isItalic"].Value<bool>();

                    if (bold)
                        xm.SetBold(name);
                    if (italic)
                        xm.SetItalic(name);

                    xm.SetFontSize(name, json["DocInfo 2"]["HWPTAG_CHAR_SHAPE"]["CHAR_SHAPE"]["CHAR_SHAPE_" + currentstyle]["BaseSize"].Value<float>() / 100);

                }
            }
        }
    }
}