using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
//using System.Windows.Forms;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml;

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
            // 문단 내 텍스트 별로 subcontent 만들기
            // p 생성
            for (int i = 0; i < json["BodyText"]["Section_0"]["HWPTAG_PARA_LINE_SEG"]["PARA LINE SEG"].Count(); i++)
            {
                // 문단 ID 가져오기
                int shapeID = json["BodyText"]["Section_0"]["PARAMETER_List"]["PARA_" + i + "_HWPTAG_PARA_HEADER"]["ShapeId"].Value<int>();
                string sID = shapeID.ToString();

                //전체 텍스트 가져오기
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

                // 스타일별 텍스트 개수
                int spancount = json["BodyText"]["Section_0"]["PARAMETER_List"]["PARA_" + i + "_HWPTAG_PARA_CHAR_SHAPE"]["PositonShapeIdPairList"].Count(); // p 안 text style count
                // 스타일별 텍스트 아이디
                int pstyle = json["BodyText"]["Section_0"]["PARAMETER_List"]["PARA_" + i + "_HWPTAG_PARA_CHAR_SHAPE"]["PositonShapeIdPairList"]["PositonShapeIdPairList_0"]["ShapeId"].Value<int>();
                
                int current_position;
                int next_position;

                string name = "P" + (i + 1);

                // 텍스트별 위치 비교해서 자르기
                for (int j = 0; j < spancount; j++)
                {
                    //스타일이 시작되는 위치
                    current_position = json["BodyText"]["Section_0"]["PARAMETER_List"]["PARA_" + i + "_HWPTAG_PARA_CHAR_SHAPE"]["PositonShapeIdPairList"]["PositonShapeIdPairList_" + j]["Position"].Value<int>();

                    string subcontent;
                    if (j < spancount - 1)
                    {
                        next_position = json["BodyText"]["Section_0"]["PARAMETER_List"]["PARA_" + i + "_HWPTAG_PARA_CHAR_SHAPE"]["PositonShapeIdPairList"]["PositonShapeIdPairList_" + (j + 1)]["Position"].Value<int>();
                        if (i == 0)
                        {
                            current_position = current_position - 16 < 0 ? 0 : current_position - 16;
                            next_position -= 16;
                        }
                        subcontent = pcontent.Substring(current_position, next_position - current_position); //처음포지션부터 글자 수만큼 자르기
                    }
                    else //마지막 텍스트
                    {
                        if (i == 0 && j != 0)
                        {
                            current_position -= 16;
                        }
                        subcontent = pcontent.Substring(current_position);
                    }

                    //스타일 추가 및 내용 추가
                    //스타일의 아이디
                    int currentstyle = json["BodyText"]["Section_0"]["PARAMETER_List"]["PARA_" + i + "_HWPTAG_PARA_CHAR_SHAPE"]["PositonShapeIdPairList"]["PositonShapeIdPairList_" + j]["ShapeId"].Value<int>();

                    if (j == 0)
                        name = xm.AddContentP(subcontent);
                    else if (currentstyle == pstyle)
                        xm.AddContentP(i, subcontent); //pstyle과 같으면 텍스트만 추가
                    else
                        name = xm.AddContentSpan(name, subcontent); //pstyle과 다르면 span 생성 후 텍스트 추가


                    //스타일 속성 추가
                    //문단 속성 추가
                    if (j == 0)
                    {
                        //줄 간격
                        int lineSpace = json["DocInfo 2"]["HWPTAG_PARA_SHAPE"]["PARA_SHAPE"]["PARA_SHAPE_" + sID]["LineSpace"].Value<int>();
                        xm.Paragraph.SetLineSpace(name, (XmlDocument)xm.docs["content.xml"], lineSpace);

                        //문단 테두리 간격
                        double topborderSpace = json["DocInfo 2"]["HWPTAG_PARA_SHAPE"]["PARA_SHAPE"]["PARA_SHAPE_" + sID]["TopBorderSpace"].Value<double>();
                        double bottomBorderSpace = json["DocInfo 2"]["HWPTAG_PARA_SHAPE"]["PARA_SHAPE"]["PARA_SHAPE_" + sID]["BottomBorderSpace"].Value<double>();
                        double leftBorderSpace = json["DocInfo 2"]["HWPTAG_PARA_SHAPE"]["PARA_SHAPE"]["PARA_SHAPE_" + sID]["LeftBorderSpace"].Value<double>();
                        double rightBorderSpace = json["DocInfo 2"]["HWPTAG_PARA_SHAPE"]["PARA_SHAPE"]["PARA_SHAPE_" + sID]["RightBorderSpace"].Value<double>();

                        topborderSpace *= 0.01;
                        bottomBorderSpace *= 0.01;
                        leftBorderSpace *= 0.01;
                        rightBorderSpace *= 0.01;

                        xm.Paragraph.SetBorderSpace(name, (XmlDocument)xm.docs["content.xml"], topborderSpace, bottomBorderSpace, leftBorderSpace, rightBorderSpace);

                        //줄 나눔
                        //한글이 ByWord일 경우 적용 (한글, 영어 모두)
                        string byWord = json["DocInfo 2"]["HWPTAG_PARA_SHAPE"]["PARA_SHAPE"]["PARA_SHAPE_" + sID]["Property1"]["LineDivideForHangul"].Value<string>();
                        if (byWord.Equals("ByWord"))
                        {
                            xm.Paragraph.SetByWord(name, (XmlDocument)xm.docs["content.xml"]);
                        }

                        //외톨이줄 보호 여부
                        bool isProtectLoner = json["DocInfo 2"]["HWPTAG_PARA_SHAPE"]["PARA_SHAPE"]["PARA_SHAPE_" + sID]["Property1"]["isProtectLoner"].Value<bool>();
                        if (isProtectLoner)
                            xm.Paragraph.SetisProtectLoner(name, (XmlDocument)xm.docs["content.xml"]);

                    }

                    bool bold = json["DocInfo 2"]["HWPTAG_CHAR_SHAPE"]["CHAR_SHAPE"]["CHAR_SHAPE_" + currentstyle]["Property"]["isBold"].Value<bool>();
                    bool italic = json["DocInfo 2"]["HWPTAG_CHAR_SHAPE"]["CHAR_SHAPE"]["CHAR_SHAPE_" + currentstyle]["Property"]["isItalic"].Value<bool>();
                    bool kerning = json["DocInfo 2"]["HWPTAG_CHAR_SHAPE"]["CHAR_SHAPE"]["CHAR_SHAPE_" + currentstyle]["Property"]["isKerning"].Value<bool>();

                    if (bold)
                        xm.SetBold(name);
                    if (italic)
                        xm.SetItalic(name);
                    if (kerning)
                    {
                        double baseSize = json["DocInfo 2"]["HWPTAG_CHAR_SHAPE"]["CHAR_SHAPE"]["CHAR_SHAPE_" + currentstyle]["BaseSize"].Value<int>(); // pt * 100 값
                        double kerningSpace = json["DocInfo 2"]["HWPTAG_CHAR_SHAPE"]["CHAR_SHAPE"]["CHAR_SHAPE_" + currentstyle]["CharSpacebyLanguage"]["Hangul"].Value<int>(); // kerning percent 값

                        double value = baseSize * (kerningSpace / 100) / 100 * 0.0353; //pt로 환산 후 cm로 환산
                        xm.SetKerning(name, baseSize.ToString() + "cm");
                    }
                }
            }
        }
    }
}
