using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml;

namespace WindowsFormsApp1.FuncToXml
{
    public class HeadToXml
    {
        public HeadToXml()
        {

        }

        JToken jsonbinData;
        JToken jsoncontrol;
        //JToken jsonHeader;


        private void setData()
        {
        }
    
        public int bitcal(int i, int shift, byte b) //대부분 property에 있는 value를 bit별로 나누어서 넣을때 
        /** i = value 값  shift = 이동갯수 b = &연산할 값 */
        {
            byte temp = (byte)(i >> shift);
            return (int)(temp & b);
        }
        Image StringToImage(sbyte[] _Image)
        {

            sbyte[] bitmapData = new sbyte[_Image.Length];
            MemoryStream ms = new MemoryStream((byte[])((Array)_Image));
            Image returnImage = Image.FromStream(ms);
            return returnImage;

        }


        public void Run(XmlManager xm, JToken binJson,  JToken bodyJson, JToken docJson, JToken hJson, JToken fJson ,int hi, int fi)
        {
            //Image있는지 체크
            bool hasImg = false;
            
            
            
            //머리말은 스타일에들어감
            //머리말
            bool headerhasImg = false;
            bool hasHeader = false;
            int headerImgcontrolListCount = 0;
            int hasHeadernum = 0;
            int headerhasImgControlNum = 0;
           
                    
                if (hJson != null)
                {
                    
                    hasHeader = true;
                    hasHeadernum = hi;
                    for (int i = 0; i < hJson["controlList"][hasHeadernum]["paragraphList"]["paragraphList"].Count(); i++)
                        for (int j = 0; j < hJson["controlList"][hasHeadernum]["paragraphList"]["paragraphList"][i]["controlList"].Count(); j++)
                        {
                            try
                            {

                                object existImg = hJson["controlList"][hasHeadernum]["paragraphList"]["paragraphList"][i]["controlList"][j]["shapeComponentPicture"].Value<object>();
                                headerhasImg = true;
                               

                            }
                            catch (System.ArgumentNullException e)
                            {
                                headerhasImg = false;
                            }

                            //image가 존재할때만 실행
                            if (headerhasImg)
                            {
                                // int borderColor = hJson["controlList"][hasHeadernum]["paragraphList"]["paragraphList"][i]["controlList"][j]["shapeComponentPicture"]["borderColor"].Value<int>();
                                int ID = hJson["controlList"][hasHeadernum]["paragraphList"]["paragraphList"][i]["controlList"][j]["shapeComponentPicture"]["pictureInfo"]["binItemID"].Value<int>();
                                int IDindex = 0;
                                int imgWidth = hJson["controlList"][hasHeadernum]["paragraphList"]["paragraphList"][i]["controlList"][j]["header"]["width"].Value<int>();
                                int imgHeight = hJson["controlList"][hasHeadernum]["paragraphList"]["paragraphList"][i]["controlList"][j]["header"]["height"].Value<int>();
                                // String borderProperty = hJson["controlList"][hasHeadernum]["paragraphList"]["paragraphList"][i]["controlList"][j]["shapeComponentPicture"]["borderColor"].Value<String>();
                                double X = hJson["controlList"][hasHeadernum]["paragraphList"]["paragraphList"][i]["controlList"][j]["header"]["xOffset"].Value<int>();
                                double Y = hJson["controlList"][hasHeadernum]["paragraphList"]["paragraphList"][i]["controlList"][j]["header"]["yOffset"].Value<int>();
                                int property = hJson["controlList"][hasHeadernum]["paragraphList"]["paragraphList"][i]["controlList"][j]["header"]["property"]["value"].Value<int>();
                                int VertiRelTo = bitcal(property, 3, 0x3);
                                int VertiRelToarray = bitcal(property, 5, 0x7);
                                int HorzRelTo = bitcal(property, 8, 0x3);
                                int HorzRelToarray = bitcal(property, 10, 0x7);
                                int geulja = bitcal(property, 0, 0x1);
                                int through = bitcal(property, 21, 0x7);
                                int zindex = hJson["controlList"][hasHeadernum]["paragraphList"]["paragraphList"][i]["controlList"][j]["header"]["zOrder"].Value<int>();
                                int linevertical = 0;
                                if (hJson["controlList"][hasHeadernum]["paragraphList"]["paragraphList"][i]["lineSeg"] != null)
                                    linevertical = hJson["controlList"][hasHeadernum]["paragraphList"]["paragraphList"][i]["lineSeg"]["lineSegItemList"][0]["lineVerticalPosition"].Value<int>();
                                
                                for (int k = 0; k < binJson["embeddedBinaryDataList"].Count(); k++)
                                {
                                    String temp2 = binJson["embeddedBinaryDataList"][k]["name"].Value<String>();
                                    String tt = temp2.Substring(3, 4);

                                    if (ID == Convert.ToInt16(tt,16))
                                        IDindex = k;

                                }
                                String name = binJson["embeddedBinaryDataList"][IDindex]["name"].Value<String>();

                                JArray temp = binJson["embeddedBinaryDataList"][IDindex]["data"].Value<JArray>();
                                sbyte[] items = temp.ToObject<sbyte[]>();

                                //ImgNode 생성 및 Pictures폴더 child 설정
                                ImgNode node = setPicturesChild(xm, name);
                                node.img = StringToImage(items);

                                String extension = docJson["binDataList"][IDindex]["extensionForEmbedding"].Value<String>();


                                double width = Math.Round(imgWidth * 2.54 / 7200, 3);
                                double height = Math.Round((imgHeight) * 2.54 / 7200, 3);
                                string currentPath = "Pictures/" + name;
                                xm.imgstyle(VertiRelTo, VertiRelToarray, HorzRelTo, HorzRelToarray, through,true,false);
                                xm.makeimg(width, height, extension, currentPath, X, Y, geulja, zindex, linevertical, VertiRelTo, ID,true,false);

                            }
                        }
                }
                if (fJson != null)
                {
                    
                    hasHeader = true;
                    hasHeadernum = fi;
                    for (int i = 0; i < fJson["controlList"][hasHeadernum]["paragraphList"]["paragraphList"].Count(); i++)
                        for (int j = 0; j < fJson["controlList"][hasHeadernum]["paragraphList"]["paragraphList"][i]["controlList"].Count(); j++)
                        {
                            try
                            {

                                object existImg = fJson["controlList"][hasHeadernum]["paragraphList"]["paragraphList"][i]["controlList"][j]["shapeComponentPicture"].Value<object>();
                                headerhasImg = true;
                            

                            }
                            catch (System.ArgumentNullException e)
                            {
                                headerhasImg = false;
                            }

                            //image가 존재할때만 실행
                            if (headerhasImg)
                            {
                                // int borderColor = fJson["controlList"][hasHeadernum]["paragraphList"]["paragraphList"][i]["controlList"][j]["shapeComponentPicture"]["borderColor"].Value<int>();
                                int ID = fJson["controlList"][hasHeadernum]["paragraphList"]["paragraphList"][i]["controlList"][j]["shapeComponentPicture"]["pictureInfo"]["binItemID"].Value<int>();
                                int IDindex = 0;
                                int imgWidth = fJson["controlList"][hasHeadernum]["paragraphList"]["paragraphList"][i]["controlList"][j]["header"]["width"].Value<int>();
                                int imgHeight = fJson["controlList"][hasHeadernum]["paragraphList"]["paragraphList"][i]["controlList"][j]["header"]["height"].Value<int>();
                                // String borderProperty = fJson["controlList"][hasHeadernum]["paragraphList"]["paragraphList"][i]["controlList"][j]["shapeComponentPicture"]["borderColor"].Value<String>();
                                double X = fJson["controlList"][hasHeadernum]["paragraphList"]["paragraphList"][i]["controlList"][j]["header"]["xOffset"].Value<int>();
                                double Y = fJson["controlList"][hasHeadernum]["paragraphList"]["paragraphList"][i]["controlList"][j]["header"]["yOffset"].Value<int>();
                                int property = fJson["controlList"][hasHeadernum]["paragraphList"]["paragraphList"][i]["controlList"][j]["header"]["property"]["value"].Value<int>();
                                int VertiRelTo = bitcal(property, 3, 0x3);
                                int VertiRelToarray = bitcal(property, 5, 0x7);
                                int HorzRelTo = bitcal(property, 8, 0x3);
                                int HorzRelToarray = bitcal(property, 10, 0x7);
                                int geulja = bitcal(property, 0, 0x1);
                                int through = bitcal(property, 21, 0x7);
                                int zindex = fJson["controlList"][hasHeadernum]["paragraphList"]["paragraphList"][i]["controlList"][j]["header"]["zOrder"].Value<int>();

                                int linevertical = 0;
                                if(fJson["controlList"][hasHeadernum]["paragraphList"]["paragraphList"][i]["lineSeg"] != null)
                                for (int m = 0; m< fJson["controlList"][hasHeadernum]["paragraphList"]["paragraphList"][i]["lineSeg"]["lineSegItemList"].Count(); m++)
                                linevertical += fJson["controlList"][hasHeadernum]["paragraphList"]["paragraphList"][i]["lineSeg"]["lineSegItemList"][m]["lineVerticalPosition"].Value<int>();
                                for (int k = 0; k < binJson["embeddedBinaryDataList"].Count(); k++)
                                {
                                    String temp2 = binJson["embeddedBinaryDataList"][k]["name"].Value<String>();
                                    String tt = temp2.Substring(3, 4);
                                    UInt16 ta = Convert.ToUInt16(tt,16);
                                    if (ID == ta)
                                        IDindex = k;

                                }
                                String name = binJson["embeddedBinaryDataList"][IDindex]["name"].Value<String>();

                                JArray temp = binJson["embeddedBinaryDataList"][IDindex]["data"].Value<JArray>();
                                sbyte[] items = temp.ToObject<sbyte[]>();

                                //ImgNode 생성 및 Pictures폴더 child 설정
                                ImgNode node = setPicturesChild(xm, name);
                                node.img = StringToImage(items);

                                String extension = docJson["binDataList"][IDindex]["extensionForEmbedding"].Value<String>();
                                String text = fJson["controlList"][hasHeadernum]["paragraphList"]["paragraphList"][i]["text"].Value<String>();

                                double width = Math.Round(imgWidth * 2.54 / 7200, 3);
                                double height = Math.Round((imgHeight) * 2.54 / 7200, 3);
                                string currentPath = "Pictures/" + name;
                                xm.imgstyle(VertiRelTo, VertiRelToarray, HorzRelTo, HorzRelToarray, through, false, true);
                                xm.makeimg(width, height, extension, currentPath, X, Y, geulja, zindex, (int)linevertical, VertiRelTo, ID, false, true);

                            }
                        }
                }
            
        }

        public ImgNode setPicturesChild(XmlManager xm, string name)
        {
            FolderNode pictures = (FolderNode)xm.root.child["Pictures"];
            ImgNode node = new ImgNode(name, pictures);
            node.path = Application.StartupPath + @"\New File\Pictures\" + name;

            return node;
        }

    }
}
