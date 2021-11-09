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
    public class ImgToXml
    {
        public ImgToXml()
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
        void SaveImage(Image image, string filePath)
        {
            string fileExtension = Path.GetExtension(filePath);
            switch (fileExtension.ToLower())
            {

                case ".bmp": image.Save(filePath, ImageFormat.Bmp); break;
                case ".exif": image.Save(filePath, ImageFormat.Exif); break;
                case ".gif": image.Save(filePath, ImageFormat.Gif); break;
                case ".jpg": case ".jpeg": image.Save(filePath, ImageFormat.Jpeg); break;
                case ".png": image.Save(filePath, ImageFormat.Png); break;
                case ".tif": case ".tiff": image.Save(filePath, ImageFormat.Tiff); break;

            }
        }


        public void Run(XmlManager xm,JToken binJson, JToken bodyJson, JToken docJson, bool zeroCheck)
        {
            //Image있는지 체크

            bool hasImg = false;
            int hasImgControlNum;
            int ImgcontrolListCount = 0;
            try
            {
                ImgcontrolListCount = bodyJson["controlList"].Count();
            }
            catch (System.ArgumentNullException e)
            {
                ImgcontrolListCount = 0;
            }
            for (int controlList = 0; controlList < ImgcontrolListCount; controlList++)
            {
                try
                {
                    object existImg = bodyJson["controlList"][controlList]["shapeComponentPicture"].Value<object>();
                    hasImg = true;
                    hasImgControlNum = controlList;
                }
                catch (System.ArgumentNullException e)
                {
                    hasImg = false;
                }
                //image가 존재할때만 실행
                if (hasImg)
                {
                    // int borderColor = bodyjson["controlList"][controlList]["shapeComponentPicture"]["borderColor"].Value<int>();
                    int ID = bodyJson["controlList"][controlList]["shapeComponentPicture"]["pictureInfo"]["binItemID"].Value<int>();
                    int imgWidth = bodyJson["controlList"][controlList]["header"]["width"].Value<int>();
                    int imgHeight = bodyJson["controlList"][controlList]["header"]["height"].Value<int>();
                    // String borderProperty = bodyJson["controlList"][controlList]["shapeComponentPicture"]["borderColor"].Value<String>();
                    int leftTopX = bodyJson["controlList"][controlList]["shapeComponentPicture"]["leftTop"]["x"].Value<int>();
                    int leftTopY = bodyJson["controlList"][controlList]["shapeComponentPicture"]["leftTop"]["y"].Value<int>();
                    int rightTopX = bodyJson["controlList"][controlList]["shapeComponentPicture"]["rightTop"]["x"].Value<int>();
                    int rightTopY = bodyJson["controlList"][controlList]["shapeComponentPicture"]["rightTop"]["y"].Value<int>();
                    int leftBottomX = bodyJson["controlList"][controlList]["shapeComponentPicture"]["leftBottom"]["x"].Value<int>();
                    int leftBottomY = bodyJson["controlList"][controlList]["shapeComponentPicture"]["leftBottom"]["y"].Value<int>();
                    int rightBottomX = bodyJson["controlList"][controlList]["shapeComponentPicture"]["rightBottom"]["x"].Value<int>();
                    int rightBottomY = bodyJson["controlList"][controlList]["shapeComponentPicture"]["rightBottom"]["y"].Value<int>();


                        JArray temp = binJson["embeddedBinaryDataList"][ID-1]["data"].Value<JArray>();
                        sbyte[] items = temp.ToObject<sbyte[]>();

                        Image img1 = StringToImage(items);
                        String name = binJson["embeddedBinaryDataList"][ID-1]["name"].Value<String>();
                        string path = Application.StartupPath + @"\New File\Pictures\" + name;
                    String extension =  docJson["binDataList"][ID - 1]["extensionForEmbedding"].Value<String>();
                        SaveImage(img1, path);
                    
                    double lx = Math.Round(leftTopX * 2.54 /7200 , 3);
                    double ly = Math.Round(leftTopY * 2.54 / 7200, 3);
                    
                    double width = Math.Round(imgWidth * 2.54 / 7200, 3);    
                    double height = Math.Round(imgHeight * 2.54 / 7200, 3);
                    string currentPath = "Pictures/" + name;
                    xm.imgstyle(ID-1);
                    xm.makeimg(width, height, extension, currentPath, lx, ly);
                    

                    }
                }

            }
            
    }
}

