using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace WindowsFormsApp1.FuncToXml
{
    public class TextToXml  
    {
        public TextToXml()
        {
        }

        public int bitcal(int i, int shift, byte b) //대부분 property에 있는 value를 bit별로 나누어서 넣을때 
        /** i = value 값  shift = 이동갯수 b = &연산할 값 */
        {
            byte temp = (byte)(i >> shift);
            return (int)(temp & b);
        }

        public void Run(XmlManager xm, JObject json)
        {
            double pageMarginLeft = Math.Round(json["bodyText"]["sectionList"][0]["paragraphList"][0]["controlList"][0]["pageDef"]["leftMargin"].Value<double>() * 0.01 * 0.0352778, 3);
            double pageMarginRight = Math.Round(json["bodyText"]["sectionList"][0]["paragraphList"][0]["controlList"][0]["pageDef"]["rightMargin"].Value<double>() * 0.01 * 0.0352778, 3);
            double pageMarginTop = Math.Round(json["bodyText"]["sectionList"][0]["paragraphList"][0]["controlList"][0]["pageDef"]["topMargin"].Value<double>() * 0.01 * 0.0352778, 3);
            double pageMarginBottom = Math.Round(json["bodyText"]["sectionList"][0]["paragraphList"][0]["controlList"][0]["pageDef"]["bottomMargin"].Value<double>() * 0.01 * 0.0352778, 3);
            double pageMarginHeader = Math.Round(json["bodyText"]["sectionList"][0]["paragraphList"][0]["controlList"][0]["pageDef"]["headerMargin"].Value<double>() * 0.01 * 0.0352778, 3);
            double pageMarginFooter = Math.Round(json["bodyText"]["sectionList"][0]["paragraphList"][0]["controlList"][0]["pageDef"]["footerMargin"].Value<double>() * 0.01 * 0.0352778, 3);
            xm.SetPageLayout(pageMarginLeft, pageMarginRight, pageMarginTop + pageMarginHeader, pageMarginBottom + pageMarginFooter);
            xm.ContentXml = true;

            // 문단 내 텍스트 별로 subcontent 만들기
            // p 생성
            for (int s = 0; s < json["bodyText"]["sectionList"].Count(); s++)
                for (int i = 0; i < json["bodyText"]["sectionList"][s]["paragraphList"].Count(); i++)
                {
                    // 문단 ID 가져오기
                    int shapeId = json["bodyText"]["sectionList"][s]["paragraphList"][i]["header"]["paraShapeId"].Value<int>();
                    int sID = shapeId;

                    //전체 텍스트 가져오기
                    string pcontent = null; // p text
                    try
                    {
                        object obj = json["bodyText"]["sectionList"][s]["paragraphList"][i]["text"].Value<object>();

                        if (obj != null)
                            pcontent = json["bodyText"]["sectionList"][s]["paragraphList"][i]["text"].Value<string>();
                    }
                    /////////////////////////////////////////////////////////////글자가 없을 때
                    catch (Exception e)
                    {
                        int nullPStyle = json["bodyText"]["sectionList"][s]["paragraphList"][i]["charShape"]["positionShapeIdPairList"][0]["shapeId"].Value<int>();
                        float nullPFontSize = json["docInfo"]["charShapeList"][nullPStyle]["baseSize"].Value<float>() / 100;
                        string nullPName = xm.AddContentP(); // para null 인 경우 처리
                        xm.SetFontSize(nullPName, nullPFontSize);

                        //줄 간격
                        int lineSpace = json["docInfo"]["paraShapeList"][sID]["lineSpace"].Value<int>();
                        xm.Paragraph.SetLineSpace(nullPName, (XmlDocument)xm.docs["content.xml"], lineSpace * 0.01 * nullPFontSize * 0.03527);

                        //문단 테두리 간격
                        double topborderSpace = json["docInfo"]["paraShapeList"][sID]["topBorderSpace"].Value<double>();
                        double bottomBorderSpace = json["docInfo"]["paraShapeList"][sID]["bottomBorderSpace"].Value<double>();
                        double leftBorderSpace = json["docInfo"]["paraShapeList"][sID]["leftBorderSpace"].Value<double>();
                        double rightBorderSpace = json["docInfo"]["paraShapeList"][sID]["rightBorderSpace"].Value<double>();

                        topborderSpace *= 0.01;
                        bottomBorderSpace *= 0.01;
                        leftBorderSpace *= 0.01;
                        rightBorderSpace *= 0.01;

                        xm.Paragraph.SetBorderSpace(nullPName, (XmlDocument)xm.docs["content.xml"], topborderSpace, bottomBorderSpace, leftBorderSpace, rightBorderSpace);

                        //정렬
                        int temp = json["docInfo"]["paraShapeList"][sID]["property1"]["value"].Value<int>();

                        int paraAlign = bitcal(temp, 2, 0x7);
                        if (paraAlign.Equals(2))
                            xm.SetPAlign(nullPName, "end");
                        else if (paraAlign.Equals(0))
                            xm.SetPAlign(nullPName, "justify");
                        else if (paraAlign.Equals(3))
                            xm.SetPAlign(nullPName, "center");
                        else if (paraAlign.Equals(4))
                            xm.SetPAlign(nullPName, "Divide");
                        else if (paraAlign.Equals(5))
                            xm.SetPAlign(nullPName, "Distribute");
                        else if (paraAlign.Equals(1))
                            xm.SetPAlign(nullPName, "Distribute");

                        //첫줄 들여쓰기
                        //margin
                        double indent = json["docInfo"]["paraShapeList"][sID]["indent"].Value<double>();
                        double topspace = json["docInfo"]["paraShapeList"][sID]["topParaSpace"].Value<double>();
                        double bottomspace = json["docInfo"]["paraShapeList"][sID]["bottomParaSpace"].Value<double>();
                        double leftmargin = json["docInfo"]["paraShapeList"][sID]["leftMargin"].Value<double>();
                        double rightmargin = json["docInfo"]["paraShapeList"][sID]["rightMargin"].Value<double>();
                        if (indent != 0 || topspace != 0 || bottomspace != 0 || leftmargin != 0 || rightmargin != 0)
                        {
                            xm.SetPIndent(nullPName, (float)(indent / 200 * 0.0353));
                            if (indent < 0)
                                leftmargin = -indent + leftmargin;
                            xm.SetPMargin(nullPName, (float)(leftmargin / 200 * 0.0353), (float)(rightmargin / 200 * 0.0353), (float)(topspace / 200 * 0.0353), (float)(bottomspace / 200 * 0.0353));
                        }

                        continue;
                    }
                    /////////////////////////////////////////////////////////////////

                    // 스타일별 텍스트 개수
                    int spancount = json["bodyText"]["sectionList"][s]["paragraphList"][i]["charShape"]["positionShapeIdPairList"].Count(); // p 안 text style count
                                                                                                                                            // 스타일별 텍스트 아이디
                    int pstyle = json["bodyText"]["sectionList"][s]["paragraphList"][i]["charShape"]["positionShapeIdPairList"][0]["shapeId"].Value<int>();

                    int current_position;
                    int next_position;

                    string pname = "";
                    string name;

                    // 텍스트별 위치 비교해서 자르기
                    for (int j = 0; j < spancount; j++)
                    {
                        //스타일이 시작되는 위치
                        current_position = json["bodyText"]["sectionList"][s]["paragraphList"][i]["charShape"]["positionShapeIdPairList"][j]["position"].Value<int>();

                        string subcontent;
                        if (j < spancount - 1)
                        {
                            next_position = json["bodyText"]["sectionList"][s]["paragraphList"][i]["charShape"]["positionShapeIdPairList"][j + 1]["position"].Value<int>();
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
                        int currentstyle = json["bodyText"]["sectionList"][s]["paragraphList"][i]["charShape"]["positionShapeIdPairList"][j]["shapeId"].Value<int>();

                        /////////////////////////////////////
                        /*if (j == 0)
                            name = xm.AddContentP(subcontent);
                        else if (currentstyle == pstyle)
                            xm.AddContentP(i, subcontent); //pstyle과 같으면 텍스트만 추가
                        else
                            name = xm.AddContentSpan(pname, subcontent); //pstyle과 다르면 span 생성 후 텍스트 추가*/
                        ///////////////////////////////////////////////추가 수정 필요
                        if (spancount == 1)
                        {
                            pname = xm.AddContentP(pcontent);
                            name = pname;
                        }
                        else if (j == 0)
                        {
                            pname = xm.AddContentP("");
                            name = xm.AddContentSpan(pname, subcontent);
                        }
                        else
                            name = xm.AddContentSpan(pname, subcontent);
                        //


                        //스타일 속성 추가
                        //문단 속성 추가
                        double baseSize = json["docInfo"]["charShapeList"][currentstyle]["baseSize"].Value<double>(); // pt * 100 값

                        if (j == 0)
                        {
                            //줄 간격
                            int lineSpace = json["docInfo"]["paraShapeList"][sID]["lineSpace"].Value<int>();
                            xm.Paragraph.SetLineSpace(pname, (XmlDocument)xm.docs["content.xml"], lineSpace * 0.01 * baseSize * 0.01 * 0.03527);

                            //문단 테두리 간격
                            double topborderSpace = json["docInfo"]["paraShapeList"][sID]["topBorderSpace"].Value<double>();
                            double bottomBorderSpace = json["docInfo"]["paraShapeList"][sID]["bottomBorderSpace"].Value<double>();
                            double leftBorderSpace = json["docInfo"]["paraShapeList"][sID]["leftBorderSpace"].Value<double>();
                            double rightBorderSpace = json["docInfo"]["paraShapeList"][sID]["rightBorderSpace"].Value<double>();

                            topborderSpace *= 0.01;
                            bottomBorderSpace *= 0.01;
                            leftBorderSpace *= 0.01;
                            rightBorderSpace *= 0.01;

                            xm.Paragraph.SetBorderSpace(pname, (XmlDocument)xm.docs["content.xml"], topborderSpace, bottomBorderSpace, leftBorderSpace, rightBorderSpace);

                            //줄 나눔
                            //한글이 ByWord일 경우 적용 (한글, 영어 모두)



                            int temp = json["docInfo"]["paraShapeList"][sID]["property1"]["value"].Value<int>();

                            Byte byWord = (Byte)(temp >> 7);
                            if (byWord.Equals(1))
                            {
                                xm.Paragraph.SetByWord(pname, (XmlDocument)xm.docs["content.xml"]);
                            }

                            //외톨이줄 보호 여부
                            int isProtectLoner = json["docInfo"]["paraShapeList"][sID]["property1"]["value"].Value<int>();
                            int protemp = bitcal(isProtectLoner, 15, 1);
                            if (protemp == 1)
                                xm.Paragraph.SetisProtectLoner(pname, (XmlDocument)xm.docs["content.xml"]);

                            //다음 문단과 함께 여부
                            int isTogetherNextPara = json["docInfo"]["paraShapeList"][sID]["property1"]["value"].Value<int>();
                            int togtemp = bitcal(isTogetherNextPara, 16, 1);
                            if (togtemp == 1)
                                xm.Paragraph.SetisTogetherNextPara(pname, (XmlDocument)xm.docs["content.xml"]);

                            //문단 보호 여부
                            int isProtectPara = json["docInfo"]["paraShapeList"][sID]["property1"]["value"].Value<int>();
                            int paratemp = bitcal(isProtectPara, 17, 1);
                            if (paratemp == 1)
                                xm.Paragraph.SetisProtectPara(pname, (XmlDocument)xm.docs["content.xml"]);

                            //한글과 영어 간격을 자동 조절 여부
                            int isAutoAdjustGapHangulEnglish = json["docInfo"]["paraShapeList"][sID]["property2"]["value"].Value<int>();
                            int autotemp = bitcal(isAutoAdjustGapHangulEnglish, 4, 0x1);
                            if (autotemp == 1)
                                xm.Paragraph.SetisAutoAdjustGapHangulEnglish(pname, (XmlDocument)xm.docs["content.xml"]);


                            //정렬
                            int aligntemp = json["docInfo"]["paraShapeList"][sID]["property1"]["value"].Value<int>();
                            int paraAlign = bitcal(aligntemp, 2, 0x7);
                            if (paraAlign.Equals(2))
                                xm.SetPAlign(pname, "end");
                            else if (paraAlign.Equals(0))
                                xm.SetPAlign(pname, "justify");
                            else if (paraAlign.Equals(3))
                                xm.SetPAlign(pname, "center");
                            else if (paraAlign.Equals(5))
                                xm.SetPAlign(pname, "Divide");
                            else if (paraAlign.Equals(4))
                                xm.SetPAlign(pname, "Distribute");

                            //첫줄 들여쓰기
                            //margin
                            double indent = json["docInfo"]["paraShapeList"][sID]["indent"].Value<double>();
                            double topspace = json["docInfo"]["paraShapeList"][sID]["topParaSpace"].Value<double>();
                            double bottomspace = json["docInfo"]["paraShapeList"][sID]["bottomParaSpace"].Value<double>();
                            double leftmargin = json["docInfo"]["paraShapeList"][sID]["leftMargin"].Value<double>();
                            double rightmargin = json["docInfo"]["paraShapeList"][sID]["rightMargin"].Value<double>();
                            if (indent != 0 || topspace != 0 || bottomspace != 0 || leftmargin != 0 || rightmargin != 0)
                            {
                                xm.SetPIndent(pname, (float)(indent / 200 * 0.0353));
                                if (indent < 0)
                                    leftmargin = -indent + leftmargin;
                                xm.SetPMargin(pname, (float)(leftmargin / 200 * 0.0353), (float)(rightmargin / 200 * 0.0353), (float)(topspace / 200 * 0.0353), (float)(bottomspace / 200 * 0.0353));
                            }



                        }
                        int charPro = json["docInfo"]["charShapeList"][currentstyle]["property"]["value"].Value<int>();

                        int bold = bitcal(charPro, 1, 0x1);
                        int italic = bitcal(charPro, 1, 0x1);
                        int underline = bitcal(charPro, 2, 0x3);
                        //bool kerning = json["docInfo"]["charShapeList"][currentstyle]["Property"]["isKerning"].Value<bool>();

                        int fontcolor = json["docInfo"]["charShapeList"][currentstyle]["charColor"]["value"].Value<int>();
                        int strikeline = bitcal(charPro, 18, 0x7);


                        if (true)
                        {
                            //진하게
                            if (bold == 1)
                                xm.SetBold(name);
                            //기울임
                            if (italic == 1)
                                xm.SetItalic(name);

                            //커닝 odt에서 지원하지 않는 기능
                            /*
                            if (kerning)
                            {
                                baseSize = json["docInfo"]["charShapeList"][currentstyle]["baseSize"].Value<int>(); // pt * 100 값
                                double kerningSpace = json["docInfo"]["charShapeList"][currentstyle]["CharSpacebyLanguage"]["Hangul"].Value<int>(); // kerning percent 값

                                double value = baseSize * (kerningSpace / 100) / 100 * 0.0353; //pt로 환산 후 cm로 환산
                                xm.SetKerning(name, baseSize.ToString() + "cm");
                            }
                            */

                            //폰트 사이즈, 폰트 색
                            xm.SetFontSize(name, json["docInfo"]["charShapeList"][currentstyle]["baseSize"].Value<float>() / 100);
                            if (fontcolor != 0)
                            {
                                byte[] bit = BitConverter.GetBytes(fontcolor);
                                Array.Reverse(bit);
                                fontcolor = BitConverter.ToInt32(bit, 0);
                                xm.SetFontColor(name, "#" + fontcolor.ToString("X8").Substring(0, 6));
                            }


                            //윗줄 밑줄
                            if (underline.Equals(1))
                            {
                                int templineshape = json["docInfo"]["charShapeList"][currentstyle]["property"]["value"].Value<int>();
                                int lineshape = bitcal(templineshape, 4, 0xf);
                                int linecolor = json["docInfo"]["charShapeList"][currentstyle]["underLineColor"]["value"].Value<int>();
                                byte[] bit = BitConverter.GetBytes(linecolor);
                                Array.Reverse(bit);
                                linecolor = BitConverter.ToInt32(bit, 0);
                                ///lineshape 값? 왜 이렇게들어가느지
                                xm.SetUnderline(name, lineshape, "#" + linecolor.ToString("X8").Substring(0, 6));
                            }
                            else if (underline.Equals(3))
                            {
                                int templineshape = json["docInfo"]["charShapeList"][currentstyle]["property"]["value"].Value<int>();
                                int lineshape = bitcal(templineshape, 4, 0xf);
                                int linecolor = json["docInfo"]["charShapeList"][currentstyle]["underLineColor"]["value"].Value<int>();
                                byte[] bit = BitConverter.GetBytes(linecolor);
                                Array.Reverse(bit);
                                linecolor = BitConverter.ToInt32(bit, 0);
                                xm.SetUnderline(name, lineshape, "#" + linecolor.ToString("X8").Substring(0, 6));
                            }

                            //취소선 odt에서는 종류 제한, 색상 선택 불가
                            if (strikeline > 0)
                            {
                                int templineshape = json["docInfo"]["charShapeList"][currentstyle]["property"]["value"].Value<int>();
                                int lineshape = bitcal(templineshape, 4, 0xf);
                                xm.SetThroughline(name, lineshape);
                            }

                            //외곽선 한글에 종류 여러개지만 odt에서는 한개
                            if (bitcal((json["docInfo"]["charShapeList"][currentstyle]["property"]["value"].Value<int>()), 8, 0x7) != 0)
                            {
                                xm.SetOutline(name);
                            }

                            //그림자 odt한종류
                            if (bitcal((json["docInfo"]["charShapeList"][currentstyle]["property"]["value"].Value<int>()), 11, 0x3) != 0)
                            {
                                xm.SetShadow(name);
                            }

                            //음각 양각
                            if (bitcal((json["docInfo"]["charShapeList"][currentstyle]["property"]["value"].Value<int>()), 13, 0x1) != 0)
                            {
                                xm.SetRelief(name);
                            }
                            else if (bitcal((json["docInfo"]["charShapeList"][currentstyle]["property"]["value"].Value<int>()), 14, 0x1) != 0)
                            {
                                xm.SetRelief(name, "engraved");
                            }

                            //위첨자
                            //아래첨자
                            if (bitcal((json["docInfo"]["charShapeList"][currentstyle]["property"]["value"].Value<int>()), 15, 0x1) != 0)
                            {
                                xm.SetSuper(name);
                            }
                            else if (bitcal((json["docInfo"]["charShapeList"][currentstyle]["property"]["value"].Value<int>()), 16, 0x1) != 0)
                            {
                                xm.SetSub(name);
                            }

                            //글꼴 적용

                            int fontID = json["docInfo"]["charShapeList"][currentstyle]["faceNameIds"]["array"][0].Value<int>();
                            string fontName = json["docInfo"]["hangulFaceNameList"][fontID]["name"].Value<string>();
                            xm.SetFont(name, fontName);

                            //글자간격
                            double charspace = json["docInfo"]["charShapeList"][currentstyle]["charSpaces"]["array"][0].Value<double>();
                            if (charspace != 0)
                                xm.SetLetterSpace(name, (float)(baseSize * 0.01 * charspace * 0.01 * 0.0353));

                        }

                    }
                }
        }
    }
}
