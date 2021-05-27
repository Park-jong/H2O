using System;
using System.Collections.Generic;
using System.Text;

namespace H2O_odf
{
    class contentDetail
    {
        protected string name = "P1";
        protected string content = "asdaaaaaaaaaaaaaaaaaaaaaaaa";
        protected float margin_left = 0;
        protected float margin_right = 0;
        protected float margin_top = 0;
        protected float margin_bottom = 0;
        protected float text_ident = 5;
        protected float font_size = 15;
        protected float letter_space = (float)(5 * 0.0353);//1pt -> 0.0353cm 앞이 입력값
        protected string text_color = "#FF0000";
        protected bool bold = false;
        protected bool italic = false;
        protected bool line_through = false;
        protected bool underline = false;
        protected bool text_super = false;
        protected bool text_sub = false;
        protected int align = 1;//일단 0:왼쪽 정렬, 1 오른쪽정렬, 2 가운데정렬, 3 양쪽맞춤이라 가정 나중에 수정
        protected float angle = 0;//글자회전
//실행안됨 추가설정필요        protected bool angle_fix = false;//글자회전 후 폭 맞춤

        public string textStyle()
        {
            string style = "<style:style style:name=\"" + name + "\" style:family=\"paragraph\" style:parent-style-name=\"Standard\">";
            if (margin_left != 0 || margin_right != 0 || margin_top != 0 || margin_bottom != 0 || text_ident != 0 || align != 0)
            {
                switch (align)
                {
                    case 0:
                        {
                            style = style + "<style:paragraph-properties"
                                + " fo:margin-left=\"" + margin_left + "cm\""
                                + " fo:margin-right=\"" + margin_right + "cm\""
                                + " fo:margin-top=\"" + margin_top + "cm\""
                                + " fo:margin-bottom=\"" + margin_bottom + "cm\""
                                + " fo:text-align=\"start\" style:justify-single-word=\"false\""
                                + " fo:text-indent=\"" + text_ident + "cm\" style:auto-text-indent=\"false\"/>";
                            break;
                        }
                    case 1:
                        {
                            style = style + "<style:paragraph-properties"
                                + " fo:margin-left=\"" + margin_left + "cm\""
                                + " fo:margin-right=\"" + margin_right + "cm\""
                                + " fo:margin-top=\"" + margin_top + "cm\""
                                + " fo:margin-bottom=\"" + margin_bottom + "cm\""
                                + " fo:text-align=\"end\" style:justify-single-word=\"false\""
                                + " fo:text-indent=\"" + text_ident + "cm\" style:auto-text-indent=\"false\"/>";
                            break;
                        }
                    case 2:
                        {
                            style = style + "<style:paragraph-properties"
                                + " fo:margin-left=\"" + margin_left + "cm\""
                                + " fo:margin-right=\"" + margin_right + "cm\""
                                + " fo:margin-top=\"" + margin_top + "cm\""
                                + " fo:margin-bottom=\"" + margin_bottom + "cm\""
                                + " fo:text-align=\"center\" style:justify-single-word=\"false\""
                                + " fo:text-indent=\"" + text_ident + "cm\" style:auto-text-indent=\"false\"/>";
                            break;
                        }
                    case 3:
                        {
                            style = style + "<style:paragraph-properties"
                                + " fo:margin-left=\"" + margin_left + "cm\""
                                + " fo:margin-right=\"" + margin_right + "cm\""
                                + " fo:margin-top=\"" + margin_top + "cm\""
                                + " fo:margin-bottom=\"" + margin_bottom + "cm\""
                                + " fo:text-align=\"justify\" style:justify-single-word=\"false\""
                                + " fo:text-indent=\"" + text_ident + "cm\" style:auto-text-indent=\"false\"/>";
                            break;
                        }
                    default:
                        break;
                }
            }


            string rsid = " officeooo:rsid=\"0007e3a5\" officeooo:paragraph-rsid=\"0007e3a5\"";
            string text_line_through = "";
            string text_position = "";
            string font_style = "";
            string text_underline = "";
            string font_weight = "";
            string font_style_asian = "";
            string font_weight_asian = "";
            string font_style_complex = "";
            string font_weight_complex = "";
            string text_color_content = "";
            string font_size_str = "";
            string font_size_asian = "";
            string font_size_complex = "";
            string letter_space_str = "";
            string angle_str = "";

            if(font_size != 12)
            {
                font_size_str = " fo:font-size=\"" + font_size + "pt\"";
                font_size_asian = " fo:font-size-asian=\"" + font_size + "pt\"";
                font_size_complex = " fo:font-size-complex=\"" + font_size + "pt\"";
            }

            if(text_color != "#000000")
            {
                text_color_content = " fo:color=\"" + text_color + "\" loext:opacity=\"100%\"";
            }

            if (line_through)
            {
                text_line_through = " style:text-line-through-style=\"solid\" style:text-line-through-type=\"single\"";
            }

            if (text_super)
            {
                text_position = " style:text-position=\"super 58%\"";//윗 첨자
            }
            else if (text_sub)
            {
                text_position = " style:text-position=\"sub 58%\"";//아랫 첨자
            }

            if (underline)
            {
                text_underline = " style:text-underline-style=\"solid\" style:text-underline-width=\"auto\" style:text-underline-color=\"font-color\"";
            }
            if (bold)
            {
                font_weight = " fo:font-weight=\"bold\"";
                font_weight_asian = " style:font-weight-asian=\"bold\"";
                font_weight_complex = " style:font-weight-complex=\"bold\"";
            }
            if (italic)
            {
                font_style = " fo:font-style=\"italic\"";
                font_style_asian = " style:font-style-asian=\"italic\"";
                font_style_complex = " style:font-style-complex=\"italic\"";
            }
            if(letter_space != 0)
            {
                letter_space_str = " fo:letter-spacing=\"" + letter_space + "cm\"";
            }

            if (angle != 0)
            {
                angle_str = " style:text-rotation-angle=\"" + angle + "\" style:text-rotation-scale=\"line-height\"";
                /*넓이를 알아야 적용가능 변수 많음           
                if (angle_fix)
                {
                    angle_str = " style:text-rotation-angle=\"" + angle + "\" style:text-rotation-scale=\"fixed\" style:text-scale=\"76%\"";
                }
                else
                {
                    angle_str = " style:text-rotation-angle=\"" + angle + "\" style:text-rotation-scale=\"line-height\"";
                }
                */
                angle_str = " style:text-rotation-angle=\"" + angle + "\" style:text-rotation-scale=\"line-height\"";
            }
            style = style + "<style:text-properties" + 
                text_color_content + 
                text_line_through + 
                text_position +
                letter_space_str + 
                font_size_str +
                font_style + 
                text_underline + 
                font_weight + 
                rsid + 
                font_size_asian +
                font_style_asian + 
                font_weight_asian + 
                font_size_complex + 
                font_style_complex + 
                font_weight_complex + 
                angle_str + "/></style:style>";
            
            return style;
        }

        public virtual string text()
        {
            string str = "<text:p text:style-name=\""+name+"\">"+content+"</text:p>";
            return str;
        }
    }
/*
    class pContent : contentDetail
    {
        tContent[] subcontent = new tContent[3];
        public pContent() {
            for (int i = 0; i < 3; i++ ) {
                subcontent[i] = new tContent();
                
            }
            subcontent[0].bold = true;
        }

        public override string text()
        {
            string str = "<text:p text:style-name=\"" + name + "\">" + content;
            for(int i = 0; i < 3; i++)
            {
                str = str + subcontent[i].text();
            }
            str = str + "</text:p>";
            return str;
        }
    }
    class tContent : contentDetail
    {
        public string name = "T1";
        public string content = "poi";

        public override string text()
        {
            string str = "<text:span text:style-name=\"" + name + "\">" + content + "</text:span>";
            return str;
        }
    }*/
}
