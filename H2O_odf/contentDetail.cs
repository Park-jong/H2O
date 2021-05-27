using System;
using System.Collections.Generic;
using System.Text;

namespace H2O_odf
{
    class contentDetail
    {
        string name = "P1";
        string content = "asd";
        float margin_left = 0;
        float margin_right = 0;
        float text_ident = 0;
        string text_color = " fo:color=\"#000000\" loext:opacity=\"100%\"";
        bool bold = false;
        bool italic = false;
        bool line_through = false;
        bool underline = false;
        bool text_super = false;
        bool text_sub = false;

        public string textStyle()
        {
            string style = "<style:style style:name=\"" + name + "\" style:family=\"paragraph\" style:parent-style-name=\"Standard\">";
            if (margin_left != 0 || margin_right != 0 || text_ident != 0)
            {
                style = style + "<style:paragraph-properties"
                  + " fo:margin-left=\"" + margin_left + "cm\""
                  + " fo:margin-right=\"" + margin_right + "cm\""
                  + " fo:text-indent=\"" + text_ident + "cm\" style:auto-text-indent=\"false\"/>";
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

            if (line_through)
            {
                text_line_through = " style:text-line-through-style=\"solid\" style:text-line-through-type=\"single\"";
            }

            if (text_super)
            {
                text_position = " style:text-position=\"super 58%\"";
            }
            else if (text_sub)
            {
                text_position = " style:text-position=\"sub 58%\"";
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
            style = style + "<style:text-properties" + 
                text_color + 
                text_line_through + 
                text_position + 
                font_style + 
                text_underline + 
                font_weight + 
                rsid + 
                font_style_asian + 
                font_weight_asian + 
                font_style_complex + 
                font_weight_complex + "/></style:style>";
            
            return style;
        }

        public string text()
        {
            string str = "<text:p text:style-name=\""+name+"\">"+content+"</text:p>";
            return str;
        }

    }
}
