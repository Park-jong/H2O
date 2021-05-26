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
            style = style + "<style:text-properties officeooo:rsid=\"0007e3a5\" officeooo:paragraph-rsid=\"000b59b8\"/></style:style>";
            return style;
        }

        public string text()
        {
            string str = "<text:p text:style-name=\""+name+"\">"+content+"</text:p>";
            return str;
        }

    }
}
