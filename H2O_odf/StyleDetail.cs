using System;
using System.Collections.Generic;
using System.Text;

namespace H2O_odf
{
    class StyleDetail
    {
        float layout_left = 2;
        float layout_right = 2;
        float layout_top = 2;
        float layout_bottom = 2;

        float page_left = 0;
        float page_right = 0;
        float page_text_indent = 0;
        public string pagelayout()
        {
            string str = "<style:page-layout-properties fo:page-width=\"21.001cm\" fo:page-height=\"29.7cm\" style:num-format=\"1\" style:print-orientation=\"portrait\" " +
                "fo:margin-top=\""+layout_top+ "cm\" fo:margin-bottom=\"" + layout_bottom + "cm\" fo:margin-left=\"" + layout_left + "cm\" fo:margin-right=\"" + layout_right + "cm\" style:writing-mode=\"lr-tb\" " +
                "style:layout-grid-color=\"#c0c0c0\" style:layout-grid-lines=\"20\" style:layout-grid-base-height=\"0.706cm\" style:layout-grid-ruby-height=\"0.353cm\" style:layout-grid-mode=\"none\" style:layout-grid-ruby-below=\"false\" style:layout-grid-print=\"false\" style:layout-grid-display=\"false\" " +
                "style:footnote-max-height=\"0cm\">";
            return str;
        }

        public string pagemargin()
        {
            string str = "<style:paragraph-properties ";
            if(page_left != 0 || page_right != 0 || page_text_indent != 0)
            {
                str = str + "<fo:margin-left=\"" + page_left + "cm\" fo:margin-right=\"" + page_right + "cm\" fo:text-indent=\"" + page_text_indent + "cm\"  style:auto-text-indent=\"false\"";
            }
            str = str + "style:text-autospace=\"none\"/>";
            return str;
        }
    }
}
