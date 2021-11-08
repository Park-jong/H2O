using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace WindowsFormsApp1.FuncToXml
{
    public class GridToXml
    {
        public GridToXml()
        {

        }

        public void Run(XmlManager xm, JToken json, JToken docJson, bool zeroCheck)
        {
            //table있는지 체크

            bool hasTable = false;
            int hasTableControlNum;
            int controlListCount = 0;
            try
            {
                controlListCount = json["controlList"].Count();
            }
            catch (System.ArgumentNullException e)
            {
                controlListCount = 0;
            }
            for (int controlList = 0; controlList < controlListCount; controlList++)
            {
                try
                {
                    object existTable = json["controlList"][controlList]["table"].Value<object>();
                    hasTable = true;
                    hasTableControlNum = controlList;
                }
                catch (System.ArgumentNullException e)
                {
                    hasTable = false;
                }
                //table이 존재할때만 실행
                if (hasTable)
                {

                    int rowCount = json["controlList"][controlList]["table"]["rowCount"].Value<int>();
                    int columnCount = json["controlList"][controlList]["table"]["columnCount"].Value<int>();
                    double tableMatginTop = Math.Round(json["controlList"][controlList]["table"]["topInnerMargin"].Value<int>() * 0.01 * 0.0352778, 3);
                    double tableMatginBottom = Math.Round(json["controlList"][controlList]["table"]["bottomInnerMargin"].Value<int>() * 0.01 * 0.0352778, 3);
                    double tableMatginLeft = Math.Round(json["controlList"][controlList]["table"]["leftInnerMargin"].Value<int>() * 0.01 * 0.0352778, 3);
                    double tableMatginRight = Math.Round(json["controlList"][controlList]["table"]["rightInnerMargin"].Value<int>() * 0.01 * 0.0352778, 3); ;

                    string table = xm.MakeTable(rowCount, columnCount);
                    xm.setTable(table, Math.Round(json["controlList"][controlList]["header"]["width"].Value<int>() * 0.01 * 0.0352778, 3));
                    //for(int c = 0; c < columnCount; c++)
                    //{
                    //    int colWidth = json["bodyText"]["sectionList"][s]["paragraphList"][i]["controlList"][controlList]["rowList"][0]["cellList"][c]["listHeader"]["width"].Value<int>();
                    //    xm.setCol(table, c, Math.Round(colWidth * 0.01 * 0.0352778, 3));
                    //}
                    //for (int c = 0; c < rowCount; c++)
                    //{
                    //    int rowHeight = json["bodyText"]["sectionList"][s]["paragraphList"][i]["controlList"][controlList]["rowList"][c]["cellList"][0]["listHeader"]["height"].Value<int>();
                    //    xm.setRow(table, c, Math.Round(rowHeight * 0.01 * 0.0352778, 3));
                    //}
                    for (int rowIndex = 0; rowIndex < json["controlList"][controlList]["rowList"].Count(); rowIndex++)
                    {
                        for (int colIndex = 0; colIndex < json["controlList"][controlList]["rowList"][rowIndex]["cellList"].Count(); colIndex++)
                        {
                            int rowNum = json["controlList"][controlList]["rowList"].Count();
                            int colNum = json["controlList"][controlList]["rowList"][rowIndex]["cellList"].Count();
                            int cellWidth = json["controlList"][controlList]["rowList"][rowIndex]["cellList"][colIndex]["listHeader"]["width"].Value<int>();
                            int cellHeight = json["controlList"][controlList]["rowList"][rowIndex]["cellList"][colIndex]["listHeader"]["height"].Value<int>();
                            int column_index = json["controlList"][controlList]["rowList"][rowIndex]["cellList"][colIndex]["listHeader"]["colIndex"].Value<int>();
                            int row_index = json["controlList"][controlList]["rowList"][rowIndex]["cellList"][colIndex]["listHeader"]["rowIndex"].Value<int>();
                            int margin_top = json["controlList"][controlList]["rowList"][rowIndex]["cellList"][colIndex]["listHeader"]["topMargin"].Value<int>();
                            int margin_bottom = json["controlList"][controlList]["rowList"][rowIndex]["cellList"][colIndex]["listHeader"]["bottomMargin"].Value<int>();
                            int margin_left = json["controlList"][controlList]["rowList"][rowIndex]["cellList"][colIndex]["listHeader"]["leftMargin"].Value<int>();
                            int margin_right = json["controlList"][controlList]["rowList"][rowIndex]["cellList"][colIndex]["listHeader"]["rightMargin"].Value<int>();
                            xm.setCol(table, colIndex, Math.Round(cellWidth * 0.01 * 0.0352778, 3));
                            xm.setRow(table, rowIndex, Math.Round(cellHeight * 0.01 * 0.0352778, 3));
                            xm.SetCell(table, colNum, rowNum, column_index, row_index, Math.Round(cellHeight * 0.01 * 0.0352778, 3), Math.Round(cellWidth * 0.01 * 0.0352778, 3), Math.Round(margin_top * 0.01 * 0.0352778, 3), Math.Round(margin_bottom * 0.01 * 0.0352778, 3), Math.Round(margin_left * 0.01 * 0.0352778, 3), Math.Round(margin_right * 0.01 * 0.0352778, 3));
                        }
                    }
                }
            }
        }
    }
}
