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

        JToken jsonTable;
        JToken jsonRowList;
        JToken jsonHeader;

        int property; //표75 속성
        int rowCount; //표75 RowCount
        int nCols; //표75 nCols
        int cellSpacing; //표75 CellSpacing

        int columnCount;

        private void setData()
        {
            property = jsonTable["property"]["value"].Value<int>();
            rowCount = jsonTable["rowCount"].Value<int>();
            nCols = 0; //값 찾지 못함
            cellSpacing = jsonTable["cellSpacing"].Value<int>();

            columnCount = jsonTable["columnCount"].Value<int>();
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
                    //현재 테이블 JToken Setting
                    jsonTable = json["controlList"][controlList]["table"];
                    jsonRowList = json["controlList"][controlList]["rowList"];
                    jsonHeader = json["controlList"][controlList]["header"];

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
                    setData(); //table 관련 data Setting

                    double tableMatginTop = Math.Round(jsonTable["topInnerMargin"].Value<int>() * 0.01 * 0.0352778, 3);
                    double tableMatginBottom = Math.Round(jsonTable["bottomInnerMargin"].Value<int>() * 0.01 * 0.0352778, 3);
                    double tableMatginLeft = Math.Round(jsonTable["leftInnerMargin"].Value<int>() * 0.01 * 0.0352778, 3);
                    double tableMatginRight = Math.Round(jsonTable["rightInnerMargin"].Value<int>() * 0.01 * 0.0352778, 3); ;

                    string table = xm.MakeTable(rowCount, columnCount);
                    xm.setTable(table, Math.Round(jsonHeader["width"].Value<int>() * 0.01 * 0.0352778, 3));
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
                    for (int rowIndex = 0; rowIndex < jsonRowList.Count(); rowIndex++)
                    {
                        for (int colIndex = 0; colIndex < jsonRowList[rowIndex]["cellList"].Count(); colIndex++)
                        {
                            int rowNum = jsonRowList.Count();
                            int colNum = jsonRowList[rowIndex]["cellList"].Count();
                            int cellWidth = jsonRowList[rowIndex]["cellList"][colIndex]["listHeader"]["width"].Value<int>();
                            int cellHeight = jsonRowList[rowIndex]["cellList"][colIndex]["listHeader"]["height"].Value<int>();
                            int column_index = jsonRowList[rowIndex]["cellList"][colIndex]["listHeader"]["colIndex"].Value<int>();
                            int row_index = jsonRowList[rowIndex]["cellList"][colIndex]["listHeader"]["rowIndex"].Value<int>();
                            int margin_top = jsonRowList[rowIndex]["cellList"][colIndex]["listHeader"]["topMargin"].Value<int>();
                            int margin_bottom = jsonRowList[rowIndex]["cellList"][colIndex]["listHeader"]["bottomMargin"].Value<int>();
                            int margin_left = jsonRowList[rowIndex]["cellList"][colIndex]["listHeader"]["leftMargin"].Value<int>();
                            int margin_right = jsonRowList[rowIndex]["cellList"][colIndex]["listHeader"]["rightMargin"].Value<int>();
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
