using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace ExcelTools.FunctionCSVConverter.ConvererMgr
{
    // 一张excel表生成一个json文件
    class JsonMgr : SingleSheetConverter
    {
        public JsonMgr():
            base()
        {
        }

        protected override void AddBody()
        {
            this.m_builder.Append("{\r\n");

            int writeRows = 0;
            foreach (int row in this.m_sheet.itemPos.Keys)
            {
                ++writeRows;
                string keys = this.m_sheet.cells[row][101].value;
                this.m_builder.Append("\t\"" + keys + "\": {\r\n");

                int writeCol = 0;
                foreach (int col in this.m_sheet.headers.Keys)
                {
                    ++writeCol;
                    string cell = this.m_sheet.cells[row][col].value;

                    // TODO 文字列cell替换为Text表Id，并做记录

                    this.m_builder.Append("\t\t");
                    try
                    {
                        AppendCellForJS(this.m_builder, this.m_sheet.headers[col].titleName, cell);    
                    }
                    catch (JsonReaderException jex)
                    {
                        //Exception in parsing json
                        throw jex;
                    }
                    catch (Exception ex) //some other exception
                    {
                        throw ex;
                    }

                    if (writeCol < this.m_sheet.headers.Count)
                    {
                        this.m_builder.Append(",\r\n");
                    }
                    else
                    {
                        this.m_builder.Append("\r\n");
                    }
                }

                if (writeRows < this.m_sheet.itemPos.Count)
                {
                    this.m_builder.Append("\t},\r\n");
                }
                else
                {
                    this.m_builder.Append("\t}\r\n");
                }

                
            }

            this.m_builder.Append("}\r\n");
        }


    }
}
