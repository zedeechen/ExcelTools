using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace ExcelTools.FunctionCSVConverter.ConvererMgr
{

    class SingleSheetConverter
    {
        protected FunctionSheet m_sheet;
        protected StringBuilder m_builder;
        protected int maxKeyLength;

        public SingleSheetConverter()
        {
            this.maxKeyLength = 0;
            this.m_sheet = null;
            this.m_builder = new StringBuilder();
        }

        public void loadSheet(FunctionSheet sheet)
        {
            this.m_sheet = sheet;
            foreach (int col in this.m_sheet.headers.Keys)
            {
                if (this.maxKeyLength < this.m_sheet.headers[col].titleName.Length)
                {
                    this.maxKeyLength = this.m_sheet.headers[col].titleName.Length;
                }
            }
        }

        public void BuildContent()
        {
            this.m_builder.Clear();
            this.AddHeader();
            this.AddBody();
            this.AddFooter();
        }

        public string GetResult()
        {
            return this.m_builder.ToString();
        }

        protected virtual void AddHeader()
        {
            // do nothing
        }

        protected virtual void AddBody()
        {
            // do nothing
        }

        protected virtual void AddFooter()
        {
            // do nothing
        }

        protected void AppendCellForJS(
                        StringBuilder builder,
                        string key,
                        String value)
        {
            StringBuilder propertyContent = new StringBuilder();
            bool isJSON = Regex.IsMatch(key, "^JSON_.*$");
            if (!value.Equals(""))
            {
                if (isJSON)
                {
                    List<string> errorMsgs = new List<string>();
                    try
                    {
                        var obj = JToken.Parse(value);
                    }
                    catch (JsonReaderException jex)
                    {
                        //Exception in parsing json
                        Console.WriteLine(jex.Message);
                        throw jex;
                    }
                    catch (Exception ex) //some other exception
                    {
                        Console.WriteLine(ex.ToString());
                        throw ex;
                    }
                    propertyContent.Append(value);
                }
                else
                {
                    bool bWithQuota = ((!Regex.IsMatch(value, "^-?[0-9]+$") && !Regex.IsMatch(value, "^(-?\\d+)(\\.\\d+)?$")) || value.IndexOf('\"') != -1);

                    if (bWithQuota)
                    {
                        propertyContent.Append("\"");
                    }

                    foreach (char c in value)
                    {
                        if (c == '\"')
                        {
                            propertyContent.Append("\\");
                        }
                        propertyContent.Append(c);
                    }

                    if (bWithQuota)
                    {
                        propertyContent.Append("\"");
                    }
                }
            }
            else
            {
                propertyContent.Append("0");
            }

            int tabCount = (this.maxKeyLength + 2) / 4 + 1;
            int tempTabCount = (key.Length + 2) / 4;

            builder.Append("\"" + key + "\"");
            for (int i = 0; i < tabCount - tempTabCount; i++)
            {
                builder.Append("\t");
            }
            builder.Append(":\t" + propertyContent.ToString());
        }

    }
}
