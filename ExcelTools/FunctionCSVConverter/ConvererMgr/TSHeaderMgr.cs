using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelTools.FunctionCSVConverter.ConvererMgr
{
    // 所有excel表生成一个ts文件
    class TSHeaderMgr
    {
        private string exportFilename { get; set; }

        private StringBuilder contentBuilder;
        private Dictionary<string, SheetInfo> sheetInfos;

        class Field
        {
            public string fieldName;
            public string chineseComments;      // 中文注释
        }

        class SheetInfo
        {
            public string filename;             // 文件名
            public string classname;            // 类名
            public int maxFieldLength;          // 最大字段长度
            public List<Field> fieldLists;     // 字段列表
            
            public SheetInfo()
            {
                this.filename = "";
                this.classname = "";
                this.maxFieldLength = 0;
                this.fieldLists = new List<Field>();
            }
        }

        // constructor
        public TSHeaderMgr()
        {
            this.contentBuilder = new StringBuilder();
            this.exportFilename = "config_struct.ts";
            this.sheetInfos = new Dictionary<string, SheetInfo>();
        }

        public void Clear()
        {
            this.contentBuilder = new StringBuilder();
            this.sheetInfos.Clear();
        }

        public string GetExportFileName()
        {
            return this.exportFilename;
        }

        // 返回最终结果
        public string GetResult()
        {
            return this.contentBuilder.ToString();
        }

        // 添加新表
        public void AddNewSheet(string filename, string classname, FunctionSheet sheet)
        {
            SheetInfo info = new SheetInfo();
            info.filename = filename;
            info.classname = classname;

            foreach (int colId in sheet.headers.Keys)
            {
                string propertyName = sheet.headers[colId].titleName;
                string chineseComments = sheet.headers[colId].titleChineseName;
                if (info.maxFieldLength < propertyName.Length)
                {
                    info.maxFieldLength = propertyName.Length;
                }
                Field field = new Field();
                field.fieldName = propertyName;
                field.chineseComments = chineseComments;
                info.fieldLists.Add(field);
            }

            if ( this.sheetInfos.ContainsKey( filename ) ) return;
            else this.sheetInfos.Add(filename, info);
        }

        public void BuildContent()
        {
            this.contentBuilder.Clear();

            this.AddHeader();
            this.AddBody();
            this.AddFooter();
        }

        // 预加头部
        private void AddHeader()
        {
            this.contentBuilder.Append("//********** header **********//\r\n" +
                "import fs = require('fs');\r\n" +
                "import CustomError = require('../../util/errors');\r\n" +
                "import ERRC = require('../../util/error_code');\r\n" +
                "\r\n\r\n");
        }

        private void AddBody()
        {
            this.contentBuilder.Append("//********** body **********//\r\n");
            foreach (KeyValuePair<string, SheetInfo> sheet in this.sheetInfos)
            {
                SheetInfo info = sheet.Value;

                this.AddBodyItem(ref this.contentBuilder, ref info);
            }
        }

        // use by AddBody
        private void AddBodyItem(ref StringBuilder builder, ref SheetInfo info)
        {
            string mgr = info.classname + "Mgr";
            string config = info.classname + "Config";
            int tabCount = (info.maxFieldLength + 5) / 4 + 1;

            builder.Append("// " + info.filename + "\r\n");
            builder.Append("export class " + info.classname + " {\r\n");
            // declare
            foreach (Field field in info.fieldLists)
            {
                builder.Append("\t" + field.fieldName + ":any;");

                // for comments align
                int tempTabCount = (field.fieldName.Length + 5) / 4;
                for (int i = 0; i < tabCount - tempTabCount; i++)
                {
                    builder.Append("\t");
                }
                builder.Append("//" + field.chineseComments + "\r\n");
            }
            // constructor
            builder.Append("\tconstructor(data) {\r\n");
            foreach (Field field in info.fieldLists)
            {
                builder.Append("\t\tthis." + field.fieldName + " = data." + field.fieldName + ";\r\n");
            }
            builder.Append("\t}\r\n");
            builder.Append("}\r\n");

            // 格式不要动
            builder.Append(
"class " + mgr + @" {
    " + config + @" : {[ID:number]: " + info.classname+ @"} = {};
    constructor(data) {
        this." + config + @" = {};
        Object.keys(data).forEach((key) => {
            this." + config + @"[data[key].ID] = new " + info.classname + @"(data[key]);
        });
    }
    public get(ID:number):" + info.classname+ @" {
        var config = this." + config + @"[ID];
        if (!config) {
            throw new CustomError.UserError(ERRC.COMMON.CONFIG_NOT_FOUND, {
                msg: 'COMMON.CONFIG_NOT_FOUND, " + info.filename + @", ID=' + ID
            })
        }
        return config;
    }
    public all():{[ID:number]: " + info.classname+ @"} {
        return this." + config + @";
    }
}" + "\r\n\r\n");
        }

        private void AddFooter()
        {
            this.contentBuilder.Append("//********** footer **********//\r\n");
            this.contentBuilder.Append("export class ConfigMgr {\r\n");
            foreach (KeyValuePair<string, SheetInfo> sheet in this.sheetInfos)
            {
                SheetInfo info = sheet.Value;
                this.contentBuilder.Append("\t" + info.filename + ":" + info.classname + "Mgr = null;\r\n");
            }
            this.contentBuilder.Append("\r\n");
            this.contentBuilder.Append("\tpublic loadAllConfig(jsonDir) {\r\n" +
                                            "\t\tvar contents, json;\r\n\r\n");

            foreach (KeyValuePair<string, SheetInfo> sheet in this.sheetInfos)
            {
                SheetInfo info = sheet.Value;

                this.AddFooterItem(ref this.contentBuilder, ref info);
            }
            this.contentBuilder.Append("\t}\r\n");
            this.contentBuilder.Append("}\r\n");
        }

        // use by AddFooter
        private void AddFooterItem(ref StringBuilder builder, ref SheetInfo info)
        {
            string mgr = info.classname + "Mgr";
            builder.Append("\t\t/// " + info.filename + @"
        try {
            contents = fs.readFileSync(jsonDir + '" + info.filename + @".json');
            json = JSON.parse(contents);
            this." + info.filename + " = new " + mgr + @"(json);
        } catch (err) {
            throw new Error('" + info.filename + @".json read failed');
        }" + "\r\n");
        }
    }

}
