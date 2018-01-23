using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelTools
{
    public class FunctionSheet
    {
        public Dictionary<int, TitleConfig> headers { get; set; }
        public SortedDictionary<int, int> itemPos { get; set; }
        public SortedDictionary<int, Dictionary<int, Node>> cells { get; set; }
        public bool bIsAscending;

        public FunctionSheet()
        {
            headers = new Dictionary<int, TitleConfig>();
            itemPos = new SortedDictionary<int, int>();
            cells = new SortedDictionary<int, Dictionary<int, Node>>();
        }

        // 获取行数
        public int GetDataRowCount() 
        {
            return itemPos.Count;
        }
        // 获取列数
        public int GetDataColCount()
        {
            return headers.Count;
        }

        public class TitleConfig
        {
            public int titleIndex;
            public int titlePos;
            public string titleName;
            public string titleChineseName;
            public bool bIsServerOnly;
            public bool bIsTxtCol;
            public TitleConfig( int index, int pos, string name, string nameChs, bool bServerOnly, bool bTxtCol )
            {
                this.titleIndex = index;
                this.titlePos = pos;
                this.titleName = name;
                this.titleChineseName = nameChs;
                this.bIsServerOnly = bServerOnly;
                this.bIsTxtCol = bTxtCol;
            }
            public TitleConfig( int index, int pos, string nameChs, bool bServerOnly )
            {
                this.titleIndex = index;
                this.titlePos = pos;
                this.titleChineseName = nameChs;
                this.bIsServerOnly = bServerOnly;
            }
        }
    }
}
