using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelTools
{
    class CustomSheet
    {
        public Dictionary<string, Dictionary<int, Item>> dictCustomItem { get; set; }
        public CustomSheet()
        {
            dictCustomItem = new Dictionary<string, Dictionary<int, Item>>();
        }

        public class Item
        {
            public int status; // 1表示修改，2表示新增
            public Dictionary<int, string> dictOpCell;
            public Item( int sta )
            {
                this.status = sta;
                dictOpCell = new Dictionary<int,string>();
            }
        }
    }
}
