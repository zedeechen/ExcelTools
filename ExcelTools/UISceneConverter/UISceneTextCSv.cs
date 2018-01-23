using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelTools
{
    class UISceneTextCsv
    {
        public SortedDictionary<int, string> dictText { get; set; }
        public UISceneTextCsv()
        {
            dictText = new SortedDictionary<int, string>();
        }
    }
}
