using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelTools
{
    class TextSheet
    {
        public SortedDictionary<int, string> dictText { get; set; }
        public TextSheet()
        {
            dictText = new SortedDictionary<int, string>();
        }
        public object[,] ToObject()
        {
            object[,] values = new object[dictText.Count, 2];
            int i = 0;
            foreach ( var item in dictText.Keys )
            {
                values[i, 0] = item;
                values[i, 1] = dictText[item];
                i++;
            }
            return values;
        }
    }
}
